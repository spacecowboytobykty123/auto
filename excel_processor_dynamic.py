import pandas as pd
import time
import psycopg2
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import requests
import helpers


def get_id_with_leading_zeros(value):
    if pd.isna(value):
        return None
    str_value = str(value).strip()
    if 'E+' in str_value.upper() or 'e+' in str_value:
        try:
            float_val = float(str_value)
            int_val = int(float_val)
            str_value = str(int_val)
        except:
            pass
    if '.' in str_value:
        try:
            float_val = float(str_value)
            if float_val.is_integer():
                str_value = str(int(float_val))
        except:
            pass
    return str_value


def fetch_html_with_timeout(url, timeout=5):
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Ошибка загрузки HTML: {e}")
        return None


def analyze_application_from_html(html_content):
    try:
        soup = BeautifulSoup(html_content, "html.parser")
        conclusion = helpers.analyzeFullApplication(soup)
        return conclusion
    except Exception as e:
        return f"Ошибка анализа HTML: {e}"


def query_postgres_by_app_id(app_id, conn):
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT i.id, i.requestnumber, s.name AS status_name, h.created_at
                FROM history.application_info i
                LEFT JOIN history.status_history h ON h.applicationinfo_id = i.id
                LEFT JOIN history.status_go s ON s.id = h.statusgo_id
                WHERE i.requestnumber = %s
                ORDER BY h.created_at ASC;            """, (app_id,))
            result = cur.fetchone()
            if result:
                status = result[2] or "Статус отсутствует"
                date = result[3].strftime('%Y-%m-%d %H:%M') if result[3] else "Дата отсутствует"
                return f"Статус: {status}, Дата: {date}"
            else:
                return "Нет истории по номеру заявления"
    except Exception as e:
        return f"Ошибка запроса: {e}"


def preserve_excel_formatting(original_file, output_file, df_updated, sheet_index):
    try:
        wb = load_workbook(original_file)
        ws = wb.worksheets[sheet_index]
        comment_col_index = None
        for col_idx, cell in enumerate(ws[1], 1):
            if cell.value and 'Комментарий АО НИТ' in str(cell.value):
                comment_col_index = col_idx
                break
        if comment_col_index is None:
            return False
        for row_idx in range(2, len(df_updated) + 2):
            df_row_idx = row_idx - 2
            comment_value = df_updated.iloc[df_row_idx]['Комментарий АО НИТ']
            if pd.notna(comment_value) and comment_value != '':
                ws.cell(row=row_idx, column=comment_col_index).value = comment_value
        wb.save(output_file)
        return True
    except Exception as e:
        print(f"Ошибка сохранения Excel: {e}")
        return False


def process_html_sheet(file_path, output_path):
    try:
        df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl', dtype=str)
    except Exception as e:
        print(f"Ошибка чтения первой страницы Excel: {e}")
        return

    identifier_col = comment_col = None
    for col in df.columns:
        if 'Идентификатор' in str(col):
            identifier_col = col
        elif 'Комментарий АО НИТ' in str(col):
            comment_col = col

    if identifier_col is None:
        print("Столбец идентификатора не найден на первой странице")
        return

    if comment_col is None:
        comment_col = 'Комментарий АО НИТ'
        df[comment_col] = ''

    print(f"Найден столбец идентификатора: {identifier_col}")
    print(f"Найден столбец комментариев: {comment_col}")
    print(f"Всего строк для обработки: {len(df)}")

    base_url = "http://192.168.130.100/csp/iiscon/isc.util.About.cls?Action=4&appId="
    successful_count = 0
    failed_count = 0

    for index, row in df.iterrows():
        app_id_raw = row[identifier_col]
        if pd.isna(app_id_raw) or str(app_id_raw).strip() in ('', 'nan'):
            continue

        app_id = get_id_with_leading_zeros(app_id_raw)
        if app_id is None:
            continue

        url = base_url + app_id
        print(f"\n[{index + 1}/{len(df)}] Обработка заявки ID: {app_id}")

        html_content = fetch_html_with_timeout(url, timeout=5)
        if html_content:
            conclusion = analyze_application_from_html(html_content)
            df.at[index, comment_col] = conclusion
            successful_count += 1
            print(f"✓ Успешно: {conclusion}")
        else:
            df.at[index, comment_col] = "Ошибка: не удалось получить данные"
            failed_count += 1
            print("✗ Не удалось получить данные")

        time.sleep(0.5)

    preserved = preserve_excel_formatting(file_path, output_path, df, sheet_index=0)
    if not preserved:
        print("⚠️ Форматирование не сохранено, сохраняем обычным способом")
        df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\n=== РЕЗУЛЬТАТЫ ===")
    print(f"Успешно обработано: {successful_count}")
    print(f"Ошибок: {failed_count}")
    print(f"Результаты сохранены в: {output_path}")


def process_pep_sheet_with_full_analysis(file_path, output_path, conn):
    try:
        df = pd.read_excel(file_path, sheet_name=2, engine='openpyxl', dtype=str)
    except Exception as e:
        print(f"Ошибка чтения листа 'ПЭП': {e}")
        return

    identifier_col = comment_col = pep_col = None
    for col in df.columns:
        if 'Идентификатор заявки' in str(col):
            identifier_col = col
        elif 'Комментарий АО НИТ' in str(col):
            comment_col = col
        elif 'Номер заявления' in str(col):
            pep_col = col

    if identifier_col is None or pep_col is None:
        print("Не найдены нужные столбцы на листе 'ПЭП'")
        return

    if comment_col is None:
        comment_col = 'Комментарий АО НИТ'
        df[comment_col] = ''

    print(f"Найден столбец идентификатора: {identifier_col}")
    print(f"Найден столбец PEP: {pep_col}")
    print(f"Найден столбец комментариев: {comment_col}")
    print(f"Всего строк для обработки: {len(df)}")

    successful_count = 0
    failed_count = 0

    for index, row in df.iterrows():
        app_id = row[identifier_col]
        pep_id = row[pep_col]
        if pd.isna(app_id) or pd.isna(pep_id):
            continue

        print(f"\n[{index + 1}/{len(df)}] Обработка PEP-записи: ID={app_id}, Номер={pep_id}")
        html_content = fetch_html_with_timeout(f"http://192.168.130.100/csp/iiscon/isc.util.About.cls?Action=4&appId={app_id}")
        html_conclusion = analyze_application_from_html(html_content) if html_content else "Ошибка загрузки HTML"
        db_conclusion = query_postgres_by_app_id(pep_id, conn)
        df.at[index, comment_col] = f"HTML: {html_conclusion} | БД: {db_conclusion}"
        successful_count += 1
        time.sleep(0.2)

    preserved = preserve_excel_formatting(file_path, output_path, df, sheet_index=2)
    if not preserved:
        print("⚠️ Форматирование не сохранено, сохраняем обычным способом")
        df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\n=== РЕЗУЛЬТАТЫ (ПЭП) ===")
    print(f"Успешно обработано: {successful_count}")
    print(f"Ошибок: {len(df) - successful_count}")
    print(f"Результаты сохранены в: {output_path}")


def process_combined_excel_pipeline(file_path, output_path):
    process_html_sheet(file_path, output_path)
    try:
        conn = psycopg2.connect(
            host="192.168.175.27",
            port=5432,
            dbname="egov",
            user="alisher_ibrayev",
            password="ASTkazkorp2010!@#",
            sslmode="require"  # или "disable", если точно без SSL
        )
        process_pep_sheet_with_full_analysis(file_path, output_path, conn)
        conn.close()
    except Exception as e:
        print(f"Ошибка подключения к базе данных: {e}")
