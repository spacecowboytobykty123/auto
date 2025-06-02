import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime
import helpers
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import copy


def fetch_html_with_timeout(url, timeout=5):
    """
    Fetch HTML from URL with timeout. Returns HTML content or None if failed.
    """
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        return response.text
    except requests.exceptions.Timeout:
        print(f"Timeout ({timeout}s) exceeded for URL: {url}")
        return None
    except requests.exceptions.RequestException as e:
        print(f"Error fetching {url}: {e}")
        return None


def analyze_application_from_html(html_content):
    """
    Analyze HTML content and return conclusion based on status sequence and technical errors.
    """
    try:
        soup = BeautifulSoup(html_content, "html.parser")

        # Use the comprehensive analysis that includes both status and error checking
        conclusion = helpers.analyzeFullApplication(soup)
        return conclusion

    except Exception as e:
        return f"Ошибка анализа: {str(e)}"


def preserve_excel_formatting(original_file, output_file, df_updated):
    """
    Preserve original Excel formatting while updating the data.
    """
    try:
        # Load the original workbook with formatting
        wb = load_workbook(original_file)
        # Always work with the first sheet (index 0)
        ws = wb.worksheets[0]

        # Find the comment column index in the worksheet
        comment_col_index = None
        for col_idx, cell in enumerate(ws[1], 1):  # Check first row (headers)
            if cell.value and 'Комментарий АО НИТ' in str(cell.value):
                comment_col_index = col_idx
                break

        if comment_col_index is None:
            print("Warning: Could not find 'Комментарий АО НИТ' column in original file")
            return False

        # Update only the comment column data, preserving all formatting
        for row_idx in range(2, len(df_updated) + 2):  # Start from row 2 (skip header)
            df_row_idx = row_idx - 2  # Convert to DataFrame index
            if df_row_idx < len(df_updated):
                comment_value = df_updated.iloc[df_row_idx]['Комментарий АО НИТ']
                if pd.notna(comment_value) and comment_value != '':
                    ws.cell(row=row_idx, column=comment_col_index).value = comment_value

        # Save the workbook with preserved formatting
        wb.save(output_file)
        return True

    except Exception as e:
        print(f"Error preserving formatting: {e}")
        return False


def get_id_with_leading_zeros(value):
    """
    Extract ID ensuring leading zeros are preserved, using the "00" + str approach.
    """
    if pd.isna(value):
        return None

    # Convert to string and strip whitespace
    str_value = str(value).strip()

    # If it's in scientific notation (like 2.27E+09), convert back to integer string
    if 'E+' in str_value.upper() or 'e+' in str_value:
        try:
            # Convert scientific notation to integer, then to string
            float_val = float(str_value)
            int_val = int(float_val)
            str_value = str(int_val)
        except:
            pass

    # Remove decimal point if present (like "123456.0" -> "123456")
    if '.' in str_value:
        try:
            float_val = float(str_value)
            if float_val.is_integer():
                str_value = str(int(float_val))
        except:
            pass



    return str_value


def process_excel_with_dynamic_fetch(excel_file_path, output_file_path=None):
    """
    Process Excel file by fetching HTML for each application ID dynamically.
    Preserves original formatting and handles leading zeros correctly.
    """
    # Read Excel file with string dtype to preserve leading zeros
    try:
        # First, try to read with original formatting preserved
        df = pd.read_excel(excel_file_path, engine='openpyxl', dtype=str)

        # Also read without dtype=str to get proper column detection
        df_detect = pd.read_excel(excel_file_path, engine='openpyxl')

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Find the required columns
    identifier_col = None
    comment_col = None

    for col in df.columns:
        if 'Идентификатор заявки' in str(col) or 'Идентификатор' in str(col):
            identifier_col = col
        elif 'Комментарий АО НИТ' in str(col):
            comment_col = col

    if identifier_col is None:
        print("Ошибка: не найден столбец с идентификатором заявки")
        print("Доступные столбцы:", df.columns.tolist())
        return

    if comment_col is None:
        print("Предупреждение: не найден столбец 'Комментарий АО НИТ', создаем новый")
        comment_col = 'Комментарий АО НИТ'
        df[comment_col] = ''

    print(f"Найден столбец идентификатора: {identifier_col}")
    print(f"Найден столбец комментариев: {comment_col}")
    print(f"Всего строк для обработки: {len(df)}")

    # Base URL template
    base_url = "http://192.168.130.100/csp/iiscon/isc.util.About.cls?Action=4&appId="

    # Process each row
    successful_count = 0
    failed_count = 0

    for index, row in df.iterrows():
        # Get the application ID with proper leading zeros
        app_id_raw = row[identifier_col]

        # Skip if ID is empty or NaN
        if pd.isna(app_id_raw) or str(app_id_raw).strip() == '' or str(app_id_raw).strip() == 'nan':
            continue

        # Convert to string, clean up, and add "00" prefix
        app_id = get_id_with_leading_zeros(app_id_raw)
        if app_id is None:
            continue

        url = base_url + app_id

        print(f"\\n[{index + 1}/{len(df)}] Обработка заявки ID: {app_id}")

        # Fetch HTML with timeout
        html_content = fetch_html_with_timeout(url, timeout=5)

        if html_content:
            # Analyze the HTML and get conclusion
            conclusion = analyze_application_from_html(html_content)
            df.at[index, comment_col] = conclusion
            successful_count += 1
            print(f"✓ Успешно: {conclusion}")
        else:
            df.at[index, comment_col] = "Ошибка: не удалось получить данные"
            failed_count += 1
            print("✗ Не удалось получить данные")

        # Small delay to avoid overwhelming the server
        time.sleep(0.5)

    # Save results with preserved formatting
    if output_file_path is None:
        output_file_path = excel_file_path.replace('.xlsx', '_processed.xlsx')

    try:
        # Try to preserve original formatting
        formatting_preserved = preserve_excel_formatting(excel_file_path, output_file_path, df)

        if not formatting_preserved:
            print("Warning: Could not preserve original formatting, saving as new file...")
            # Fallback: save without formatting preservation
            df.to_excel(output_file_path, index=False, engine='openpyxl')

        print(f"\\n=== РЕЗУЛЬТАТЫ ===")
        print(f"Успешно обработано: {successful_count}")
        print(f"Ошибок: {failed_count}")
        print(f"Результаты сохранены в: {output_file_path}")

        if formatting_preserved:
            print("✓ Оригинальное форматирование сохранено")
        else:
            print("⚠ Форматирование не сохранено")

    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


def main():
    excel_file = "Павлодарская область_Апрель_75.xlsx"
    output_file = "Павлодарская область_Апрель_75_processed.xlsx"

    print("=== ОБРАБОТКА EXCEL С ДИНАМИЧЕСКИМ ПОЛУЧЕНИЕМ HTML ===")
    process_excel_with_dynamic_fetch(excel_file, output_file)


if __name__ == "__main__":
    main()