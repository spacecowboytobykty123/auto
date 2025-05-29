import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime
import helpers


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


def process_excel_with_dynamic_fetch(excel_file_path, output_file_path=None):
    """
    Process Excel file by fetching HTML for each application ID dynamically.
    """
    # Read Excel file
    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
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
        app_id = row[identifier_col]

        # Skip if ID is empty or NaN
        if pd.isna(app_id) or str(app_id).strip() == '':
            continue

        # Convert to string and clean up
        app_id = "00" + str(app_id).strip()
        url = base_url + app_id

        print(f"\n[{index + 1}/{len(df)}] Обработка заявки ID: {app_id}")

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

    # Save results
    if output_file_path is None:
        output_file_path = excel_file_path.replace('.xlsx', '_processed.xlsx')

    try:
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        print(f"\n=== РЕЗУЛЬТАТЫ ===")
        print(f"Успешно обработано: {successful_count}")
        print(f"Ошибок: {failed_count}")
        print(f"Результаты сохранены в: {output_file_path}")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


def main():
    excel_file = "Павлодарская область_Апрель_75.xlsx"
    output_file = "Павлодарская область_Апрель_75_processed.xlsx"

    print("=== ОБРАБОТКА EXCEL С ДИНАМИЧЕСКИМ ПОЛУЧЕНИЕМ HTML ===")
    process_excel_with_dynamic_fetch(excel_file, output_file)


if __name__ == "__main__":
    main()