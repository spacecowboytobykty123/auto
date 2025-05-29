from datetime import datetime


def checkChanges(historyTable):
    """Legacy function - prints all changes data."""
    changesTable = historyTable[4]
    changesTrs = changesTable.find_all("tr")

    for tr in changesTrs:
        changesTds = tr.find_all("td")
        chData = [td.get_text(strip=True) for td in changesTds]
        print(chData)


def isMadeItInTime(lastMod, deadline):
    """Check if the task was completed before the deadline."""
    last_changed = datetime.strptime(lastMod, "%Y-%m-%d %H:%M:%S.%f")
    deadline = datetime.strptime(deadline, "%Y-%m-%d %H:%M:%S.%f")

    if last_changed < deadline:
        return True
    else:
        return False


def checkErrors(historyTable):
    """Legacy function - checks for errors in the history table."""
    needTabled = historyTable[4]
    trs = needTabled.find_all("tr")

    for tr in trs:
        tds = tr.find_all("td")
        data = [td.get_text(strip=True) for td in tds]
        if len(data) > 2 and data[2] == '2':
            print("Error!")
        else:
            print("No error!")


def findErrorsTable(soup):
    """
    Find the errors table that comes after the <br><b>Очередь уведомлений isc.kzcon.ens.MsgQueue</b> element.
    """
    try:
        # Look for the specific text pattern
        target_elements = soup.find_all(['b', 'strong'])
        errors_table = None

        for element in target_elements:
            if element and 'Очередь уведомлений isc.kzcon.ens.MsgQueue' in element.get_text():
                # Found the target element, now find the next table
                next_sibling = element.parent
                while next_sibling:
                    next_sibling = next_sibling.find_next_sibling()
                    if next_sibling and next_sibling.name == 'table':
                        errors_table = next_sibling
                        break
                break

        # Alternative approach: look for table after br tag containing the text
        if not errors_table:
            br_elements = soup.find_all('br')
            for br in br_elements:
                next_element = br.next_sibling
                if next_element and 'Очередь уведомлений isc.kzcon.ens.MsgQueue' in str(next_element):
                    # Find the next table after this br
                    next_table = br.find_next('table')
                    if next_table:
                        errors_table = next_table
                        break

        return errors_table
    except Exception as e:
        print(f"Error finding errors table: {e}")
        return None

def load_error_mapping(filepath):
    error_mapping = {}
    with open(filepath, 'r', encoding='utf-8') as file:
        for line in file:
            if ' - ' in line:
                error_text, owner = line.strip().split(' - ', 1)
                error_mapping[error_text.strip()] = owner.strip()
                print(error_text)
                print(owner)
                print(error_mapping)
    return error_mapping


def find_error_owner(error_message, error_mapping):
    for error_text, owner in error_mapping.items():
        if error_text in error_message:
            return owner
    return "Неизвестно"


def checkTechnicalErrors(soup):
    """
    Check for technical errors in the notification queue table.
    Returns a string with error information or empty string if no errors.
    """
    try:
        errors_table = findErrorsTable(soup)
        if not errors_table:
            return ""

        # Get all rows from the errors table
        rows = errors_table.find_all('tr')
        if len(rows) < 2:  # Need at least header + 1 data row
            return ""

        # Find column indices by examining the header row
        header_row = rows[0]
        headers = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]

        queue_type_index = -1
        last_error_index = -1

        for i, header in enumerate(headers):
            if 'QueueType' in header:
                queue_type_index = i
            elif 'LastError' in header:
                last_error_index = i

        if queue_type_index == -1 or last_error_index == -1:
            return ""

        # Check data rows for errors
        error_messages = []

        for row in rows[1:]:  # Skip header row
            cells = row.find_all(['td', 'th'])
            if len(cells) > max(queue_type_index, last_error_index):
                queue_type = cells[queue_type_index].get_text(strip=True)
                last_error = cells[last_error_index].get_text(strip=True)

                # Check if QueueType is 2 or 4 (error states)
                if queue_type in ['2', '4'] and last_error:
                    error_messages.append(last_error)

        # Format error messages
        # тут текст ошибки
        if error_messages:
            result = "Однако ошибка - " + error_messages[0] + " . "
            for additional_error in error_messages[1:]:
                result += "А также ошибка - " + additional_error + " . "
            return result

        return ""
    except Exception as e:
        return f"Ошибка при проверке технических ошибок: {e}"


def analyzeStatusSequence(historyTable):
    """
    Analyze the sequence of statuses and return appropriate conclusion text.
    """
    changesTable = historyTable[4]
    changesTrs = changesTable.find_all("tr")

    # Extract all statuses in chronological order (skip header row)
    statuses = []
    for tr in changesTrs[1:]:  # Skip header row
        changesTds = tr.find_all("td")
        if len(changesTds) > 6:  # Make sure we have enough columns
            status = changesTds[6].get_text(strip=True)  # newStatus column
            if status:  # Only add non-empty statuses
                statuses.append(status)

    if not statuses:
        return "Нет данных о статусах"

    # Get the last status
    last_status = statuses[-1]

    # Count occurrences of each status type
    accepted_count = statuses.count('ACCEPTED')
    launched_count = statuses.count('LAUNCHED')

    # Check for priority statuses (STARTED, FINISHED, READY, HANDED, CANCELED)
    # These override all previous logic
    priority_statuses = ['STARTED', 'FINISHED', 'READY', 'HANDED', 'CANCELED']

    # If any priority status is present, use that logic
    for status in reversed(statuses):  # Check from last to first
        if status in priority_statuses:
            if status == 'STARTED':
                return "ГУ на исполнении. Рассмотреть на стороне ГО."
            elif status in ['FINISHED', 'READY', 'HANDED']:
                return "ГУ оказана несвоевременно. Рассмотреть на стороне ГО."
            elif status == 'CANCELED':
                return "ГУ отменена."
            break

    # Logic for ACCEPTED and LAUNCHED statuses only
    # (only applies if no priority statuses were found)

    # Case: Only one ACCEPTED status
    if accepted_count == 1 and launched_count == 0:
        return "ГУ принята от заявителя. Рассмотреть на стороне ГО."

    # Case: Two consecutive ACCEPTED statuses with no LAUNCHED
    elif accepted_count == 2 and launched_count == 0:
        return "Оператор цон не провел через накопитель Б"

    # Case: Two ACCEPTED and one LAUNCHED afterwards
    elif accepted_count == 2 and launched_count == 1:
        return "ГУ на исполнении. Рассмотреть на стороне ГО"

    # Case: One ACCEPTED and one LAUNCHED
    elif accepted_count == 1 and launched_count == 1:
        return "Рассмотреть на SHEP"

    # Default case for other combinations
    else:
        return f"Неопределенная последовательность статусов: {' -> '.join(statuses)}"


def analyzeFullApplication(soup):
    """
    Complete analysis including both status sequence and technical errors.
    Returns the final conclusion with error information if present.
    """
    try:
        historyTable = soup.find_all("table")

        if len(historyTable) < 5:
            return "Ошибка: недостаточно таблиц в HTML"

        # Get the basic conclusion from status analysis
        basic_conclusion = analyzeStatusSequence(historyTable)

        # Check for technical errors
        error_info = checkTechnicalErrors(soup)

        # Combine conclusion with error information
        final_conclusion = basic_conclusion + error_info

        return final_conclusion

    except Exception as e:
        return f"Ошибка анализа: {str(e)}"


def printStatusHistory(historyTable):
    """
    Print the status history for debugging purposes.
    """
    changesTable = historyTable[4]
    changesTrs = changesTable.find_all("tr")

    print("=== История статусов ===")
    for i, tr in enumerate(changesTrs):
        changesTds = tr.find_all("td")
        if len(changesTds) > 6:
            date = changesTds[2].get_text(strip=True) if len(changesTds) > 2 else "N/A"
            old_status = changesTds[7].get_text(strip=True) if len(changesTds) > 7 else "N/A"
            new_status = changesTds[6].get_text(strip=True) if len(changesTds) > 6 else "N/A"

            if i == 0:  # Header row
                print(f"{'#':<3} {'Дата':<25} {'Старый статус':<15} {'Новый статус':<15}")
                print("-" * 65)
            else:
                print(f"{i:<3} {date:<25} {old_status:<15} {new_status:<15}")
    print("=" * 65)