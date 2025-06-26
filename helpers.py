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


def getDeadlineFromMainTable(soup):
    """
    Find the deadline from the main properties table with "Основные свойства заявки" header.
    """
    try:
        # Look for the text "Основные свойства заявки"
        elements = soup.find_all('b')

        for element in elements:
            if element and 'Основные свойства заявки' in element.get_text():
                # Found the target element, now find the next table
                table = element.find_next('table')
                if table:
                    # Look for deadline column in the header row
                    rows = table.find_all('tr')
                    if len(rows) >= 2:  # Need header + data row
                        header_row = rows[0]
                        data_row = rows[1]

                        header_cells = header_row.find_all(['td', 'th'])
                        data_cells = data_row.find_all(['td', 'th'])

                        # Find deadline column index
                        deadline_index = -1
                        for i, cell in enumerate(header_cells):
                            if 'deadline' in cell.get_text(strip=True).lower():
                                deadline_index = i
                                break

                        # Get deadline value
                        if deadline_index != -1 and deadline_index < len(data_cells):
                            deadline_value = data_cells[deadline_index].get_text(strip=True)
                            if deadline_value:
                                return deadline_value
                break

        return None
    except Exception as e:
        print(f"Error finding deadline: {e}")
        return None


def checkStatusDeadline(status_create_date, deadline):
    """
    Check if status was created before deadline.
    """
    try:
        if not status_create_date or not deadline:
            return False

        # Skip non-date entries like "currentState"
        if not status_create_date.replace('-', '').replace(':', '').replace('.', '').replace(' ', '').isdigit():
            return False

        # Parse both dates
        status_date = datetime.strptime(status_create_date, "%Y-%m-%d %H:%M:%S.%f")
        deadline_date = datetime.strptime(deadline, "%Y-%m-%d %H:%M:%S.%f")

        return status_date <= deadline_date
    except Exception as e:
        print(f"Error checking status deadline: {e}")
        return False


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
        if error_messages:
            result = ". Однако ошибка - " + error_messages[0] + " ."
            for additional_error in error_messages[1:]:
                result += "А также ошибка - " + additional_error + " ."
            return result

        return ""
    except Exception as e:
        return f"Ошибка при проверке технических ошибок: {e}"



def analyzeStatusSequence(historyTable, soup):
    """
    Analyze the sequence of statuses and return appropriate conclusion text.
    """
    changesTable = historyTable[4]
    changesTrs = changesTable.find_all("tr")

    # Extract all statuses with their create dates in chronological order (skip header row)
    statuses = []
    status_dates = []

    for tr in changesTrs[1:]:  # Skip header row
        changesTds = tr.find_all("td")
        if len(changesTds) > 6:  # Make sure we have enough columns
            status = changesTds[6].get_text(strip=True)  # newStatus column
            create_date = changesTds[2].get_text(strip=True) if len(changesTds) > 2 else ""  # createDate column
            if status:  # Only add non-empty statuses
                statuses.append(status)
                status_dates.append(create_date)

    if not statuses:
        return "Нет данных о статусах"

    # Count occurrences of each status type
    accepted_count = statuses.count('ACCEPTED')
    launched_count = statuses.count('LAUNCHED')

    # Get deadline from main table
    deadline = getDeadlineFromMainTable(soup)

    # Check for priority statuses (STARTED, FINISHED, READY, HANDED, CANCELED)
    # These override all previous logic
    priority_statuses = ['STARTED', 'FINISHED', 'READY', 'HANDED', 'CANCELED']

    # If any priority status is present, use that logic
    for status in reversed(statuses):  # Check from last to first
        if status in priority_statuses:
            if status == 'STARTED':
                return "ГУ на исполнении", deadline
            elif status in ['FINISHED', 'READY', 'HANDED']:
                # Find the FIRST occurrence of any final status for deadline checking
                first_final_index = -1
                for i, s in enumerate(statuses):
                    if s in ['FINISHED', 'READY', 'HANDED']:
                        first_final_index = i
                        break

                if first_final_index != -1:
                    status_create_date = status_dates[first_final_index] if first_final_index < len(
                        status_dates) else ""

                    if deadline and status_create_date:
                        is_on_time = checkStatusDeadline(status_create_date, deadline)
                        if is_on_time:
                            return "ГУ оказана своевременно", deadline
                        else:
                            return "ГУ оказана несвоевременно", deadline
                    else:
                        return "ГУ завершена (не удалось проверить сроки)", deadline
                else:
                    return "ГУ завершена", deadline
            elif status == 'CANCELED':
                return "ГУ отменена", deadline
            break

    # Logic for ACCEPTED and LAUNCHED statuses only
    # (only applies if no priority statuses were found)

    # NEW: Check for ACCEPTED -> LAUNCHED -> ACCEPTED pattern
    if len(statuses) >= 3:
        last_three = statuses[-3:]
        if last_three == ['ACCEPTED', 'LAUNCHED', 'ACCEPTED']:
            return "ГУ не доставлена до исполнителя", deadline

    # Case: Only one ACCEPTED status
    if accepted_count == 1 and launched_count == 0:
        return "ГУ принята от заявителя", deadline

    # Case: Two consecutive ACCEPTED statuses with no LAUNCHED
    elif accepted_count == 2 and launched_count == 0:
        return "Оператор цон не провел через накопитель Б", deadline

    # Case: Two ACCEPTED and one LAUNCHED afterwards
    elif accepted_count == 2 and launched_count == 1:
        return "ГУ на исполнении", deadline

    # Case: One ACCEPTED and one LAUNCHED
    elif accepted_count == 1 and launched_count == 1:
        return "Рассмотреть на SHEP", deadline

    # Default case for other combinations
    else:
        return f"Неопределенная последовательность статусов: {' -> '.join(statuses)}", deadline


def analyzeFullApplication(soup, return_deadline=False):
    """
    Complete analysis including both status sequence and technical errors.
    Returns the final conclusion with error information if present.
    """
    try:
        historyTable = soup.find_all("table")

        if len(historyTable) < 5:
            return ("Ошибка: недостаточно таблиц в HTML", None) if return_deadline else "Ошибка: недостаточно таблиц в HTML"

        # Get the basic conclusion from status analysis
        basic_conclusion, deadline = analyzeStatusSequence(historyTable, soup)

        # Check for technical errors
        error_info = checkTechnicalErrors(soup)

        # Determine if we should add "Рассмотреть на стороне ГО."
        # Only add if there are NO technical errors AND conclusion needs it
        should_add_go_review = False

        if not error_info:  # No technical errors
            conclusions_needing_go_review = [
                "ГУ принята от заявителя",
                "ГУ на исполнении",
                "ГУ оказана несвоевременно"
            ]

            if basic_conclusion in conclusions_needing_go_review:
                should_add_go_review = True

        # Build final conclusion
        final_conclusion = basic_conclusion

        if should_add_go_review:
            final_conclusion += ". Рассмотреть на стороне ГО."

        # Add error information if present
        final_conclusion += error_info

        return (final_conclusion, deadline) if return_deadline else final_conclusion

    except Exception as e:
        result = f"Ошибка анализа: {str(e)}"
        return (result, None) if return_deadline else result


def analyzeDBStatuses(lastStatus):
    # lastStatus = statuses[-1]
    if lastStatus == "IN_PROCESSING":
        return "ГУ на исполнении. Рассмотреть на стороне ГО."
    elif lastStatus == "SENT" or "ACCEPTED":
        return "Рассмотреть на SHEP"
    elif lastStatus in ["COMPLETED", "CANCELLED", "APPROVED"]:
        return "FINISHED"
    elif lastStatus == "CREATED":
        return "REGISTERED"
    else:
        return "Не нашли соответствия"




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


def hasTechErrors(statusList):
    for status in statusList:
        if status[2] == "TECH_ERROR":
            return True
    return False