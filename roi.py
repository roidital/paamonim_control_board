from time import sleep
from typing import Final
import datetime

from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QMessageBox, QLineEdit
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from login.login import do_login
from openpyxl import Workbook


URL_PAGE: Final[str] = 'https://app.paamonim.org.il/budgets'


def filter_unit_name_with_search_button(driver, unit_name):
    # navigate to the urlpage
    driver.get(URL_PAGE)

    # wait for the dropdown to be clickable
    dropdown = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, 'betterselecter-sel')))
    dropdown.click()

    # wait for the options to be visible
    options = WebDriverWait(driver, 5).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@class="betterselecter-op"]')))

    # find the relevant unit we wish to analyze
    found_option = False
    for option in options:
        if unit_name in option.text:
        #if option.text == u'-------------- מרכז שרון – {}'.format(unit_name):
        #if option.text == u'-------------- מרכז שרון – רועי דיטל':
            print("### Found the option")
            option.click()
            found_option = True
            break

    if not found_option:
        QMessageBox.information(None, "שגיאה", "לא נמצאה יחידה עם שם זה")
        exit(0)
        
    driver.find_element(By.ID, 'searchButton').click()

    # find the table
    tables = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'tbl_chart')))

    # select the second table
    table = tables[1]
    return table


def main():
    app = QApplication([])
    driver, unit_name = do_login()
    if not driver:
        print("error occured. exiting gracefully")
        exit(0)

    table = filter_unit_name_with_search_button(driver, unit_name)
    

    # find all rows in the table whose id starts with 'family_'
    rows = table.find_elements(By.XPATH, './/tr[starts-with(@id, "family_")]')
    # find the headers of the table
    headers = table.find_elements(By.TAG_NAME, 'th')

    # find the index of the column "פגישה אחרונה"
    last_meeting_index = None
    next_meeting_index = None
    case_old_index = None
    budget_date_index = None
    for i, header in enumerate(headers):
        if header.text == u'פגישה אחרונה':
            last_meeting_index = i
            print(f'### last meeting index:{last_meeting_index}')
        if header.text == u'פגישה הבאה':
            next_meeting_index = i
            print(f'### last meeting index:{next_meeting_index}')
        if header.text == u'וותק':
            case_old_index = i
            print(f'### case old index:{case_old_index}')
        if header.text == u'תקציב בתוקף':
            budget_date_index = i
            print(f'### budget date index:{budget_date_index}')
        if u'מלווה' in header.text:
            assigned_to_index = i
            print(f'### assigned_to_index: {assigned_to_index}')

    # if the column was not found, print an error message and exit
    if not last_meeting_index or not next_meeting_index:
        print('### error: Column not found')
        driver.quit()
        exit(1)

    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Extract and write headers
    header_values = [header.text for header in headers]
    ws.append(header_values)

    # Iterate through each row and extract data
    for row_index, row in enumerate(rows, start=1):
        # Assuming each row contains multiple cells and you want to extract text from each cell
        cells = row.find_elements(By.TAG_NAME, "td")
        row_data = [cell.text for cell in cells]
        ws.append(row_data)
        # for col_index, cell in enumerate(cells, start=1):
        #     ws.cell(row=row_index, column=col_index, value=cell.text)

    # Save the workbook
    wb.save("output1.xlsx")

    alerts = []
    # iterate over the rows
    for row in rows:
        # find all cells in the row
        cells = row.find_elements(By.TAG_NAME, 'td')

        # get budget_date value
        budget_date = cells[budget_date_index].text.strip()
        # get the cell's value
        case_old_value = cells[case_old_index].text.strip()
        # split the value into words
        words = case_old_value.split()
        # check if the first word is a number
        if words[0].isdigit():
            # convert the first word to an integer
            case_old_days = int(words[0])
            # if the number is greater than 45, print a message
            if case_old_days > 45 and not budget_date:
                alerts.append(f'למשפחת {cells[0].text} של המלווה {cells[assigned_to_index].text} אין תקציב והליווי כבר בן יותר מ-45 יום')

        if not cells[last_meeting_index].text.strip():
            alerts.append(f'למשפחת {cells[0].text} של המלווה {cells[assigned_to_index].text} אין פגישה אחרונה בתיק')
        else:
            # parse the date string
            last_meeting_date = datetime.datetime.strptime(cells[last_meeting_index].text.strip(), "%d-%m-%y")
            # get the current date
            current_date = datetime.datetime.now()
            # get the current month and the previous month
            current_month = current_date.month
            previous_month = current_month - 1 if current_month != 1 else 12
            # if the last meeting date's month is not the same as the current month or the previous month, print the alert
            if last_meeting_date.month != current_month and last_meeting_date.month != previous_month:
                alerts.append(f'לא התקיימה פגישה עם משפחת {cells[0].text} של המלווה {cells[assigned_to_index].text}, בחודש האחרון')

        if not cells[next_meeting_index].text.strip():
            alerts.append(f'לא נקבעה הפגישה הבאה למשפחת {cells[0].text} של המלווה {cells[assigned_to_index].text}')

    # concat all the alerts into one message (line by line)
    alert_message = "\n".join(alerts)
    bullet_points = "\n".join([u"\u2022 " + alert for alert in alerts])

    QMessageBox.information(None, "התראות בצוות שלך:", bullet_points)
    #messagebox.showinfo(":התראות בצוות שלך", alert_message)
    QMessageBox.information(None, "סיום", "אשמח לשמוע הצעות לשיפור שאפשר להוסיף או לשנות או כל בעייה שנתקלת בה, אני זמין במייל או בווטסאפ 0544661404 רועי roidital@gmail.com תודה ולהתראות")

    # close the browser
    driver.quit()
    #input("Press Enter to exit...")


if __name__ == "__main__":
    main()



