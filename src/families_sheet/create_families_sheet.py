import openpyxl
import datetime
from src.common.common_utils import filter_unit_name_with_search_button, set_cell_value, __apply_border_to_team_table
from src.common.constants import URL_FAMILIES_STATUS_PAGE, LIGHT_BLUE_FILL, FAMILIES_SHEET_NAME, YELLOW_FILL, \
    FAMILIES_SHEET_FIRST_COLUMN_INDEX, FAMILIES_SHEET_LAST_COLUMN_INDEX, DAYS_WITHOUT_BUDGET_LIMIT
from selenium.webdriver.common.by import By


def create_families_sheet(wb, sheet_name, browser, start_row, tutor_to_families, unit_name):
    # to reset the checkboxes checked by previous steps
    browser.get(URL_FAMILIES_STATUS_PAGE)
    filter_unit_name_with_search_button(browser, unit_name)

    sheet = wb[sheet_name]

    rows = browser.find_elements(By.XPATH, './/tr[starts-with(@id, "family_")]')

    i = start_row
    for (tutor, families) in tutor_to_families.items():
        # create a header line for this tutor
        sheet.merge_cells(start_row=i, start_column=FAMILIES_SHEET_FIRST_COLUMN_INDEX, end_row=i, end_column=FAMILIES_SHEET_LAST_COLUMN_INDEX)
        merged_cell = sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX)
        set_cell_value(merged_cell, tutor, fill=LIGHT_BLUE_FILL)
        i += 1

        # for each family of this tutor search for the family name in the html and copy relevant fields to excel
        for family in families:
            # find the row in the html that contains the family name
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if family[0] in cells[0].text:
                    # copy the row to the excel sheet (it takes only selected fields from the row)
                    # todo: this is a horrible code:), please think of a nicer way to do that
                    row_data = [cells[0].text, cells[1].text, cells[2].text, cells[12].text,
                                cells[13].text, cells[14].text, cells[7].text, cells[15].text, cells[9].text,
                                cells[11].text, cells[10].text, cells[8].text, cells[6].text]
                    # write row_data to the excel at row
                    for col, cell in enumerate(row_data, start=FAMILIES_SHEET_FIRST_COLUMN_INDEX):
                        set_cell_value(sheet.cell(row=i, column=col), cell)
                        # Adjust the width of the column to text length #todo: put this adjustment into a function
                        column_letter = openpyxl.utils.get_column_letter(col)
                        if len(sheet.cell(i, col).value) > sheet.column_dimensions[column_letter].width:
                            sheet.column_dimensions[column_letter].width = len(sheet.cell(i, col).value)
                    num_skip_lines = write_family_alerts(cells, sheet, i)
                    i += num_skip_lines if num_skip_lines > 0 else 1
                    break

    num_of_table_rows = i
    __apply_border_to_team_table(wb[FAMILIES_SHEET_NAME], 1, num_of_table_rows - 1,
                                 FAMILIES_SHEET_FIRST_COLUMN_INDEX,
                                 (FAMILIES_SHEET_LAST_COLUMN_INDEX-FAMILIES_SHEET_FIRST_COLUMN_INDEX))


def write_family_alerts(cells, sheet, row):
    alerts = []
    if not cells[8].text and int(cells[6].text.split()[0]) > DAYS_WITHOUT_BUDGET_LIMIT:
        alerts.append("ליווי בן יותר מ-45 יום ועדיין ללא תקציב ")
    if not cells[12].text.strip():
        alerts.append("אין פגישה אחרונה בתיק")
    else:
        # parse the date string
        last_meeting_date = datetime.datetime.strptime(cells[12].text.strip(), "%d-%m-%y")
        # get the current date
        current_date = datetime.datetime.now()
        # get the current month and the previous month
        current_month = current_date.month
        previous_month = current_month - 1 if current_month != 1 else 12
        # if the last meeting date's month is not the same as the current month or the previous month, print the alert
        if last_meeting_date.month != current_month and last_meeting_date.month != previous_month:
            alerts.append(
                f'לא התקיימה פגישה בחודש הנוכחי או הקודם')
    if not cells[13].text.strip():
        alerts.append("לא נקבעה הפגישה הבאה")

    for i, alert in enumerate(alerts, start=1):
        if i > 1:
            sheet.insert_rows(row + i - 1)
        set_cell_value(sheet.cell(row=row + i - 1, column=FAMILIES_SHEET_LAST_COLUMN_INDEX), alert, fill=YELLOW_FILL)
        # Adjust the width of the column to text length #todo: put this adjustment into a function
        column_letter = openpyxl.utils.get_column_letter(FAMILIES_SHEET_LAST_COLUMN_INDEX)
        if len(sheet.cell(row + i - 1, FAMILIES_SHEET_LAST_COLUMN_INDEX).value) > sheet.column_dimensions[column_letter].width:
            sheet.column_dimensions[column_letter].width = len(sheet.cell(row+i-1, FAMILIES_SHEET_LAST_COLUMN_INDEX).value)

    print(f'### alerts for family {cells[0].text}: {alerts}')
    return len(alerts)

