import os
import sys
import asyncio
import nest_asyncio
import platform

# Add the project root directory to the PYTHONPATH
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
sys.path.append(project_root)

import openpyxl
# from PyQt5.QtSql import userName, password

from login.login import auto_login
from src.common.common_utils import set_cell_value
from src.common.constants import EXCEL_FILENAME, FAMILIES_SHEET_NAME, \
    URL_FAMILIES_STATUS_PAGE, FamilyStatus, YELLOW_FILL, BOLD_FONT, FAMILIES_SHEET_LAST_COLUMN_INDEX
from src.create_teams_list_sheet import create_teams_list_sheet, collect_families_data

nest_asyncio.apply()

def init_workbook(excel_filename):
    # copy the template file to the new excel file
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Directory of the script
    project_root = os.path.dirname(script_dir)  # Project root directory
    
    system = platform.system()
    if system == "Windows":
        os.system(f'copy /y "{project_root}\\cockpit_template.xlsx" "{script_dir}\\{excel_filename}"')
        excel_path = script_dir+'\\'+excel_filename
    elif system == "Linux":
        os.system(f'cp -f {project_root}/cockpit_template.xlsx {script_dir}/{excel_filename}')
        excel_path = script_dir+'/'+excel_filename
    else:
        print("### ERROR: your OS is not supported, currently only Windows and Linux are supported by this app")
        exit(1)
    
    # check if copy command ended successfully
    if not os.path.exists(excel_path):
        print("Error copying the template file")
        exit(1)

    # Load the Excel file
    wb = openpyxl.load_workbook(excel_path)

    # patch (can be moved to the cockpit_template.xlsx) - freeze the top 3 rows of the familes sheet
    sheet = wb[FAMILIES_SHEET_NAME]
    sheet.freeze_panes = sheet['A4']
    return wb


def save_workbook(wb):
    wb.save('cockpit.xlsx')
    os.environ['EXCEL_FILENAME'] = 'cockpit.xlsx'


async def main(browser, unit_name, do_teams_list_sheet, do_families_sheet, do_email_list_sheet, lock):
    # app = QApplication([])
    # browser, unit_name = _do_login()
    if not browser:
        print("error occurred. exiting gracefully")
        exit(0)

    wb = init_workbook(EXCEL_FILENAME)

    if do_teams_list_sheet:
        team_leader_to_families = await create_teams_list_sheet(browser, unit_name, wb, do_email_list_sheet, lock)
        if team_leader_to_families is None:
            return None

    if do_families_sheet:
        if not do_teams_list_sheet:
            _, team_leader_to_families = await collect_families_data(browser, unit_name,
                                                                     URL_FAMILIES_STATUS_PAGE,
                                                                     FamilyStatus.ACTIVE, wb, do_email_list_sheet, lock)
            if team_leader_to_families is None:
                return None

        # sheet = wb[FAMILIES_SHEET_NAME]
        # ret_create_families = await create_families_sheet(sheet, browser, FAMILIES_SHEET_FIRST_ROW_NUM, team_leader_to_families, unit_name, do_email_list_sheet, lock)
        # if not ret_create_families:
        #     return None
    sheet = wb[FAMILIES_SHEET_NAME]
    sort_sheet_by_column(sheet, 1)

    save_workbook(wb)
    print(f'### DONE. Your output file cockpit.xlsx was successfully created here - in current directory')
    return True


def sort_sheet_by_column(sheet, column_index):
    # Extract all rows from the sheet, starting from the 4th row (first row after headers)
    data = list(sheet.iter_rows(min_row=4, values_only=True))

    # Sort the data (excluding the headers) by the specified column
    sorted_data = sorted(data, key=lambda row: row[column_index])

    # Clear the fill, font and alignment off the alerts column
    clear_alerts_column(sheet, FAMILIES_SHEET_LAST_COLUMN_INDEX)

    # Write the headers and sorted data back to the sheet
    for row_idx, row in enumerate(sorted_data, start=4):
        for col_idx, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.value = None # need to clear old value in case new value is blank (None)
            set_cell_value(cell, value, adjust_width=True)

    restore_attributes_for_alerts_column(sheet, FAMILIES_SHEET_LAST_COLUMN_INDEX)


def clear_alerts_column(sheet, column_index):
    for row in range(4, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=column_index)
        cell.value = None
        cell.fill = openpyxl.styles.PatternFill(fill_type=None)
        cell.font = openpyxl.styles.Font()
        cell.alignment = openpyxl.styles.Alignment()


def restore_attributes_for_alerts_column(sheet, column_index):
    for row in range(4, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=column_index)
        if cell.value is not None:
            cell.fill = YELLOW_FILL
            cell.font = BOLD_FONT
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True)


async def local_main():
    lock = asyncio.Lock()
    # read username, password and unit_name from the command line
    username = sys.argv[1]
    password = sys.argv[2]
    unit_name = sys.argv[3]
    print(f"### Starting main username: {username}  password: {password}. unit_name: {unit_name}\n")
    browser = await auto_login(username, password)
    if not browser:
        print("יוזר או סיסמא שגויים")
        exit(1)
    ret_value = await main(browser, unit_name, True, True, False, lock)
    await browser.close()
    if not ret_value:
        print(f"היחידה שהזנת {unit_name} לא נמצאה, אנא וודא/י שהקלדת נכון ללא רווחים וסימני פיסוק")
        exit(1)
    # print("הפעולה הסתיימה בהצלחה")

    # file_path = os.path.join(os.getcwd(), 'cockpit.xlsx')
    # save_workbook(wb, file_path)
    # print(f"Excel file created successfully at: {file_path}")


if __name__ == "__main__":
    asyncio.run(local_main())
