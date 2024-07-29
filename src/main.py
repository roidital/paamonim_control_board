import os
import openpyxl

from src.common.common_utils import set_cell_value
from src.common.constants import EXCEL_FILENAME, FAMILIES_SHEET_NAME, \
    URL_FAMILIES_STATUS_PAGE, FamilyStatus, YELLOW_FILL, BOLD_FONT, FAMILIES_SHEET_LAST_COLUMN_INDEX
from src.create_teams_list_sheet import create_teams_list_sheet, collect_families_data


def init_workbook(excel_filename):
    # copy the template file to the new excel file
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Directory of the script
    project_root = os.path.dirname(script_dir)  # Project root directory
    os.system(f'cp -f {project_root}/cockpit_template.xlsx {script_dir}/{excel_filename}')

    # check if prev command ended successfully
    excel_path = script_dir+'/'+excel_filename
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
    print(f'### DONE')
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
