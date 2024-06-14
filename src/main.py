import os
import openpyxl
from flask import session
import tempfile
from src.common.constants import EXCEL_FILENAME, FAMILIES_SHEET_NAME, FAMILIES_SHEET_FIRST_ROW_NUM, \
    URL_FAMILIES_STATUS_PAGE, FamilyStatus
from src.families_sheet.create_families_sheet import create_families_sheet
from src.teams_list_sheet.create_teams_list_sheet import create_teams_list_sheet, collect_tutor_families


def init_workbook(excel_filename):
    # copy the template file to the new excel file
    os.system(f'cp -f cockpit_template.xlsx {excel_filename}')

    # check if prev command ended successfully
    if not os.path.exists(excel_filename):
        print("Error copying the template file")
        exit(1)

    # Load the Excel file
    wb = openpyxl.load_workbook(excel_filename)
    return wb


def save_workbook(wb):
    # Create a temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False)
    # Save the workbook to the temporary file
    wb.save(temp_file.name)
    # Store the temporary file's name in the session
    session['temp_file'] = temp_file.name


def main(browser, unit_name, username, password, do_teams_list_sheet, do_families_sheet):
    # app = QApplication([])
    # browser, unit_name = _do_login()
    if not browser:
        print("error occurred. exiting gracefully")
        exit(0)

    wb = init_workbook(EXCEL_FILENAME)

    if do_teams_list_sheet:
        team_leader_to_families = create_teams_list_sheet(browser, unit_name, wb)

    if do_families_sheet:
        if not do_teams_list_sheet:
            _, team_leader_to_families = collect_tutor_families(browser, unit_name,
                                                                                URL_FAMILIES_STATUS_PAGE,
                                                                                FamilyStatus.ACTIVE)
        create_families_sheet(wb, FAMILIES_SHEET_NAME, browser, FAMILIES_SHEET_FIRST_ROW_NUM, team_leader_to_families, unit_name, username, password)

    save_workbook(wb)
    print(f'### DONE')
