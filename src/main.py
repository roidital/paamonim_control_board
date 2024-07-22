import os
import openpyxl
from src.common.constants import EXCEL_FILENAME, FAMILIES_SHEET_NAME, FAMILIES_SHEET_FIRST_ROW_NUM, \
    URL_FAMILIES_STATUS_PAGE, FamilyStatus
from src.create_families_sheet import create_families_sheet
from src.create_teams_list_sheet import create_teams_list_sheet, collect_tutor_families


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
        team_leader_to_families = await create_teams_list_sheet(browser, unit_name, wb)
        if team_leader_to_families is None:
            return None

    if do_families_sheet:
        if not do_teams_list_sheet:
            _, team_leader_to_families = await collect_tutor_families(browser, unit_name,
                                                                                URL_FAMILIES_STATUS_PAGE,
                                                                                FamilyStatus.ACTIVE)
            if team_leader_to_families is None:
                return None

        sheet = wb[FAMILIES_SHEET_NAME]
        ret_create_families = await create_families_sheet(sheet, browser, FAMILIES_SHEET_FIRST_ROW_NUM, team_leader_to_families, unit_name, do_email_list_sheet, lock)
        if not ret_create_families:
            return None

    save_workbook(wb)
    print(f'### DONE')
    return True
