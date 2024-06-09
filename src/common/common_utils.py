from openpyxl.styles import Alignment
import unicodedata
import openpyxl
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from src.common.constants import LEFT_TOP_BORDER, RIGHT_TOP_BORDER, TOP_BORDER, BOTTOM_BORDER, LEFT_BOTTOM_BORDER, \
    RIGHT_BOTTOM_BORDER, LEFT_BORDER, RIGHT_BORDER, BOLD_FONT, FamilyStatus, HEADERS_ROW_NUM, LIGHT_BLUE_FILL


# app = QApplication([])  # QApplication instance is required for QMessageBox


def normalize_string(s):
    """
    Normalize a string by removing diacritics and converting to lowercase.
    """
    return unicodedata.normalize("NFD", s).casefold()


def set_cell_value(cell, value, fill=None, font=BOLD_FONT, adjust_width=False):
    cell.value = value
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    if adjust_width:
        __adjust_column_width_to_text(cell)


def filter_unit_name_no_search_button(browser, filter_by, unit_name):
    # wait for the dropdown to be clickable
    dropdown = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, 'betterselecter-sel')))
    dropdown.click()

    # wait for the options to be visible
    WebDriverWait(browser, 5).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@class="betterselecter-op"]')))

    # count the rows before filtering the table
    current_row_count = len(browser.find_elements(By.XPATH, f'.//tr[starts-with(@id, {filter_by})]'))
    # print(f'### current_row_count: {current_row_count}')

    try:
        option = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.XPATH, f'//div[@class="betterselecter-op" and contains(text(), "{unit_name}")]')))
        print("### Found the option")
        option.click()
    except TimeoutException:
        print(f'### unit name: {unit_name} not found')
        # QMessageBox.information(None, "שגיאה", "לא נמצאה יחידה עם שם זה")
        exit(0)

    def __rows_have_updated(browser):
        new_row_count = len(browser.find_elements(By.XPATH, f'.//tr[starts-with(@id, {filter_by})]'))
        # print(f'### new_row_count: {new_row_count}')
        return new_row_count != current_row_count

    # wait for the table to be updated after unit filter
    try:
        WebDriverWait(browser, 5).until(__rows_have_updated)
    except TimeoutException:
        print("got timeout while waiting for rows number to change")


def filter_unit_name_with_search_button(browser, unit_name, families_status = FamilyStatus.ACTIVE):
    # wait for the dropdown to be clickable
    dropdown = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, 'betterselecter-sel')))
    dropdown.click()

    # wait for the options to be visible
    options = WebDriverWait(browser, 5).until(
        EC.visibility_of_all_elements_located((By.XPATH, '//div[@class="betterselecter-op"]')))

    # find the relevant unit we wish to analyze
    found_option = False
    for option in options:
        if unit_name in option.text:
            print(f"### Found the option {unit_name}")
            option.click()
            found_option = True
            break

    if not found_option:
        # QMessageBox.information(None, "שגיאה", "לא נמצאה יחידה עם שם זה")
        print(f'### ERROR - unit name: {unit_name} not found')
        exit(0)

    if families_status == FamilyStatus.READY_TO_START:
        browser.find_element(By.ID, 'started').click()
        browser.find_element(By.ID, 'ordered').click()

    browser.find_element(By.ID, 'searchButton').click()

    # find the table
    tables = WebDriverWait(browser, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'tbl_chart')))

    # select the second table
    table = tables[1]
    return table


def set_sum_formula_to_cell(sheet, start_row, end_row, column_index, divide_by_2=False):
    column_letter = openpyxl.utils.get_column_letter(column_index)
    formula = f'=SUM({column_letter}{start_row}:{column_letter}{end_row - 1})'
    set_cell_value(sheet.cell(row=end_row, column=column_index), (formula + '/2' if divide_by_2 else formula), LIGHT_BLUE_FILL)


def __apply_border_to_team_table(sheet, start_row, end_row, first_column_index, last_column_shift):
    sheet.cell(row=start_row, column=first_column_index).border = LEFT_TOP_BORDER
    sheet.cell(row=start_row, column=first_column_index + last_column_shift).border = RIGHT_TOP_BORDER
    for column in range(first_column_index + 1, first_column_index + last_column_shift):
        sheet.cell(row=start_row, column=column).border = TOP_BORDER
        sheet.cell(row=end_row, column=column).border = BOTTOM_BORDER

    sheet.cell(row=end_row, column=first_column_index).border = LEFT_BOTTOM_BORDER
    sheet.cell(row=end_row, column=first_column_index + last_column_shift).border = RIGHT_BOTTOM_BORDER
    for row in range(start_row + 1, end_row):
        sheet.cell(row=row, column=first_column_index).border = LEFT_BORDER
        sheet.cell(row=row, column=first_column_index + last_column_shift).border = RIGHT_BORDER


def __find_header_index(sheet, header_name):
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(row=HEADERS_ROW_NUM, column=col).value
        if cell_value and normalize_string(cell_value) == normalize_string(header_name):
            return col

    print(f"Header '{header_name}' not found.")
    return None


def __adjust_column_width_to_text(cell):
    column_letter = openpyxl.utils.get_column_letter(cell.column)
    cell_value_str = str(cell.value)
    if len(cell_value_str) > cell.parent.column_dimensions[column_letter].width:
        cell.parent.column_dimensions[column_letter].width = len(cell_value_str) * 1.1
