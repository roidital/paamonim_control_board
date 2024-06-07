import re
import openpyxl
import datetime
from src.common.common_utils import filter_unit_name_with_search_button, set_cell_value, __apply_border_to_team_table
from src.common.constants import URL_FAMILIES_STATUS_PAGE, LIGHT_BLUE_FILL, FAMILIES_SHEET_NAME, YELLOW_FILL, \
    FAMILIES_SHEET_FIRST_COLUMN_INDEX, FAMILIES_SHEET_LAST_COLUMN_INDEX, DAYS_WITHOUT_BUDGET_LIMIT, MAIN_LOGIN_URL, \
    BUDGET_AND_BALANCES_PAGE
from selenium.webdriver.common.by import By
from collections import defaultdict
import asyncio
from pyppeteer import launch
import nest_asyncio
nest_asyncio.apply()


def create_families_sheet(wb, sheet_name, browser, start_row, tutor_to_families, unit_name, username, password):
    # to reset the checkboxes checked by previous steps
    browser.get(URL_FAMILIES_STATUS_PAGE)
    filter_unit_name_with_search_button(browser, unit_name)

    sheet = wb[sheet_name]

    rows = browser.find_elements(By.XPATH, './/tr[starts-with(@id, "family_")]')

    i = start_row
    family_data_dict = defaultdict(lambda: [])
    for (tutor, families) in tutor_to_families.items():
        # create a header line for this tutor
        # print(f'### tutor: {tutor} line {i}')
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
                    # get the the family's id number (from html)
                    family_id = row.get_attribute('id').split('_')[1]  # Get the id attribute
                    #print(f'### family id: {family_id}')  # Print the number
                    retrieve_data_from_common_families_table(row, family_id, family_data_dict)
                    family_data_dict[family_id]['line_num'] = i
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX),
                                   family_data_dict[family_id]['unit_name'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+1),
                                   family_data_dict[family_id]['family_name'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+2),
                                   family_data_dict[family_id]['city'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+3),
                                   family_data_dict[family_id]['case_age'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+4),
                                   family_data_dict[family_id]['last_meeting_date'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+5),
                                   family_data_dict[family_id]['next_meeting_date'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+6),
                                   family_data_dict[family_id]['num_of_meetings'], adjust_width=True)
                    if 'num_cancelled_meetings' in family_data_dict[family_id]:
                        set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+7),
                                       family_data_dict[family_id]['num_cancelled_meetings'], adjust_width=True)
                    set_cell_value(sheet.cell(row=i, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX+17),
                                   family_data_dict[family_id]['last_osh_stats'], adjust_width=True)

                    num_skip_lines = write_family_alerts(family_data_dict[family_id], sheet, i)
                    i += num_skip_lines if num_skip_lines > 0 else 1
                    break
    asyncio.run(browser_dispatcher(family_data_dict, username, password))
    #print(f'### AFTER family_data_dict: {family_data_dict}')
    for family_id in family_data_dict.keys():
        line_num = family_data_dict[family_id]['line_num']
        if 'budget_income' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 8),
                           family_data_dict[family_id]['budget_income'], adjust_width=True)
        if 'budget_expense' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 9),
                           family_data_dict[family_id]['budget_expense'], adjust_width=True)
        if 'budget_diff' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 10),
                           family_data_dict[family_id]['budget_diff'], adjust_width=True)
        set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 11),
                       family_data_dict[family_id]['total_debts'], adjust_width=True)
        set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 12),
                       family_data_dict[family_id]['monthly_debts_payment'], adjust_width=True)
        set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 13),
                       family_data_dict[family_id]['unsettled_debts'], adjust_width=True)
        if 'month_income' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 14),
                           family_data_dict[family_id]['month_income'], adjust_width=True)
        if 'month_expense' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 15),
                           family_data_dict[family_id]['month_expense'], adjust_width=True)
        if 'last_month_diff' in family_data_dict[family_id]:
            set_cell_value(sheet.cell(row=line_num, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 16),
                           family_data_dict[family_id]['last_month_diff'], adjust_width=True)

    num_of_table_rows = i
    __apply_border_to_team_table(wb[FAMILIES_SHEET_NAME], 1, num_of_table_rows - 1,
                                 FAMILIES_SHEET_FIRST_COLUMN_INDEX,
                                 (FAMILIES_SHEET_LAST_COLUMN_INDEX-FAMILIES_SHEET_FIRST_COLUMN_INDEX))


# family_data_dict is an output parameter, a dictionary populated families data
def retrieve_data_from_common_families_table(row, family_id, family_data_dict):
    cells = row.find_elements(By.TAG_NAME, "td")
    if cells[0].text:
        # todo: consider moving the keys in the inner dict to constants in constants.py
        family_data_dict[family_id] = {'family_name': cells[0].text,
                                       'unit_name': cells[1].text,
                                       'city': cells[2].text,
                                       'last_meeting_date': cells[12].text,
                                       'next_meeting_date': cells[13].text,
                                       # 'num_of_meetings': cells[14].text,
                                       #  'num_cancelled_meetings': cells[14].text,
                                       'last_shikuf_bitsua': cells[7].text,
                                       'last_osh_stats': cells[15].text,
                                       'total_debts': cells[9].text,
                                       'monthly_debts_payment': cells[11].text,
                                       'unsettled_debts': cells[10].text,
                                       'budget': cells[8].text,
                                       'case_age': cells[6].text}
        # parse the number of meetings and the number of cancelled meetings from the format "num_meetings (num_cancelled_meetings)"
        match = re.match(r"(\d+) \((\d+)\)", cells[14].text)
        if match:
            family_data_dict[family_id]['num_of_meetings'] = match.group(1)
            family_data_dict[family_id]['num_cancelled_meetings'] = match.group(2)
        else:
            family_data_dict[family_id]['num_of_meetings'] = cells[14].text


async def auto_login(username, password):
    options = {
        'ignoreHTTPSErrors': True,
        'args': ['--no-sandbox'],
        'handleSIGINT': False,
        'handleSIGTERM': False,
        'handleSIGHUP': False
    }
    browser = await launch(options=options)
    page = await browser.newPage()

    # navigate to the login page
    await page.goto(MAIN_LOGIN_URL)

    # Perform login
    await page.type('input[name=login]', username)
    await page.type('input[name=password]', password)
    await page.click('#loginBtn')

    # Wait for navigation to complete
    await page.waitForNavigation()

    return browser


async def browser_dispatcher(family_data_dict, username, password):
    # Perform login
    browser = await auto_login(username, password)

    tasks = [fetch_family_data(browser, family_id, family_data_dict) for family_id in family_data_dict.keys() if family_data_dict[family_id]['last_shikuf_bitsua'] != '']
    pages_content = await asyncio.gather(*tasks)

    for page_content in pages_content:
        print(f'### page_content: {page_content}')

    await browser.close()


async def fetch_family_data(browser, family_id, family_data_dict):
    page = await browser.newPage()
    await page.goto(BUDGET_AND_BALANCES_PAGE + family_id)

    try:
        await page.waitForSelector('#expenseTable')
    except:
        print(f'### ERROR - family {family_id} got timed out while waiting for #expenseTable')
        await page.close()
        return 'done with family_id: ' + family_id

    budget_income = await page.evaluate('''() => {
            const td = document.querySelector('#sumTable tr.incomeTitle td.sumLine1Col1');
            return td ? td.innerHTML : null;
        }''')
    print(f'### income: {budget_income}')
    family_data_dict[family_id]['budget_income'] = budget_income
    budget_expense = await page.evaluate('''() => {
                const td = document.querySelector('#sumTable tr.expenseTitle td.sumLine2Col1');
                return td ? td.innerHTML : null;
            }''')
    print(f'### expense: {budget_expense}')
    family_data_dict[family_id]['budget_expense'] = budget_expense
    if budget_income and budget_expense:
        print(f'### budget diff: {int(budget_income) - int(budget_expense)}')
        family_data_dict[family_id]['budget_diff'] = int(budget_income) - int(budget_expense)
    month_income = await page.evaluate('''() => {
                    const td = document.querySelector('#sumTable tr.incomeTitle td.sumLine1Col3');
                    return td ? td.innerHTML : null;
                }''')
    print(f'### last month income: {month_income}')
    family_data_dict[family_id]['month_income'] = month_income
    month_expense = await page.evaluate('''() => {
                        const td = document.querySelector('#sumTable tr.expenseTitle td.sumLine2Col3');
                        return td ? td.innerHTML : null;
                    }''')
    print(f'### last month expense: {month_expense}')
    family_data_dict[family_id]['month_expense'] = month_expense
    if month_income and month_expense:
        print(f'### last month diff: {int(month_income) - int(month_expense)}')
        family_data_dict[family_id]['last_month_diff'] = int(month_income) - int(month_expense)

    await page.close()
    return 'done with family_id: ' + family_id


def write_family_alerts(cells, sheet, row):
    print(f'### alerts. cells: {cells}')
    alerts = []
    if not cells['budget'] and int(cells['case_age'].split()[0]) > DAYS_WITHOUT_BUDGET_LIMIT:
        alerts.append("ליווי בן יותר מ-45 יום ועדיין ללא תקציב ")
    if not cells['last_meeting_date'].strip():
        alerts.append("אין פגישה אחרונה בתיק")
    else:
        # parse the date string
        last_meeting_date = datetime.datetime.strptime(cells['last_meeting_date'].strip(), "%d-%m-%y")
        # get the current date
        current_date = datetime.datetime.now()
        # get the current month and the previous month
        current_month = current_date.month
        previous_month = current_month - 1 if current_month != 1 else 12
        # if the last meeting date's month is not the same as the current month or the previous month, print the alert
        if last_meeting_date.month != current_month and last_meeting_date.month != previous_month:
            alerts.append(
                f'לא התקיימה פגישה בחודש הנוכחי או הקודם')
    if not cells['next_meeting_date'].strip():
        alerts.append("לא נקבעה הפגישה הבאה")

    for i, alert in enumerate(alerts, start=1):
        if i > 1:
            sheet.insert_rows(row + i - 1)
        set_cell_value(sheet.cell(row=row + i - 1, column=FAMILIES_SHEET_LAST_COLUMN_INDEX), alert, fill=YELLOW_FILL)
        # Adjust the width of the column to text length #todo: put this adjustment into a function
        column_letter = openpyxl.utils.get_column_letter(FAMILIES_SHEET_LAST_COLUMN_INDEX)
        if len(sheet.cell(row + i - 1, FAMILIES_SHEET_LAST_COLUMN_INDEX).value) > sheet.column_dimensions[column_letter].width:
            sheet.column_dimensions[column_letter].width = len(sheet.cell(row+i-1, FAMILIES_SHEET_LAST_COLUMN_INDEX).value)

    print(f'### alerts for family {cells["family_name"]}: {alerts}')
    return len(alerts)
