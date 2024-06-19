import re
import datetime
from src.common.common_utils import filter_unit_name_with_search_button, set_cell_value, __apply_border_to_team_table
from src.common.constants import URL_FAMILIES_STATUS_PAGE, YELLOW_FILL, \
    FAMILIES_SHEET_FIRST_COLUMN_INDEX, FAMILIES_SHEET_LAST_COLUMN_INDEX, DAYS_WITHOUT_BUDGET_LIMIT, \
    BUDGET_AND_BALANCES_PAGE, FAMILY_NAME, UNIT_NAME, CITY, LAST_MEETING_DATE, NEXT_MEETING_DATE, LAST_SHIKUF_BITSUA, \
    LAST_OSH_STATS, TOTAL_DEBTS, MONTHLY_DEBTS_PAYMENT, UNSETTLED_DEBTS, BUDGET, CASE_AGE, NUM_OF_MEETINGS, \
    NUM_CANCELLED_MEETINGS, BUDGET_INCOME, BUDGET_EXPENSE, BUDGET_DIFF, MONTH_INCOME, MONTH_EXPENSE, LAST_MONTH_DIFF, \
    TUTOR, DAYS_WITHOUT_FIRST_MEETING_LIMIT, OSH_STATS_PAGE, CURRENT_MONTH_OSH, LAST_MONTH_OSH
from collections import defaultdict
import asyncio
import nest_asyncio
nest_asyncio.apply()


async def create_families_sheet(sheet, browser, start_row, team_leader_to_families, unit_name):
    # to reset the checkboxes checked by previous steps
    page = await browser.newPage()
    await page.goto(URL_FAMILIES_STATUS_PAGE)
    await filter_unit_name_with_search_button(page, unit_name)

    rows = await page.querySelectorAll('tr[id^="family_"]')

    i = start_row
    family_data_dict = defaultdict(lambda: [])
    for (team_leader, families) in team_leader_to_families.items():
        # for each family of this tutor search for the family name in the html and copy relevant fields to excel
        for family in families:
            # find the row in the html that contains the family name
            for row in rows:
                cells = await row.querySelectorAll('td')
                cell0_value = await page.evaluate('(element) => element.textContent', cells[0])
                if family in cell0_value:
                    # get the the family's id number (from html)
                    family_id = await (await row.getProperty('id')).jsonValue()
                    family_id = family_id.split('_')[1]
                    await retrieve_data_from_common_families_table(page, row, family_id, family_data_dict)
                    family_data_dict[family_id]['line_num'] = i
                    set_values_from_common_families_table_to_excel(family_data_dict[family_id], sheet)

                    write_family_alerts(family_data_dict[family_id], sheet, i)
                    i += 1
                    break
    # retrieve the budget and balances data for each family (parallel execution)
    await browser_dispatcher(family_data_dict, browser)
    # print(f'### AFTER family_data_dict: {family_data_dict}')

    for family_id in family_data_dict.keys():
        set_budget_and_balances_to_excel(family_data_dict[family_id], sheet)

    num_of_table_rows = i
    __apply_border_to_team_table(sheet, 1, num_of_table_rows - 1,
                                 FAMILIES_SHEET_FIRST_COLUMN_INDEX,
                                 (FAMILIES_SHEET_LAST_COLUMN_INDEX-FAMILIES_SHEET_FIRST_COLUMN_INDEX))


# family_data_dict is an output parameter, a dictionary populated families data
async def retrieve_data_from_common_families_table(page, row, family_id, family_data_dict):
    cells = await row.querySelectorAll('td')
    cell0_value = await page.evaluate('(element) => element.textContent', cells[0])
    cell1_value = await page.evaluate('(element) => element.textContent', cells[1])
    cell2_value = await page.evaluate('(element) => element.textContent', cells[2])
    cell3_value = await page.evaluate('(element) => element.textContent', cells[3])
    cell6_value = await page.evaluate('(element) => element.textContent', cells[6])
    cell7_value = await page.evaluate('(element) => element.textContent', cells[7])
    cell8_value = await page.evaluate('(element) => element.textContent', cells[8])
    cell9_value = await page.evaluate('(element) => element.textContent', cells[9])
    cell10_value = await page.evaluate('(element) => element.textContent', cells[10])
    cell11_value = await page.evaluate('(element) => element.textContent', cells[11])
    cell12_value = await page.evaluate('(element) => element.textContent', cells[12])
    cell13_value = await page.evaluate('(element) => element.textContent', cells[13])
    cell14_value = await page.evaluate('(element) => element.textContent', cells[14])
    cell15_value = await page.evaluate('(element) => element.textContent', cells[15])
    if cell0_value:
        # todo: consider moving the keys in the inner dict to constants in constants.py
        family_data_dict[family_id] = {FAMILY_NAME: cell0_value,
                                       UNIT_NAME: cell1_value,
                                       CITY: cell2_value,
                                       TUTOR: cell3_value,
                                       LAST_MEETING_DATE: cell12_value,
                                       NEXT_MEETING_DATE: cell13_value,
                                       LAST_SHIKUF_BITSUA: cell7_value,
                                       LAST_OSH_STATS: cell15_value,
                                       TOTAL_DEBTS: cell9_value,
                                       MONTHLY_DEBTS_PAYMENT: cell11_value,
                                       UNSETTLED_DEBTS: cell10_value,
                                       BUDGET: cell8_value,
                                       CASE_AGE: cell6_value}
        # parse the number of meetings and the number of cancelled meetings from the format "num_meetings (num_cancelled_meetings)"
        match = re.match(r"(\d+) \((\d+)\)", cell14_value)
        if match:
            family_data_dict[family_id][NUM_OF_MEETINGS] = match.group(1)
            family_data_dict[family_id][NUM_CANCELLED_MEETINGS] = match.group(2)
        else:
            family_data_dict[family_id][NUM_OF_MEETINGS] = cell14_value


def set_values_from_common_families_table_to_excel(family_data, sheet):
    row = family_data['line_num']
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX),
                   family_data[UNIT_NAME], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 1),
                   family_data[TUTOR], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 2),
                   family_data[FAMILY_NAME], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 3),
                   family_data[CITY], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 4),
                   family_data[CASE_AGE], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 5),
                   family_data[LAST_MEETING_DATE], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 6),
                   family_data[NEXT_MEETING_DATE], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 7),
                   family_data[NUM_OF_MEETINGS], adjust_width=True)
    if NUM_CANCELLED_MEETINGS in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 8),
                       family_data[NUM_CANCELLED_MEETINGS], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 18),
                   family_data[LAST_OSH_STATS], adjust_width=True)


def set_budget_and_balances_to_excel(family_data, sheet):
    row = family_data['line_num']
    if BUDGET_INCOME in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 9),
                       family_data[BUDGET_INCOME], adjust_width=True)
    if BUDGET_EXPENSE in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 10),
                       family_data[BUDGET_EXPENSE], adjust_width=True)
    if BUDGET_DIFF in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 11),
                       family_data[BUDGET_DIFF], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 12),
                   family_data[TOTAL_DEBTS], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 13),
                   family_data[MONTHLY_DEBTS_PAYMENT], adjust_width=True)
    set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 14),
                   family_data[UNSETTLED_DEBTS], adjust_width=True)
    if MONTH_INCOME in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 15),
                       family_data[MONTH_INCOME], adjust_width=True)
    if MONTH_EXPENSE in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 16),
                       family_data[MONTH_EXPENSE], adjust_width=True)
    if LAST_MONTH_DIFF in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 17),
                       family_data[LAST_MONTH_DIFF], adjust_width=True)
    if CURRENT_MONTH_OSH in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 19),
                       family_data[CURRENT_MONTH_OSH], adjust_width=True)
    if LAST_MONTH_OSH in family_data:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 20),
                       family_data[LAST_MONTH_OSH], adjust_width=True)
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_FIRST_COLUMN_INDEX + 21),
                       int(family_data[CURRENT_MONTH_OSH].replace(',', ''))-int(family_data[LAST_MONTH_OSH].replace(',', '')), adjust_width=True)


async def browser_dispatcher(family_data_dict, browser):
    # in case the family doesn't have a shikuf/bitsua - it means the BUDGET_AND_BALANCES_PAGE page doesn't exist
    # which will follow a timeout exception in the fetch_family_data function
    tasks = [fetch_family_data(browser, family_id, family_data_dict) for family_id in family_data_dict.keys()
             if family_data_dict[family_id]['last_shikuf_bitsua'].strip() != '']
    osh_tasks = [fetch_family_osh_data(browser, family_id, family_data_dict) for family_id in family_data_dict.keys()]
    tasks.extend(osh_tasks)
    pages_content = await asyncio.gather(*tasks)

    # for page_content in pages_content:
    #     print(f'### page_content: {page_content}')

    await browser.close()


async def fetch_family_osh_data(browser, family_id, family_data_dict):
    page = await browser.newPage()
    try:
        await page.goto(OSH_STATS_PAGE + family_id, timeout=5000)
    except:
        print(f'### ERROR: family {family_id} got timedout while browsing to OSH page')
        return 'timeout for OSH page. family_id: ' + family_id
    
    rows = await page.querySelectorAll('tbody tr')
    if len(rows)> 0:
        tds = await rows[0].querySelectorAll('td')
        current_month_osh_td = tds[2] # the osh value is in the 3rd <td> element
        current_month_osh_value = await page.evaluate('(element) => element.textContent', current_month_osh_td)
        #print(f'### current_month_osh_value: {current_month_osh_value}')
        family_data_dict[family_id][CURRENT_MONTH_OSH] = current_month_osh_value
        if len(rows) > 1:
            tds = await rows[1].querySelectorAll('td')
            last_month_osh_td = tds[2] # the osh value is in the 3rd <td> element
            last_month_osh_value = await page.evaluate('(element) => element.textContent', last_month_osh_td)
            #print(f'### last_month_osh_value: {last_month_osh_value}')
            family_data_dict[family_id][LAST_MONTH_OSH] = last_month_osh_value

    return 'done with osh for family_id: ' + family_id


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
    # print(f'### income: {budget_income}')
    family_data_dict[family_id][BUDGET_INCOME] = budget_income
    budget_expense = await page.evaluate('''() => {
                const td = document.querySelector('#sumTable tr.expenseTitle td.sumLine2Col1');
                return td ? td.innerHTML : null;
            }''')
    # print(f'### expense: {budget_expense}')
    family_data_dict[family_id][BUDGET_EXPENSE] = budget_expense
    if budget_income and budget_expense:
        # print(f'### budget diff: {int(budget_income) - int(budget_expense)}')
        family_data_dict[family_id][BUDGET_DIFF] = int(budget_income) - int(budget_expense)
    month_income = await page.evaluate('''() => {
                    const td = document.querySelector('#sumTable tr.incomeTitle td.sumLine1Col3');
                    return td ? td.innerHTML : null;
                }''')
    # print(f'### last month income: {month_income}')
    family_data_dict[family_id][MONTH_INCOME] = month_income
    month_expense = await page.evaluate('''() => {
                        const td = document.querySelector('#sumTable tr.expenseTitle td.sumLine2Col3');
                        return td ? td.innerHTML : null;
                    }''')
    # print(f'### last month expense: {month_expense}')
    family_data_dict[family_id][MONTH_EXPENSE] = month_expense
    if month_income and month_expense:
        # print(f'### last month diff: {int(month_income) - int(month_expense)}')
        family_data_dict[family_id][LAST_MONTH_DIFF] = int(month_income) - int(month_expense)

    await page.close()
    return 'done with family_id: ' + family_id


def write_family_alerts(family_data, sheet, row):
    # print(f'### alerts. cells: {family_data}')
    alerts = []
    if not family_data[BUDGET] and int(family_data[CASE_AGE].split()[0]) > DAYS_WITHOUT_BUDGET_LIMIT:
        alerts.append("ליווי בן יותר מ-45 יום ועדיין ללא תקציב ")
    if not family_data[LAST_MEETING_DATE].strip():
        if int(family_data[CASE_AGE].split()[0]) > DAYS_WITHOUT_FIRST_MEETING_LIMIT:
            alerts.append("אין פגישה אחרונה בתיק שמתנהל מעל 30 יום")
    else:
        # parse the date string
        last_meeting_date = datetime.datetime.strptime(family_data[LAST_MEETING_DATE].strip(), "%d-%m-%y")
        current_date = datetime.datetime.now()
        current_month = current_date.month
        previous_month = current_month - 1 if current_month != 1 else 12
        # if the last meeting date's month is not the same as the current month or the previous month, add an alert
        if last_meeting_date.month != current_month and last_meeting_date.month != previous_month:
            alerts.append(
                f'לא התקיימה פגישה בחודש הנוכחי או הקודם')
    if not family_data[NEXT_MEETING_DATE].strip():
        alerts.append("לא נקבעה הפגישה הבאה")
    if MONTH_INCOME in family_data and BUDGET_INCOME in family_data:
        if family_data[MONTH_INCOME] < float(family_data[BUDGET_INCOME]*0.75):
            alerts.append("הכנסה חודשית נמוכה מ-75% מהתקציב החודשי")
    if MONTH_EXPENSE in family_data and BUDGET_EXPENSE in family_data:
        if family_data[MONTH_EXPENSE] > float(family_data[BUDGET_EXPENSE]*1.3):
            alerts.append("הוצאה חודשית גבוהה ביותר מ-30% מהתקציב החודשי")

    # concat all the alerts into one string with a new line separator
    alerts = '\n'.join(alerts)
    if alerts:
        set_cell_value(sheet.cell(row=row, column=FAMILIES_SHEET_LAST_COLUMN_INDEX), alerts, fill=YELLOW_FILL, adjust_width=True, wrap_text=True)

    # print(f'### alerts for family {family_data[FAMILY_NAME]}: {alerts}')
