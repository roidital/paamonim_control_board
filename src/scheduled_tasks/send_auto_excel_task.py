import sys
import os
script_dir = os.path.dirname(os.path.abspath(__file__))  # Directory of the script
project_root = os.path.dirname(os.path.dirname(script_dir))  # Project root directory
# Append the project root to sys.path
sys.path.append(project_root)

from login.login import auto_login
from src.common.common_utils import connect_to_db, send_email
from mysql.connector import Error
import asyncio
from src.main import main

# create a lock to be used later on but it must be created in the main thread where asyncio.run() is called
lock = asyncio.Lock()


async def fetch_all_users_details_from_db():
    connection = connect_to_db()
    if connection:
        try:
            cursor = connection.cursor()
            fetch_query = "SELECT username, password, unit_name FROM users_details"
            cursor.execute(fetch_query)
            for (username, password, unit_name) in cursor:
                ret = await generate_auto_excel(username, password, unit_name)
                if not ret:
                    print(f'### ERROR: generate_auto_excel() returned None')

        except Error as e:
            print(f"Failed to fetch data: {e}")
        finally:
            cursor.close()
            connection.close()


async def generate_auto_excel(username, password, unit_name):
    browser = await auto_login(username, password)
    if not browser:
        print(f'### ERROR: wrong username or password fetched from DB. username: {username}')
        return None

    ret_value = await main(browser, unit_name, True, True, False, lock)
    await browser.close()
    if not ret_value:
        print(f'### error occurred while creating the auto excel for user: {username} unit_name: {unit_name}')
        return None
    attachment_filename = os.environ.get('EXCEL_FILENAME', '')
    if attachment_filename is None or not os.path.exists(attachment_filename):
        print(f"### failed to create the excel file for user: {username} of unit: {unit_name}. No email will be sent")
        return None

    send_email(username, "האקסל המעודכן החודשי שלך הגיע", "מצורף בזאת האקסל המעודכן החודשי שלך, אם אינך מעוניינ/ת בקבלת מייל זה בעתיד אנא שלח/י לי מייל חוזר עם בקשה להסירך", attachment_filename)
    os.remove(attachment_filename)
    # cleanup temp files which were written during the excel creation
    os.system('rm -rf /tmp/.*')
    return True

if __name__ == "__main__":
    current_day = datetime.now().day
    if current_day == 1: # sends the excel file only if it's the first day of the month
        asyncio.run(fetch_all_users_details_from_db())
        
