from app import cleanup
from login.login import auto_login
from src.common.common_utils import connect_to_db, send_email
from mysql.connector import Error
import asyncio
from src.main import main
from flask import session
import os
#import nest_asyncio

#nest_asyncio.apply()

# create a lock to be used later on but it must be created in the main thread where asyncio.run() is called
lock = asyncio.Lock()


def fetch_all_users_details_from_db():
    connection = connect_to_db()
    if connection:
        try:
            cursor = connection.cursor()
            fetch_query = "SELECT username, password, unit_name FROM users_details"
            cursor.execute(fetch_query)
            for (username, password, unit_name) in cursor:
                #print(f"Username: {username}, Password: {password}, Unit Name: {unit_name}")
                ret = generate_auto_excel(username, password, unit_name)
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
        print(f'### error occurred while creating the auto excel')
        return None
    attachment_filename = session.get('temp_file')
    send_email(username, "your Paamonim's Excel file is attached", "attached below is your Excel", attachment_filename)
    os.remove(attachment_filename)
    cleanup()


if __name__ == "__main__":
    asyncio.run(fetch_all_users_details_from_db)
