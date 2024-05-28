import os
from typing import Final
# from pyvirtualdisplay import Display
from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from flask import flash, redirect, url_for

CRED_FILE: Final[str] = "paamonim_cred.txt"
LOGIN_URL: Final[str] = 'https://app.paamonim.org.il'


def _do_login(username, password, unit_name):
    # try to load user/password from early logins
    # creds_exists = False
    # if os.path.exists(CRED_FILE):
    #     with open(CRED_FILE, "r") as file:
    #         username = file.readline().rstrip("\n")
    #         password = file.readline().rstrip("\n")
    #         creds_exists = True
    # else:
    #     username, ok = QInputDialog.getText(None, "התחברות", "מה היוזר(האימייל) שלך בפעמונים?")
    #     password, ok = QInputDialog.getText(None, "התחברות", "מה הסיסמת התחברות שלך בפעמונים?", QLineEdit.Password)
    #
    
    # create a new virtual display
    #display = Display(visible=0, size=(800, 600))
    #display.start()

    # create an Options instance
    # options = Options()
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")

    # create a new Chrome browser instance
    browser = webdriver.Chrome(options=options)

    # navigate to the login page
    browser.get(LOGIN_URL)

    # find the username and password fields and enter your credentials
    browser.find_element(By.NAME, 'login').send_keys(username)
    browser.find_element(By.NAME, 'password').send_keys(password)

    try:
        # submit the form
        browser.find_element(By.ID, 'loginBtn').click()

        logout_button = browser.find_element(By.XPATH, "//a[text()='התנתק']")
        # if reach here with no exception - it means logout button is found - login successful
        print("Login successful")

        # if not creds_exists:
        #     save_cred = QMessageBox.question(None, "התחברות", "תרצה שהיוזר והסיסמא שלך ישמרו? (כך שלא תצטרך להזין אותם להבא", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        #     if save_cred == QMessageBox.Yes:
        #         with open(CRED_FILE, "w") as file:
        #             file.write(username + "\n")
        #             file.write(password + "\n")

        # unit_name, ok = QInputDialog.getText(None, "התחברות", "מה שמך כפי שהוא מופיע במערכת, למשל עבור 'מרכז שרון - רועי דיטל' פשוט רשום 'רועי דיטל' בשורה זו")
        print(f"unit_name: {unit_name}")
        #QMessageBox.information(None, "הודעה", "לאחר הודעה זו יפתח חלון טרמינל שיריץ את התוכנה, אל תסגור/י חלון זה, פשוט המתן/י מספר שניות לתוצאה")

        return browser, unit_name

    except NoSuchElementException:
        print("login failed")
        browser.quit()
        return None, None
