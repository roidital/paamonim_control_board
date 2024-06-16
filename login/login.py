from typing import Final
from pyppeteer import launch

from src.common.constants import MAIN_LOGIN_URL

CRED_FILE: Final[str] = "paamonim_cred.txt"
LOGIN_URL: Final[str] = 'https://app.paamonim.org.il'


# def _do_login(username, password):
#     # try to load user/password from early logins
#     # creds_exists = False
#     # if os.path.exists(CRED_FILE):
#     #     with open(CRED_FILE, "r") as file:
#     #         username = file.readline().rstrip("\n")
#     #         password = file.readline().rstrip("\n")
#     #         creds_exists = True
#     # else:
#     #     username, ok = QInputDialog.getText(None, "התחברות", "מה היוזר(האימייל) שלך בפעמונים?")
#     #     password, ok = QInputDialog.getText(None, "התחברות", "מה הסיסמת התחברות שלך בפעמונים?", QLineEdit.Password)
#     #
#
#     # create a new virtual display
#     #display = Display(visible=0, size=(800, 600))
#     #display.start()
#
#     # create an Options instance
#     # options = Options()
#     options = webdriver.ChromeOptions()
#     options.add_argument("--no-sandbox")
#     options.add_argument("--headless")
#     options.add_argument("--disable-gpu")
#
#     # create a new Chrome browser instance
#     browser = webdriver.Chrome(options=options)
#
#     # navigate to the login page
#     browser.get(LOGIN_URL)
#
#     # find the username and password fields and enter your credentials
#     browser.find_element(By.NAME, 'login').send_keys(username)
#     browser.find_element(By.NAME, 'password').send_keys(password)
#
#     try:
#         # submit the form
#         browser.find_element(By.ID, 'loginBtn').click()
#
#         logout_button = browser.find_element(By.XPATH, "//a[text()='התנתק']")
#         # if reach here with no exception - it means logout button is found - login successful
#         print("Login successful")
#
#         # if not creds_exists:
#         #     save_cred = QMessageBox.question(None, "התחברות", "תרצה שהיוזר והסיסמא שלך ישמרו? (כך שלא תצטרך להזין אותם להבא", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
#         #     if save_cred == QMessageBox.Yes:
#         #         with open(CRED_FILE, "w") as file:
#         #             file.write(username + "\n")
#         #             file.write(password + "\n")
#
#         # unit_name, ok = QInputDialog.getText(None, "התחברות", "מה שמך כפי שהוא מופיע במערכת, למשל עבור 'מרכז שרון - רועי דיטל' פשוט רשום 'רועי דיטל' בשורה זו")
#         #print(f"unit_name: {unit_name}")
#         #QMessageBox.information(None, "הודעה", "לאחר הודעה זו יפתח חלון טרמינל שיריץ את התוכנה, אל תסגור/י חלון זה, פשוט המתן/י מספר שניות לתוצאה")
#
#         return browser
#
#     except NoSuchElementException:
#         print("login failed")
#         browser.quit()
#         return None
async def auto_login(username, password):
    options = {
        'ignoreHTTPSErrors': True,
        'args': ['--no-sandbox --headless --disable-gpu'],
        # since we are using Flask to launch this webapp - this flow is not running in main thread, in python signals
        # can be set only in main thread, so we need to disable the signals handling
        'handleSIGINT': False,
        'handleSIGTERM': False,
        'handleSIGHUP': False
    }
    browser = await launch(options=options)
    page = await browser.newPage()

    # navigate to the login page
    await page.goto(MAIN_LOGIN_URL)

    await page.type('input[name=login]', username)
    await page.type('input[name=password]', password)
    await page.click('#loginBtn')

    # Wait for navigation to complete
    await page.waitForNavigation()

    # Check if login was successful
    try:
        await page.waitForSelector('a[href="/login/logout"]', timeout=10000)
    except:
        print("Login failed")
        return None

    print(f'### Login successful')

    return browser
