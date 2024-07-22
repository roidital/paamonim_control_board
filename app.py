import asyncio
import nest_asyncio
from flask import Flask, render_template, request, send_file, abort, redirect, url_for, flash

from login.login import auto_login
from src.main import main
import os
from time import sleep
from src.register_to_auto_excel import register_user_to_receive_auto_excel

nest_asyncio.apply()

app_dir = os.path.dirname(os.path.abspath(__file__))
template_dir = os.path.join(app_dir, 'templates')
print(f"template_dir: {template_dir}")
app = Flask(__name__, template_folder=template_dir)
app.secret_key = os.urandom(24)


@app.route('/')
def home():
    return redirect(url_for('do_login'))


@app.route('/login', methods=['GET'])
def login_form():
    # Render the login form
    return render_template('login.html')


@app.route('/login', methods=['POST'])
def do_login():
    main_ret = asyncio.run(async_main())
    if main_ret:
        return main_ret
    return redirect(url_for('download_excel'))


async def async_main():
    # create a lock to be used later on but it must be created in the main thread where asyncio.run() is called
    lock = asyncio.Lock()
    # Get form data
    username = request.form.get('username')
    password = request.form.get('password')
    unit_name = request.form.get('unit_name')
    if not input_validation(username, password, unit_name):
        flash("אחד או יותר מהשדות ריקים, אנא מלא/י את כל השדות: שם משתמש, סיסמא ושם יחידה")
        return redirect(url_for('do_login'))
    do_teams_list_sheet = 'create_teams_list_sheet' in request.form
    do_families_sheet = 'create_families_sheet' in request.form
    create_email_list = 'create_email_list' in request.form
    do_register_to_auto_excel = 'register_to_auto_excel' in request.form
    browser = await auto_login(username, password)
    if not browser:
        flash("שגיאת התחברות, אנא בדוק/י שהיוזר והסיסמא נכונים")
        return redirect(url_for('do_login'))

    ret_value = await main(browser, unit_name, do_teams_list_sheet, do_families_sheet, create_email_list, lock)

    if do_register_to_auto_excel:
        register_user_to_receive_auto_excel(username, password, unit_name)

    await browser.close()
    if not ret_value:
        flash(f"היחידה שהזנת {unit_name} לא נמצאה, אנא וודא/י שהקלדת נכון ללא רווחים וסימני פיסוק")
        return redirect(url_for('do_login'))
    return None


def input_validation(username, password, unit_name):
    return username and password and unit_name.split()


@app.route('/download', methods=['GET'])
def download_excel():
    # Retrieve the temporary file's name from the session
    temp_file_name = os.environ.get('EXCEL_FILENAME', '')

    if temp_file_name is None or not os.path.exists(temp_file_name):
        sleep(3) # allow one retry
        if temp_file_name is None or not os.path.exists(temp_file_name):
            flash("הורדת הקובץ נכשלה")
            abort(404)

    # Create a Flask response with the Excel file
    response = send_file(temp_file_name, as_attachment=True, download_name='cockpit.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Delete the temporary file
    os.remove(temp_file_name)
    cleanup()
    return response


def cleanup():
    os.system('rm -rf /tmp/.*')


if __name__ == '__main__':
    app.run(debug=True)
