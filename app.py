from flask import Flask, render_template, request, session, send_file, abort, redirect, url_for, flash
from src.main import main
from login.login import _do_login
import os

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
    # Get form data
    username = request.form.get('username')
    password = request.form.get('password')
    unit_name = request.form.get('unit_name')
    do_teams_list_sheet = 'create_teams_list_sheet' in request.form
    do_families_sheet = 'create_families_sheet' in request.form
    browser, unit_name = _do_login(username, password, unit_name)
    if not browser:
        flash("שגיאת התחברות, אנא בדוק/י שהיוזר והסיסמא נכונים")
        return redirect(url_for('do_login'))

    main(browser, unit_name, username, password, do_teams_list_sheet, do_families_sheet)
    return redirect(url_for('download_excel'))


@app.route('/download', methods=['GET'])
def download_excel():
    # Retrieve the temporary file's name from the session
    temp_file_name = session.get('temp_file')

    if temp_file_name is None or not os.path.exists(temp_file_name):
        # If there's no file to download, send a 404 Not Found response
        abort(404)

    # Create a Flask response with the Excel file
    response = send_file(temp_file_name, as_attachment=True, download_name='cockpit.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Delete the temporary file
    os.remove(temp_file_name)

    return response


if __name__ == '__main__':
    app.run(debug=True)
