import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

current_day = datetime.now().day
if current_day == 1:
    with open('/home/roidital/paamonim_control_board/emails.txt', 'r') as file:
        email_list = [line.strip() for line in file]

    # email content
    subject = "תזכורת מפעמונים למלא יתרת עוש באפליקציה"
    body = "היום הראשון לחודש - היכנסו לאפליקציה למלא את יתרת העו״ש של היום, כך תוכלו לעקוב אחר התקדמותכם מחודש לחודש ולוודא שהרישום שלכם באפליקציה תואם את מה שקורה בפועל"

    # your email credentials
    your_email = "roidital@gmail.com"
    your_password = "zshx lkzh bhpl rrqm"

    # login to the email server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(your_email, your_password)

    # send the email to each address
    for email in email_list:
        if '@' in email:
            # setup the email
            msg = MIMEMultipart()
            msg['From'] = your_email
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            text = msg.as_string()
            server.sendmail(your_email, email, text)

    # logout of the email server
    server.quit()
