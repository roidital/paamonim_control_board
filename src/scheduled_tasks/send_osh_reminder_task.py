from datetime import datetime

from src.common.common_utils import send_email


def send_osh_reminder_email():
    current_day = datetime.now().day
    if current_day == 1: # sends the reminder email only if it's the first day of the month
        with open('/home/roidital/paamonim_control_board/emails.txt', 'r') as file:
            email_list = [line.strip() for line in file]
            for email_address in email_list:
                if '@' in email_address:
                    subject = "תזכורת מפעמונים למלא יתרת עוש באפליקציה"
                    body = "היום הראשון לחודש - היכנסו לאפליקציה למלא את יתרת העו״ש של היום, כך תוכלו לעקוב אחר התקדמותכם מחודש לחודש ולוודא שהרישום שלכם באפליקציה תואם את מה שקורה בפועל"
                    body += "\n במידה ואינך מעוניינ/ת להמשיך ולקבל תזכורות אלו - אנא שלח/י לי מייל חוזר להסירך ואדאג לכך"
                    send_email(email_address, subject, body)


if __name__ == "__main__":
    send_osh_reminder_email()