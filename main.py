import os
import yagmail
import pandas as pd
from datetime import datetime
import sched
import time

# create two-step verification:
#   generate password in my accountgoogle -> security -> apps passwords -> generate password for mail ->
#   copy key to working_sec_pwd
#   enter your mail in my_mail
#   you can fill in mail_to_send_errors with your mail to receive errors. For example mail doesn't exist

text_body = open("text.txt", 'rb')
body = text_body.read()
body = body.decode('utf-8')
text_body.close()
subject_text = open("subject.txt", 'rb')
subject = subject_text.read()
subject = subject.decode('utf-8')
subject_text.close()
attachment = 'attachment\\file_to_send.xlsx'

my_mail = "your_mail"
working_sec_pwd = "your_generated_password"


def get_receivers(filename):
    a = pd.read_excel(filename)
    a1 = a.values
    return a1


def send_mail():
    if os.path.exists('receivers.xlsx'):
        receivers = get_receivers('receivers.xlsx')
    elif os.path.exists('receivers.xls'):
        receivers = get_receivers('receivers.xls')
    else:
        receivers = []
        print("list is empty")
    yagmail.register(my_mail, working_sec_pwd)
    yag = yagmail.SMTP(my_mail)
    for mail in receivers:
        try:
            yag.send(to=mail[0], subject=subject, contents=body, attachments=attachment)
            print("\nmail send to: ", mail[0])
        except Exception:
            yag.send(to="mail_to_send_errors", subject="error in mail",
                     contents=str(str(mail[0]) + "got error while sending"))
            pass
    i = str(input("press q and enter to finish"))
    while i != "q":
        i = input("press q and enter to finish")


if __name__ == '__main__':
    a = []
    g = int(input("input 1 to start email_send\n"
                  "input 2 to schedule email_send\n"
                  "input 0 to exit\n"))
    if g == 0:
        exit()
    elif g == 1:
        send_mail()
    elif g == 2:
        a = input("input time to start email send in format yyyy-mm-dd hh:mm:ss\n")
        now = datetime.now()
        time_to_set = datetime.strptime(a, '%Y-%m-%d %H:%M:%S')
        difference = time_to_set - now
        seconds_delay = difference.total_seconds()
        print(seconds_delay)
        scheduler = sched.scheduler(time.time, time.sleep)
        scheduler.enter(seconds_delay, 1, send_mail)
        scheduler.run()
