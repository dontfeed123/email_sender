'''
    Version 0: The function functional
    Have:
        - Send emails to a specification list of emails
    Todo:
        - Errors Handling (important)
        - Multi-Threading
        - Log File
        - Not tested with a big amount of emails
'''

import smtplib
import openpyxl

from email.mime.text import MIMEText


class Global_Storage:
    '''
        This class stores all global variables, to all functions access it
    '''

    def __init__(self, email="testeautomateemails@gmail.com", password="testeTeste21-06-2018"):
        self.my_email = email  # This is to login in smpt server (needs to be gmail)
        self.my_password = password  # This is to login in smpt server


def email_sender(send_to: list, subject: str, message: str) -> None:

    '''
        This function send emails from a sender, that is the email where the emails will be sended from,
        to a list of emails, with the same subject and message.

        This function assumes that all emails are valid.
    '''

    # Bind server to smpt server, running on port 587 (default), this server is used to send emails thru gmail
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()

    # Make Login
    server.login(my_variable_globals.my_email, my_variable_globals.my_password)  # login into smpt server

    for email_to_send in send_to:
        # Create the body of the message
        msg = MIMEText(message)
        msg['From'] = my_variable_globals.my_email
        msg['To'] = email_to_send
        msg['Subject'] = subject
        # Tell server to send the email
        server.sendmail(my_variable_globals.my_email, email_to_send, msg.as_string())



if __name__ == '__main__':
    my_variable_globals = Global_Storage()  # If needs to change email,and password, only needs to add it as agrument
    email_list = ["a.jplmendes@esmcastilho.pt"]
    import time

    t1 = time.time()
    for _ in range(20):
        email_sender(email_list, "Hello World", "Teste de emails 20*")
    t2 = time.time()

    print("Done in " + str((t2 - t1) / 20) + " seconds")
