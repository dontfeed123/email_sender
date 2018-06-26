#!/usr/bin/python


import smtplib

sender = 'mendes@localhost'

receivers = ['dontfeed123@hotmail.com']

message = """From: mendes@localhost

To: dontfeed123@hotmail.com

Subject: SMTP e-mail test


This is a test e-mail message.
"""

try:

    server = smtplib.SMTP('localhost')
    server.set_debuglevel(1)
    server.sendmail(sender, receivers, message)

    print("Successfully sent email")

except smtplib.SMTPException:

    print("Error: unable to send email")