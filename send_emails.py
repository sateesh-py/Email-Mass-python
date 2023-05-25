import smtplib
import ssl
import email
import certifi
import pandas as pd

from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path

import os


def smtpmail(mail_reciver,  subject, msg, files):
    user = "YOUR_EMAIL_ADDRESS"
    pwd = "YOUR_EMAIL_PASSWORD"
    try:

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = user
        message["To"] = mail_reciver
        message["Subject"] = subject
        message["Bcc"] = mail_reciver  # Recommended for mass emails

        # Add body to email
        message.attach(MIMEText(msg, "plain"))
        files = files.split(',')
        for file in files:  # add files to the message
            file_path = file
            attachment = MIMEApplication(
                open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition',
                                  'attachment', filename=os.path.basename(file_path))
            message.attach(attachment)

            # Encode file in ASCII characters to send by email
            encoders.encode_base64(attachment)

        # Add header as key/value pair to attachment part

        text = message.as_string()
        with smtplib.SMTP("smtp.gmail.com", 587) as server:  # smtp.web.de
            server.starttls()  # Secure the connection
            server.login(user, pwd)
            server.sendmail(user, mail_reciver, text)
            server.quit()
        print(f"{mail_reciver} --->> sent")

    except Exception as e:
        print("\n::::: -----> Having Problme <----- ::::: ")
        print(f"{mail_reciver} <<>> not sent")
        print('-------------------')
        print(e)
        print('-------------------')
        return e


if __name__ == '__main__':
    # read excel file
    df = pd.read_excel(
        'emailList.xlsx')

    for i in df.index:
        smtpmail(mail_reciver=df['email'][i], subject=df['subject']
                 [i], msg=df['body'][i], files=df['attachements'][i])
