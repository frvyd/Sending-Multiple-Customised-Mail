import pandas as pd

# If we are going to create a pdf ourselves, we must add this library
#from fpdf import FPDF

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import smtplib

# Excel, pdf and code file must be in the same folder

data = pd.read_excel("receivers.xlsx")

# Enter the names of the columns where you get the mail and names of the participants
mails= data["mail"]
names = data["name surname"]

def send_email(mail, name):
    msg = MIMEMultipart()

    # Enter the sender's email address and password
    msg["From"] = "sender@gmail.com"
    msg["To"] = mail
    msg["Subject"] = "Certificate of participation in the ... program"

    body = f"Hello {name},\n\nThank you for your participation in the ... program. Your participation certificate is attached.\n\nSincerely,\n\nOrganization Team"

    msg.attach(MIMEText(body, "plain"))
    try:
        with open(f"{name}.pdf", "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")
            attachment.add_header("Content-Disposition", f"attachment; filename={name}.pdf")

        msg.attach(attachment)


        # Enter the settings of the SMTP server
        '''
        1. Go to https://myaccount.google.com/intro in a browser.
        2. Sign in to your Google Account.
        3. Go to the Security tab.
        4. Select App Passwords.
        5. Click Choose Password.
        6. In the Application option, select Other.
        7. Type "Program Name" in the Name field. Click Generate.
        8. Enter the generated password in the code instead of "password".
        '''
        # Remember, your account must have two-step verification turned on. Otherwise you won't be able to perform this step.
        '''
        if your email provider is Gmail:
        SMTP server address: smtp.gmail.com
        SMTP port 587
        Username: your email address
        Password: The custom application password you created for the SMTP server, not the password you use in the Gmail app
        '''
        
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("sender@gmail.com", "password")
        server.sendmail(msg["From"], msg["To"], msg.as_string())
        server.quit()

    except FileNotFoundError:
        print(f"Error: no PDF file found for {name}.")

for mail, name in zip(mails, names):
    send_email(mail, name)
