from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import smtplib
import os
import pandas as pd
import time

file_path = "list.xlsx"
df = pd.read_excel(file_path)

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
YOUR_NAME = os.getenv("YOUR_NAME")
CC_EMAIL = os.getenv("CC_EMAIL")

context = ssl.create_default_context()

with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls(context=context)
    server.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)

    for index, row in df.iterrows():
        recipient_name = row["Name"]
        recipient_email = row["Email"]

        print(f"Preparing email for {recipient_name} ({recipient_email})")

        msg = MIMEMultipart()
        msg["From"] = OUTLOOK_EMAIL
        msg["To"] = recipient_email
        msg["Cc"] = CC_EMAIL
        msg["Subject"] = 'Testing automatisation with python'

        # Email body (personalized)
        email_body = f"""Dear {recipient_name},

This is automated message sent with Python.

Sincerely,
{YOUR_NAME}

"""

        msg.attach(MIMEText(email_body, "plain"))

        server.sendmail(
            OUTLOOK_EMAIL, [recipient_email, CC_EMAIL], msg.as_string())
        print(f"✅ Email sent to {recipient_name} ({recipient_email})")
        time.sleep(3)

print("All emails sent successfully!")
