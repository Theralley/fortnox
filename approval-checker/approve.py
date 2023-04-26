import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import time


def send_email_with_link(receiver_email):
    sender_email = "info@easymarine.se"
    password = "mKTa!46MMBzi3i!"

    message = MIMEMultipart("alternative")
    message["Subject"] = "Please confirm"
    message["From"] = sender_email
    message["To"] = receiver_email

    unique_link = "http://127.0.0.1:5000/unique-link"
    text = f"Please click the following link to confirm: {unique_link}"
    html = f'<p>Please click the following link to confirm: <a href="{unique_link}">Link</a></p>'

    text_part = MIMEText(text, "plain")
    html_part = MIMEText(html, "html")

    message.attach(text_part)
    message.attach(html_part)

    with smtplib.SMTP_SSL("mail.easymarine.se", 465) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    print("Email sent with unique link")

def check_link_status():
    response = requests.get("http://127.0.0.1:5000/check-status")
    link_status = response.json()

    if link_status["clicked"]:
        print("The link has been clicked!")
    else:
        print("The link has not been clicked yet.")

receiver_email = "rasmus_hamren@msn.com"
send_email_with_link(receiver_email)

while True:
    check_link_status()
    time.sleep(5)
