import glob
import os
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import base64

def send_email(email, link):
    email_user = 'info@easymarine.se'
    email_password = 'mKTa!46MMBzi3i!'
    
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email
    msg['Subject'] = 'Komplettera din bokning hos oss'

    # Read the HTML content from the file
    with open('email_template.html', 'rb') as f:
        html_content = f.read().decode('utf-8')

    # Replace the placeholder in the HTML content with the actual link
    html_content = html_content.replace('{LINK}', link)

    # Attach the HTML content to the email
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP_SSL('mail.easymarine.se', 465)
        server.login(email_user, email_password)
        text = msg.as_string()
        server.sendmail(email_user, email, text)
        server.quit()
        print("Email sent to:", email)
    except Exception as e:
        print("Error while sending email:", e)

# Get the current working directory
current_folder = os.getcwd()
done_error_folder = os.path.join(current_folder, "done_error")

if not os.path.exists(done_error_folder):
    os.makedirs(done_error_folder)

#read the *_customer_error.txt file and extract the email address and other information correctly 
for customer_error_file in glob.glob(os.path.join(current_folder, "*_customer_error.txt")):
    email = None
    link = "your_link_here"  # Replace this with the actual link you want to send
    with open(customer_error_file, "r") as f:
        lines = f.readlines()
        for line in lines:
            if "Email:" in line:
                email = line.strip().split("Email: ")[1]
            elif "Error: Customer not found in the text file" in line and email:
                send_email(email, link)
                email = None

    # Move the file to the done_error folder
    shutil.move(customer_error_file, os.path.join(done_error_folder, os.path.basename(customer_error_file)))
