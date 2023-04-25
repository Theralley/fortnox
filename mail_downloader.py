import imaplib
import email
import os
import time

def download_attachments(mail, email_id, output_dir):
    allowed_extensions = ['.csv', '.xlsx']
    #commented_extensions = ['.ics']  # Add this extension to allowed_extensions to uncomment

    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        file_name = part.get_filename()
        if bool(file_name):
            file_extension = os.path.splitext(file_name)[1]

            if file_extension in allowed_extensions or file_extension in commented_extensions:
                file_path = os.path.join(output_dir, file_name)

                if not os.path.exists(file_path):
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print("Attachment downloaded")
                else:
                    existing_file_size = os.path.getsize(file_path)
                    new_attachment_size = len(part.get_payload(decode=True))

                    if existing_file_size != new_attachment_size:
                        file_name_without_ext, file_ext = os.path.splitext(file_name)
                        i = 1
                        while os.path.exists(os.path.join(output_dir, f"{file_name_without_ext}_{i}{file_ext}")):
                            i += 1

                        new_file_path = os.path.join(output_dir, f"{file_name_without_ext}_{i}{file_ext}")
                        with open(new_file_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        print("Attachment with different size downloaded")
                    else:
                        print(f"Attachment '{file_name}' already exists with the same size, skipping download")

def main():
    email_user = 'info@easymarine.se'
    email_password = 'mKTa!46MMBzi3i!'  # Replace with your email account's password
    output_dir = 'downloaded_attachments'

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    while True:
        print("Listening...")
        try:
            mail = imaplib.IMAP4_SSL('mail.easymarine.se', 993)
            mail.login(email_user, email_password)
            mail.select('inbox')

            _, data = mail.search(None, 'UNSEEN')
            email_ids = data[0].split()

            for email_id in email_ids:
                _, data = mail.fetch(email_id, '(RFC822)')
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                download_attachments(msg, email_id, output_dir)

            mail.logout()
        except Exception as e:
            print(f'Error: {e}')

        time.sleep(5)

if __name__ == '__main__':
    main()
