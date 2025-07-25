import smtplib
import os
from email.message import EmailMessage

def send_email_with_attachments(
        subject: str,
        body: str, attachments: list,
        recipients: list,
        sender: str = "munteanumihailcosmin@gmail.com",
        password: str = "tjnsoftsbmbyiaiw",
        smtp_server: str = 'smtp.gmail.com',
        smtp_port: int = 587):
    """Send an email with attachments.

    Args:
        sender (str): The email address of the sender.
        password (str): The password or app password for the sender's email account.
        recipients (list): A list of email addresses for the recipients.
        subject (str): The subject of the email.
        body (str): The body content of the email.
        attachments (list): A list of file paths to attach to the email.
        smtp_server (str, optional): The SMTP server to use for sending the email. Defaults to 'smtp.gmail.com'.
        smtp_port (int, optional): The port to use for the SMTP server. Defaults to 587.
    """
    for recipient in recipients:
        msg = EmailMessage()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.set_content(body)

        # Add attachments
        for file_path in attachments:
            if not os.path.isfile(file_path):
                continue
            with open(file_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(file_path)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        # Send email
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                smtp.starttls()
                smtp.login(sender, password)
                smtp.send_message(msg)
            print("Email sent successfully.")
        except Exception as e:
            print(f"Failed to send email: {e}")

# Example usage
# if __name__ == "__main__":
#     send_email_with_attachments(
#         sender_email="munteanumihailcosmin@gmail.com",
#         sender_password="tjnsoftsbmbyiaiw",  # Use app password for Gmail
#         recipient_email="munteanu@althom.de",
#         subject="Test Email with Attachments",
#         body="This is the body of the email.",
#         attachment_paths=[r"D:\IT\Software Engineer - Python & XML.pdf", r"D:\IT\logo.png"]
#     )
