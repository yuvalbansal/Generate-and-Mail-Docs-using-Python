import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sender_email = 'yuvalbansal2005.yb@gmail.com'
receiver_email = 'yuvalb22@iitk.ac.in'
subject = 'Subject of the Email'
body = 'Body of the email'
password = '' # Generate App Password

if password is None:
    raise ValueError("No email password set in environment variables")

print(f"Retrieved password: {password}")  # Debugging statement

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

filename = '999_Accepted.pdf'
attachment = open(filename, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f"attachment; filename= {filename}")

msg.attach(part)

attachment.close()

try:
    server = smtplib.SMTP('smtp.gmail.com', 587) 
    server.starttls()
    server.login(sender_email, password)
    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)
    print("Email sent successfully")
except Exception as e:
    print(f"Failed to send email: {e}")
finally:
    server.quit()