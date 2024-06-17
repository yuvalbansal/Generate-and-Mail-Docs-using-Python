import os
import csv
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

email_address = 'wceam2024@gmail.com'
email_password = '' # Generate App Password
smtp_server = 'smtp.gmail.com'
smtp_port = 587

def send_email(to_address, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = to_address
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))
    attachment = open(attachment_path, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {attachment_path}")

    msg.attach(part)

    attachment.close()

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_address, email_password)
        text = msg.as_string()
        server.sendmail(email_address, to_address, text)
        print(f'Sent {to_address}')
    except Exception as e:
        print(f"Failed {e}")
    finally:
        server.quit()

with open('Review Abstract Submission.csv', mode ='r')as file:
    csvFile = csv.reader(file)
    for lines in csvFile:
        Number = lines[0]
        Email = lines[2]
        Author = lines[3]
        Title = lines[1]
        Comments = lines[5]
        document_name = f'{Number}.pdf'
        if lines[4] == "Accepted":
            Subject = "Acceptance of Abstract Submission"
            Body = f'Dear {Author},\n\nPlease find the document enclosed.\n\nBest regards,\nWCEAM Team'
            Path = f'Letters_PDF\{Number}_Accepted.pdf'
            send_email(Email, Subject, Body, Path)
            print(f'Email sent to abstract number {Number}')
        elif lines[4] == "Rejected":  
            Subject = "Rejection of Abstract Submission"
            Body = f'Dear {Author},\n\nPlease find the document enclosed.\n\nBest regards,\nWCEAM Team'
            Path = f'Letters_PDF\{Number}_Rejected.pdf'
            send_email(Email, Subject, Body, Path)
            print(f'Email sent to abstract number {Number}')

print('All emails have been sent.')