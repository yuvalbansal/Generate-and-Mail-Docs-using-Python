import os
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email credentials
email_address = 'yuvalbansal2005.yb@gmail.com'
email_password = ''  # App-specific password
smtp_server = 'smtp.gmail.com'
smtp_port = 587

def send_email(to_address, subject, body):
    # Create the email headers and body
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = to_address
    msg['Subject'] = subject

    # Attach the message body
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Connect to the SMTP server and send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade to secure connection
            server.login(email_address, email_password)
            server.sendmail(email_address, to_address, msg.as_string())
        print(f'Successfully sent email to {to_address}')
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")

# Reading the CSV file
with open('Final Abstract List.csv', mode='r') as file:
    csvFile = csv.reader(file)
    for lines in csvFile:
        Number = lines[0]
        Email = lines[4]
        Author = lines[2]
        Title = lines[1]
        Subject = "Regarding WCEAM 2024"
        
        # Skip header row or rows without a valid abstract number
        if Number.isdigit():
            if lines[6].strip().lower() == "accepted":
                Body = (f'Dear {Author},\n\n'
                        f'This is a system-generated email. Your registration for the conference is complete, and your submission titled "{Title}" '
                        f'(Abstract Number: {Number}) has been accepted. Please submit your full papers on the EquinOCS platform as mentioned on the website: '
                        f'https://wceam2024.com/call-for-papers/ if you have not done so yet.\n\n'
                        f'Please ignore if you have already submitted your full paper.\n\n'
                        f'Thank you!\n\nBest regards,\nYuval Bansal\nWCEAM Team (IIT Kanpur)')
                send_email(Email, Subject, Body)
            elif lines[6].strip().lower() == "rejected":
                Body = (f'Dear {Author},\n\n'
                        f'This is a system-generated email. Please ignore if you have opted out of the WCEAM conference.\n\n'
                        f'Your submission titled "{Title}" (Abstract Number: {Number}) has been cancelled. If you think there is some mistake, '
                        f'please ensure that you have registered for the conference and reply to this email.\n\n'
                        f'Thank you!\n\nBest regards,\nYuval Bansal\nWCEAM Team (IIT Kanpur)')
                send_email(Email, Subject, Body)

print('All emails have been sent.')
