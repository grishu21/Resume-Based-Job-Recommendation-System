# Email Notification Utility
import smtplib
from email.message import EmailMessage

def send_email_notification(recipient, company, days_left):
    msg = EmailMessage()
    msg.set_content(f"You have an interview with {company} in {days_left} days.")
    msg['Subject'] = 'Interview Reminder'
    msg['From'] = 'your-email@example.com'
    msg['To'] = recipient

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login('your-email@example.com', 'your-password')
        server.send_message(msg)
