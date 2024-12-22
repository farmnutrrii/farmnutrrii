from apscheduler.schedulers.background import BackgroundScheduler
from django.core.mail import EmailMessage
from .utils import create_combined_excel  # Import your utility function to create the Excel file
from datetime import datetime
import os

def send_combined_report():
    # Create the Excel report for today's date
    stock_date = datetime.today().date()
    excel_file = create_combined_excel(stock_date)

    # Send email with the report attached
    subject = "Combined Report for " + str(stock_date)
    body = "Please find attached the combined report for the date: " + str(stock_date)
    from_email = 'farmnutrriichicken1@gmail.com'
    recipient_list = ['cnkumar66@gmail.com','farmnutrriichicken1@gmail.com']  # Replace with the actual recipient email

    email = EmailMessage(subject, body, from_email, recipient_list)
    email.attach_file(excel_file)
    
    try:
        email.send()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

    # Optionally, delete the temporary Excel file after sending the email
    if os.path.exists(excel_file):
        os.remove(excel_file)

def start_scheduler():
    scheduler = BackgroundScheduler()
    scheduler.add_job(send_combined_report, 'cron', hour=23, minute=30)  # Schedule to run at 10:00 AM every day
    scheduler.start()
