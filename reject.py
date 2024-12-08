import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import re
import logging
import time

# Set up logging
logging.basicConfig(filename='email_log.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Function to validate email format
def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(pattern, str(email)))  # Ensure email is a string

# Function to validate Excel file structure and data
def validate_excel(df, required_columns):
    # Clean and validate column names
    df.columns = df.columns.str.strip()

    # Check for missing required columns
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing columns: {', '.join(missing_columns)}")

    return df

# Function to set up SMTP server
def setup_smtp_server(SMTP_SERVER, SMTP_PORT, SENDER_EMAIL, SENDER_PASSWORD):
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        return server
    except smtplib.SMTPAuthenticationError:
        logging.error("Authentication failed. Check your email and password.")
        raise
    except Exception as e:
        logging.error(f"Error setting up SMTP server: {e}")
        raise

# Load and clean the Excel file
def load_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        logging.error(f"Excel file not found at path: {file_path}")
        raise
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        raise

# Function to send email
def send_email(server, sender_email, recipient_email, subject, body):
    try:
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = subject

        # Attach the body as HTML
        message.attach(MIMEText(body, 'html'))
        
        # Send the email
        server.sendmail(sender_email, recipient_email, message.as_string())
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        return False

# Main function to handle the process
def main():
    # Use the hardcoded Excel file path
    file_path = 'D:/rejectGmailPY/mustafa.xlsx'
    required_columns = ['Email']  # Required column(s)

    try:
        df = load_excel_file(file_path)
        df_cleaned = validate_excel(df, required_columns)
    except Exception as e:
        print(e)
        return

    # Prompt for email credentials
    SENDER_EMAIL = input("Enter your email: ")
    SENDER_PASSWORD = input("Enter your app password: ")

    # Set up the SMTP server
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    try:
        server = setup_smtp_server(SMTP_SERVER, SMTP_PORT, SENDER_EMAIL, SENDER_PASSWORD)
    except Exception as e:
        print(e)
        return

    counter = 0  # Email success counter

    # Iterate through each row in the cleaned Excel sheet
    for index, row in df_cleaned.iterrows():
        recipient_email = row['Email']
        recipient_name = row.get('Name', 'Applicant')  # Use 'Applicant' if Name is missing
        committee = row.get('Committee', 'the team')  # Use 'the team' if Committee is missing
        reject_reason = row.get('Reject Reason', 'We could not proceed with your application.')

        # Check if the email is valid
        if not is_valid_email(recipient_email):
            logging.warning(f"Skipping invalid email: {recipient_name} ({recipient_email})")
            continue

        # Rejection email body with recipient's name, committee, and reject reason
        body = f"""
        <html>
          <body style="background-color: #0f212b; color: #ffffff; font-family: Arial, sans-serif; padding: 20px; position: relative;">
            <img src="https://drive.google.com/uc?id=12JkGCXpaXnsj5EXUGPrgW_w9pYLbJ-LO" alt="Header Image" style="width: 100%; height: auto; margin-bottom: 20px;">
            
            <p style="color:#ffc222;">Dear {recipient_name},</p>
            <p>We appreciate the time and effort you invested in applying to Enactus Menoufia's <strong style="color:#ffc222;">{committee} committee.</strong> After a thorough review of all applications, we <strong style="color:#ffc222;"> regret</strong> to inform you that we are unable to offer you a position at this time.</p>
            <p>The primary reason for this decision is: <strong style="color:#ffc222;">{reject_reason}</strong></p>
            <p>Please know that this outcome does not diminish your skills or the passion you demonstrated in your application. We encourage you to continue seeking opportunities to grow and make an impact in other areas, and we hope to see you apply again in the future.</p>
            <p>Thank you once again for your interest in Enactus Menoufia.</p>
            <p>Best regards,<br>Enactus</p>

            <img src="https://enactusegypt.org/wp-content/uploads/2021/01/Enactus-Full-Color-2.png" alt="Logo" style="position: absolute; bottom: 10px; right: 10px; width: 100px; height: auto;">
          </body>
        </html>
        """

        # Send the email
        if send_email(server, SENDER_EMAIL, recipient_email, f"Application Update - Enactus Menoufia's {committee} Committee", body):
            counter += 1
            logging.info(f"Rejection email sent successfully to {recipient_name} ({recipient_email}). (Total sent: {counter})")
            print(f"Rejection email sent successfully to {recipient_name} ({recipient_email}).")
        else:
            print(f"Failed to send email to {recipient_name} ({recipient_email}).")

        time.sleep(2)  # To avoid hitting email server rate limits

    # Close the SMTP server
    server.quit()

    print(f"Total valid emails processed: {counter}")

if __name__ == "__main__":
    main()
