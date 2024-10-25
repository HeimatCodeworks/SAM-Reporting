import imaplib
import getpass
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import pandas as pd

# Console login for Gmail
def console_login():
    print("Welcome to the SAM Report Extractor V1")
    email_user = input("Enter your Gmail address: ").strip()  # Ensure input is captured
    app_password = getpass.getpass("Enter your app-specific password (hidden): ").strip()
    return email_user, app_password

# Connect to Gmail using IMAP
def connect_to_imap(email_user, app_password):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_user, app_password)
        print("[INFO] Successfully connected to Gmail")
        return mail
    except imaplib.IMAP4.error:
        print("[ERROR] Login failed. Please check your credentials.")
        exit(1)

# Get subject filter and date range from user input
def get_user_inputs():
    subject_filter = input("Enter the email subject filter: ")
    start_date = input("Enter the start date (e.g., 01-Oct-2024): ")
    end_date = input("Enter the end date (e.g., 24-Oct-2024): ")
    return subject_filter, start_date, end_date

# Search for emails based on user input
def search_emails(mail, subject_filter, start_date, end_date):
    try:
        mail.select('inbox')
        query = f'(SUBJECT "{subject_filter}" SINCE {start_date} BEFORE {end_date})'
        result, data = mail.search(None, query)
        email_ids = data[0].split()
        print(f"[INFO] Found {len(email_ids)} emails matching criteria.")
        return email_ids
    except Exception as e:
        print(f"[ERROR] Failed to search emails: {str(e)}")
        return []

# Extract tables from email bodies
def extract_tables_from_email(email_body):
    tables = []
    soup = BeautifulSoup(email_body, 'html.parser')
    for table in soup.find_all('table'):
        df = pd.read_html(str(table))[0]
        tables.append(df)
    return tables

# Fetch emails and parse for tables
def fetch_emails(mail, email_ids):
    tables = []
    for email_id in email_ids:
        try:
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == 'text/html':
                                email_body = part.get_payload(decode=True).decode()
                                extracted_tables = extract_tables_from_email(email_body)
                                tables.extend(extracted_tables)
                                print(f"[INFO] Extracted {len(extracted_tables)} tables from an email.")
        except Exception as e:
            print(f"[ERROR] Failed to fetch or parse email: {str(e)}")
    return tables

# Save extracted tables to Excel
def save_to_excel(tables, filename='output/output.xlsx'):
    if tables:
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
            print(f"[INFO] Data saved to {filename}")
        except Exception as e:
            print(f"[ERROR] Failed to save data to Excel: {str(e)}")
    else:
        print("[INFO] No tables found to save.")

# Logout from Gmail
def logout_from_imap(mail):
    mail.logout()
    print("[INFO] Logged out from Gmail")
