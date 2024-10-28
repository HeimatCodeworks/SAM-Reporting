import imaplib
import getpass
import email
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
from datetime import datetime

def console_login():
    print("Welcome to the SAM Email Report Extractor v0.1.4!")
    email_user = input("Enter your Gmail address: ").strip()
    app_password = getpass.getpass("Enter your app-specific password (hidden): ").strip()
    return email_user, app_password

def connect_to_imap(email_user, app_password):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_user, app_password)
        print("[INFO] Successfully connected to Gmail")
        return mail
    except imaplib.IMAP4.error:
        print("[ERROR] Login failed. Please check your credentials.")
        exit(1)

def get_user_inputs():
    subject_filter = input("Enter the email subject filter: ").strip()

    start_date = input("Enter the start date (YYYY-MM-DD): ").strip()
    start_date = format_date(start_date)

    end_date = input("Enter the end date (YYYY-MM-DD): ").strip()
    end_date = format_date(end_date)

    return subject_filter, start_date, end_date

def format_date(date_str):
    """Convert date from 'YYYY-MM-DD' to 'DD-MMM-YYYY' for IMAP, with re-prompting for invalid formats."""
    while True:
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return date_obj.strftime('%d-%b-%Y')
        except ValueError:
            print(f"[ERROR] Invalid date format: {date_str}. Please use 'YYYY-MM-DD'.")
            date_str = input("Re-enter the date (YYYY-MM-DD): ").strip()




def search_emails(mail, subject_filter, start_date, end_date):
    try:
        mail.select('"[Gmail]/All Mail"')

        query = f'(SUBJECT "{subject_filter}" SINCE {start_date} BEFORE {end_date})'
        result, data = mail.search(None, query)

        if result != "OK":
            print(f"[ERROR] IMAP search failed: {result}")
            return []

        email_ids = data[0].split()
        print(f"[INFO] Found {len(email_ids)} emails matching criteria.")
        return email_ids

    except Exception as e:
        print(f"[ERROR] Failed to search emails: {str(e)}")
        return []


def extract_log_entries(email_body, email_timestamp):
    """
    Extracts 'Log:' entries from the email's HTML content and formats them into a DataFrame.
    Adds a new column with the email's timestamp as the first column.
    """
    soup = BeautifulSoup(email_body, 'html.parser')

    logs = re.findall(r'Log:\s*([^\n<]+)', soup.get_text())
    if logs:
        if "," in logs[0]:
            delimiter = ","
        elif ";" in logs[0]:
            delimiter = ";"
        else:
            delimiter = "|"

        headers = logs[0].replace(";", "|").replace(",", "|").split("|")
        data_rows = [log.replace(";", "|").replace(",", "|").split("|") for log in logs[1:]]

        df_logs = pd.DataFrame(data_rows, columns=headers)

        df_logs.insert(0, 'Email Timestamp', email_timestamp)

        if not df_logs.empty:
            print(f"[INFO] Extracted {len(df_logs)} rows from 'Log:' entries.")
            return df_logs

    print("[INFO] No 'Log:' entries found in the email.")
    return pd.DataFrame()


def fetch_emails(mail, email_ids):
    tables = []

    for i, email_id in enumerate(email_ids):
        try:
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    email_timestamp = "Unknown"

                    received_headers = msg.get_all('Received')
                    if received_headers and len(received_headers) >= 3:
                        third_received_header = received_headers[2]
                        print(f"[DEBUG] Third 'Received' header: {third_received_header}")  # Debug line

                        match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', third_received_header)
                        if match:
                            try:
                                email_timestamp = datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S")
                                print(f"[DEBUG] Parsed timestamp: {email_timestamp}")
                            except ValueError as ve:
                                print(f"[ERROR] Failed to parse timestamp: {ve}")
                        else:
                            print("[DEBUG] No valid date-time found in the third 'Received' header.")
                    else:
                        print("[DEBUG] Less than 3 'Received' headers found.")

                    email_body = None
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            if content_type == 'text/html' and 'attachment' not in content_disposition:
                                email_body = part.get_payload(decode=True).decode(
                                    part.get_content_charset() or 'utf-8',
                                    errors='ignore'
                                )
                                break
                    else:
                        if msg.get_content_type() == 'text/html':
                            email_body = msg.get_payload(decode=True).decode(
                                msg.get_content_charset() or 'utf-8',
                                errors='ignore'
                            )

                    if email_body:
                        df_logs = extract_log_entries(email_body, email_timestamp)

                        if isinstance(df_logs, pd.DataFrame) and not df_logs.empty:
                            tables.append(df_logs)

        except Exception as e:
            print(f"[ERROR] Failed to fetch or parse email: {str(e)}")

    return tables

def save_to_excel(tables, filename='output/output.xlsx'):
    if any(isinstance(df, pd.DataFrame) and not df.empty for df in tables):
        try:
            combined_df = pd.concat([df for df in tables if isinstance(df, pd.DataFrame) and not df.empty], ignore_index=True)

            os.makedirs(os.path.dirname(filename), exist_ok=True)
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Log Data', index=False)

            print(f"[INFO] Data saved to {filename}")
            print("[INFO] Data preview:")
            print(combined_df.head())
        except Exception as e:
            print(f"[ERROR] Failed to save data to Excel: {str(e)}")
    else:
        print("[INFO] No valid data found to save.")

def logout_from_imap(mail):
    mail.logout()
    print("[INFO] Logged out from Gmail")

if __name__ == '__main__':
    email_user, app_password = console_login()
    mail = connect_to_imap(email_user, app_password)

    subject_filter, start_date, end_date = get_user_inputs()

    email_ids = search_emails(mail, subject_filter, start_date, end_date)

    extracted_tables = fetch_emails(mail, email_ids)

    save_to_excel(extracted_tables)

    logout_from_imap(mail)