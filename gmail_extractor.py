import imaplib
import getpass
import email
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
from datetime import datetime
from email.utils import parsedate_to_datetime
from email.utils import parsedate_tz, mktime_tz

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
            # Parse input date in 'YYYY-MM-DD' format
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            # Convert to 'DD-MMM-YYYY' format
            return date_obj.strftime('%d-%b-%Y')
        except ValueError:
            # If the format is incorrect, ask the user to re-enter the date
            print(f"[ERROR] Invalid date format: {date_str}. Please use 'YYYY-MM-DD'.")
            date_str = input("Re-enter the date (YYYY-MM-DD): ").strip()




def search_emails(mail, subject_filter, start_date, end_date):
    try:
        mail.select('"[Gmail]/All Mail"')

        # Ensure that the search query has valid formatting
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
    # Parse the email body using BeautifulSoup
    soup = BeautifulSoup(email_body, 'html.parser')

    # Search for "Log:" entries in the text
    logs = re.findall(r'Log:\s*([^\n<]+)', soup.get_text())
    if logs:
        # Determine the delimiter used in the first log entry
        if "," in logs[0]:
            delimiter = ","
        elif ";" in logs[0]:
            delimiter = ";"
        else:
            delimiter = "|"

        # Extract headers and data rows using the detected delimiter
        headers = logs[0].replace(";", "|").replace(",", "|").split("|")
        data_rows = [log.replace(";", "|").replace(",", "|").split("|") for log in logs[1:]]

        # Create a DataFrame from extracted data
        df_logs = pd.DataFrame(data_rows, columns=headers)

        # Add the email timestamp as the first column
        df_logs.insert(0, 'Email Timestamp', email_timestamp)

        if not df_logs.empty:
            print(f"[INFO] Extracted {len(df_logs)} rows from 'Log:' entries.")
            return df_logs

    print("[INFO] No 'Log:' entries found in the email.")
    return pd.DataFrame()  # Return an empty DataFrame if no logs are found


def fetch_emails(mail, email_ids):
    tables = []

    for i, email_id in enumerate(email_ids):
        try:
            result, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    # Initialize the timestamp as "Unknown"
                    email_timestamp = "Unknown"

                    # Extract all "Received" headers from the email
                    received_headers = msg.get_all('Received')
                    if received_headers and len(received_headers) >= 3:
                        # Access the third "Received" header
                        third_received_header = received_headers[2]
                        print(f"[DEBUG] Third 'Received' header: {third_received_header}")  # Debug line

                        # Extract the date-time part from the third "Received" header
                        match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', third_received_header)
                        if match:
                            try:
                                # Parse the date-time string
                                email_timestamp = datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S")
                                print(f"[DEBUG] Parsed timestamp: {email_timestamp}")  # Debug line
                            except ValueError as ve:
                                print(f"[ERROR] Failed to parse timestamp: {ve}")
                        else:
                            print("[DEBUG] No valid date-time found in the third 'Received' header.")
                    else:
                        print("[DEBUG] Less than 3 'Received' headers found.")

                    # Extract the HTML body from the email
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

                    # If HTML body is found, process it
                    if email_body:
                        # Extract 'Log:' entries from the HTML body, including the timestamp
                        df_logs = extract_log_entries(email_body, email_timestamp)

                        # Ensure df_logs is a DataFrame and not empty before adding it to the tables list
                        if isinstance(df_logs, pd.DataFrame) and not df_logs.empty:
                            tables.append(df_logs)

        except Exception as e:
            print(f"[ERROR] Failed to fetch or parse email: {str(e)}")

    return tables

def save_to_excel(tables, filename='output/output.xlsx'):
    # Check if tables list contains any DataFrames before attempting to concatenate
    if any(isinstance(df, pd.DataFrame) and not df.empty for df in tables):
        try:
            # Combine all DataFrames into one
            combined_df = pd.concat([df for df in tables if isinstance(df, pd.DataFrame) and not df.empty], ignore_index=True)

            os.makedirs(os.path.dirname(filename), exist_ok=True)
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Log Data', index=False)

            print(f"[INFO] Data saved to {filename}")
            print("[INFO] Data preview:")
            print(combined_df.head())  # Display the first few rows
        except Exception as e:
            print(f"[ERROR] Failed to save data to Excel: {str(e)}")
    else:
        print("[INFO] No valid data found to save.")

def logout_from_imap(mail):
    mail.logout()
    print("[INFO] Logged out from Gmail")

# Main function to execute the script
if __name__ == '__main__':
    # Login and connect to IMAP
    email_user, app_password = console_login()
    mail = connect_to_imap(email_user, app_password)

    # Get user inputs for email filtering
    subject_filter, start_date, end_date = get_user_inputs()

    # Search emails based on the provided criteria
    email_ids = search_emails(mail, subject_filter, start_date, end_date)

    # Fetch emails and extract data
    extracted_tables = fetch_emails(mail, email_ids)

    # Save extracted data to Excel
    save_to_excel(extracted_tables)

    # Logout from IMAP
    logout_from_imap(mail)

