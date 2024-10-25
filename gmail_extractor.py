import imaplib
import getpass
import email
from bs4 import BeautifulSoup
import pandas as pd
import os
import re


def console_login():
    print("Welcome to the SAM Email Report Extractor v0.1.3!")
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

    start_date = input("Enter the start date (e.g., 01-Oct-2024): ").strip()
    start_date = format_date(start_date)  # Convert to title case

    end_date = input("Enter the end date (e.g., 24-Oct-2024): ").strip()
    end_date = format_date(end_date)  # Convert to title case

    return subject_filter, start_date, end_date

def format_date(date_str):
    """Ensure the date format has the month in title case."""
    parts = date_str.split('-')
    if len(parts) == 3:
        parts[1] = parts[1].title()  # Convert month to title case
        return '-'.join(parts)
    return date_str



def search_emails(mail, subject_filter, start_date, end_date):
    try:
        mail.select('inbox')
        # Add double quotes around the subject filter for stricter matching
        query = f'(SUBJECT "{subject_filter}" SINCE {start_date} BEFORE {end_date})'
        result, data = mail.search(None, query)
        email_ids = data[0].split()
        print(f"[INFO] Found {len(email_ids)} emails matching criteria.")
        return email_ids
    except Exception as e:
        print(f"[ERROR] Failed to search emails: {str(e)}")
        return []



def extract_log_entries(email_body):
    """
    Extracts 'Log:' entries from the email's HTML content and formats them into a DataFrame.
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

        df_logs = pd.DataFrame(data_rows, columns=headers)
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
                        # If not multipart, check if it's HTML
                        if msg.get_content_type() == 'text/html':
                            email_body = msg.get_payload(decode=True).decode(
                                msg.get_content_charset() or 'utf-8',
                                errors='ignore'
                            )

                    # If HTML body is found, process it
                    if email_body:
                        # Extract 'Log:' entries from the HTML body
                        df_logs = extract_log_entries(email_body)

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
