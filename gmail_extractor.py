import imaplib
from tqdm import tqdm
import email
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
from datetime import datetime
from datetime import timedelta
import xlsxwriter
import warnings

warnings.filterwarnings("ignore", category=FutureWarning, module="pandas.core.tools")


def connect_to_imap(email_user, app_password):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_user, app_password)
        print("[INFO] Successfully connected to Gmail")
        return mail
    except imaplib.IMAP4.error:
        print("[ERROR] Login failed. Please check your email and password.")
        return None


def format_date(date_str):
    while True:
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return date_obj.strftime('%d-%b-%Y')
        except ValueError:
            print(f"[ERROR] Invalid date format: {date_str}. Please use 'YYYY-MM-DD'.")
            date_str = input("Re-enter the date (YYYY-MM-DD): ").strip()


from datetime import datetime, timedelta


def search_emails(mail, subject_filter, start_date, end_date):
    try:
        end_date_obj = datetime.strptime(end_date, '%d-%b-%Y') + timedelta(days=1)
        end_date_inclusive = end_date_obj.strftime('%d-%b-%Y')

        mail.select('"[Gmail]/All Mail"')

        query = f'(SUBJECT "{subject_filter}" SINCE {start_date} BEFORE {end_date_inclusive} NOT SUBJECT "Test Run")'
        result, data = mail.search(None, query)

        if result != "OK":
            print(f"[ERROR] IMAP search failed: {result}")
            return []

        email_ids = data[0].split()
        print(f"[INFO] Found {len(email_ids)} emails matching criteria (excluding Test Runs).")
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
        delimiter = "," if "," in logs[0] else ";" if ";" in logs[0] else "|"
        headers = logs[0].replace(";", "|").replace(",", "|").split("|")
        data_rows = [log.replace(";", "|").replace(",", "|").split("|") for log in logs[1:]]
        df_logs = pd.DataFrame(data_rows, columns=headers)
        df_logs.insert(0, 'Email Timestamp', email_timestamp)

        for col in df_logs.columns[1:]:
            df_logs[col] = pd.to_numeric(df_logs[col], errors='ignore')

        if not df_logs.empty:
            print(f"[INFO] Extracted {len(df_logs)} rows from 'Log:' entries.")
            return df_logs

    print("[INFO] No 'Log:' entries found in the email.")
    return pd.DataFrame()


def fetch_emails(mail, email_ids):
    """
    Fetches and processes emails, displaying a single pinned loading bar at the bottom.
    """
    tables = []

    # Initialize tqdm progress bar
    with tqdm(total=len(email_ids), desc="Processing Emails", unit="email", leave=True, position=0) as pbar:
        for i, email_id in enumerate(email_ids):
            try:
                # Fetch the email data
                result, msg_data = mail.fetch(email_id, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        email_timestamp = "Unknown"

                        # Extract timestamp from email's "Date" header
                        date_header = msg.get("Date")
                        if date_header:
                            try:
                                email_timestamp = datetime.strptime(date_header, "%a, %d %b %Y %H:%M:%S %z")
                                email_timestamp = email_timestamp.strftime("%Y-%m-%d %H:%M:%S")
                            except ValueError:
                                print("[WARNING] Failed to parse email timestamp.")

                        # Extract email content
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

                        # Process the email body
                        if email_body:
                            df_logs = extract_log_entries(email_body, email_timestamp)
                            if isinstance(df_logs, pd.DataFrame) and not df_logs.empty:
                                tables.append(df_logs)

                # Update the progress bar
                pbar.update(1)

            except Exception as e:
                print(f"[ERROR] Failed to fetch or parse email: {str(e)}")

    return tables



def save_to_excel(tables, filename='output/output.xlsx'):
    if any(isinstance(df, pd.DataFrame) and not df.empty for df in tables):
        try:
            combined_df = pd.concat([df for df in tables if isinstance(df, pd.DataFrame) and not df.empty],
                                    ignore_index=True)
            date_pattern = re.compile(r"^(0?[1-9]|1[0-2])/(0?[1-9]|[12][0-9]|3[01])/\d{4}$")

            for col in combined_df.columns:
                if combined_df[col].apply(lambda x: bool(date_pattern.match(str(x)))).mean() > 0.5:
                    combined_df[col] = pd.to_datetime(combined_df[col], format='%m/%d/%Y', errors='coerce').dt.date
                else:
                    combined_df[col] = pd.to_numeric(combined_df[col], errors='ignore')

            os.makedirs(os.path.dirname(filename), exist_ok=True)

            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='Log Data', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Log Data']
                date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})

                for col_num, col in enumerate(combined_df.columns):
                    if combined_df[col].dtype == 'object' and combined_df[col].apply(
                            lambda x: bool(date_pattern.match(str(x)))).mean() > 0.5:
                        worksheet.set_column(col_num, col_num, 15, date_format)

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
