from gmail_extractor import (
    console_login, connect_to_imap, get_user_inputs,
    search_emails, fetch_emails, save_to_excel, logout_from_imap
)

# Main script for Gmail Table Extraction
if __name__ == "__main__":
    # User login
    email_user, app_password = console_login()

    # Connect to Gmail
    mail = connect_to_imap(email_user, app_password)

    # Get user inputs for subject and date range
    subject_filter, start_date, end_date = get_user_inputs()

    # Search emails
    email_ids = search_emails(mail, subject_filter, start_date, end_date)

    # Fetch and extract tables from emails
    tables = fetch_emails(mail, email_ids)

    # Save tables to Excel
    save_to_excel(tables)

    # Logout from Gmail
    logout_from_imap(mail)
