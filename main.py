import msvcrt
from gmail_extractor import (
    connect_to_imap,
    search_emails,
    fetch_emails,
    save_to_excel,
    logout_from_imap,
    format_date
)
import os

def main():
    print("=== SAM Email Reporter ===")

    # Loop for login with retry option
    while True:
        email_user = input("Enter your Gmail address: ").strip()
        app_password = input("Enter your app-specific password: ").strip()

        print("\n[INFO] Connecting to Gmail...")
        mail = connect_to_imap(email_user, app_password)

        if mail:
            break  # Successful connection
        else:
            retry = input("[INFO] Would you like to try again? (y/n): ").strip().lower()
            if retry != 'y':
                print("[INFO] Exiting the program...")
                return

    # Search emails and save results in loop
    while True:
        subject_filter = input("\nEnter the email subject filter: ").strip()

        start_date = input("Enter the start date (YYYY-MM-DD): ").strip()
        start_date = format_date(start_date)

        end_date = input("Enter the end date (YYYY-MM-DD): ").strip()
        end_date = format_date(end_date)

        print(
            f"\n[INFO] Searching for emails with subject containing '{subject_filter}' from {start_date} to {end_date}..."
        )
        email_ids = search_emails(mail, subject_filter, start_date, end_date)

        if email_ids:
            print(f"[INFO] Fetching and processing {len(email_ids)} emails...")
            tables = fetch_emails(mail, email_ids)

            if any(not df.empty for df in tables):
                safe_subject_filter = subject_filter.replace(" ", "_")
                safe_filename = f"{safe_subject_filter}_{start_date}_{end_date}.xlsx"
                output_path = os.path.join("output", safe_filename)

                print(f"[INFO] Saving results to '{output_path}'...")
                save_to_excel(tables, filename=output_path)
                print(f"[INFO] Successfully saved data to '{output_path}'")
            else:
                print("[INFO] No tables found in the emails.")
        else:
            print("[INFO] No matching emails found.")

        choice = input("\nDo you want to perform another search? (y/n): ").strip().lower()
        if choice != 'y':
            print("\n[INFO] Exiting the program...")
            break

    # Logout from email
    logout_from_imap(mail)
    print("[INFO] Process completed.")

if __name__ == "__main__":
    main()
    print("\nPress any key to exit...")
    msvcrt.getch()
