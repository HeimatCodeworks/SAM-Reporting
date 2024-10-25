import msvcrt
from gmail_extractor import (
    connect_to_imap,
    search_emails,
    fetch_emails,
    save_to_excel,
    logout_from_imap
)

def main():
    print("=== SAM Email Reporter ===")
    email_user = input("Enter your Gmail address: ").strip()
    app_password = input("Enter your app-specific password: ").strip()

    print("\n[INFO] Connecting to Gmail...")
    mail = connect_to_imap(email_user, app_password)

    while True:
        # Get user inputs for search
        subject_filter = input("\nEnter the email subject filter: ").strip()
        start_date = input("Enter the start date (e.g., 01-Oct-2024): ").strip()
        end_date = input("Enter the end date (e.g., 24-Oct-2024): ").strip()

        print(f"\n[INFO] Searching for emails with subject containing '{subject_filter}' from {start_date} to {end_date}...")
        email_ids = search_emails(mail, subject_filter, start_date, end_date)

        if email_ids:
            print(f"[INFO] Fetching and processing {len(email_ids)} emails...")
            tables = fetch_emails(mail, email_ids)

            if any(not df.empty for df in tables):
                print("[INFO] Saving results to 'output/output.xlsx'...")
                save_to_excel(tables)
                print("[INFO] Successfully saved data to 'output/output.xlsx'")
            else:
                print("[INFO] No tables found in the emails.")
        else:
            print("[INFO] No matching emails found.")

        # Ask the user if they want to perform another search
        choice = input("\nDo you want to perform another search? (y/n): ").strip().lower()
        if choice != 'y':
            print("\n[INFO] Exiting the program...")
            break

    logout_from_imap(mail)
    print("[INFO] Process completed.")


if __name__ == "__main__":
    main()
    print("\nPress any key to exit...")
    msvcrt.getch()  # Wait for a key press before exiting
