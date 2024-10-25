# SAM Email Report Extractor v0.1.3

## Overview
The **SAM Email Report Extractor** is a Python-based tool designed to connect to your Gmail account, search for specific emails based on subject filters and date ranges, extract data from tables or "Log:" entries within these emails, and save the results to an Excel file. The tool can handle data formatted with delimiters such as `|`, `;`, and `,`.

## Features
- Connects to your Gmail account using IMAP.
- Searches for emails based on a specified subject and date range.
- Extracts data from tables and "Log:" entries.
- Handles data separated by commas, semicolons, and vertical bars.
- Saves extracted data to an Excel file.
- Provides error handling and logs extraction status.

## Requirements
1. Gmail account with IMAP enabled and an app-specific password.
2. Python 3.7+ installed on your computer.

## Setup
### 1. Enable IMAP in Gmail
   - Log into your Gmail account.
   - Click the **Settings** icon (gear) in the top right corner.
   - Select **See all settings**.
   - Go to the **Forwarding and POP/IMAP** tab.
   - In the **IMAP access** section, select **Enable IMAP**.
   - Click **Save Changes**.

### 2. Generate an App-Specific Password
   - Ensure that **2-Step Verification** is enabled for your Google account.
   - Go to [App Passwords](https://myaccount.google.com/apppasswords).
   - Enter an App Name (exact name doesn't matter).
   - Save the generated password somewhere secure. (This is what you will use to access your email using the tool).

## Usage
1. Run the compiled `.exe` file.
2. You will be prompted to enter:
   - Your Gmail address.
   - Your app-specific password.
   - The subject filter (part or all of the email's subject).
   - Start and end dates (e.g., `01-Oct-2024`).
3. The tool will connect to your Gmail, search for emails matching the criteria, and extract the data.
4. The extracted data is saved to an Excel file named `output.xlsx` in the `output` folder within the program directory. The file will contain:
   - Combined data from all extracted tables.
   - Data from "Log:" entries, separated by the detected delimiter.

## Troubleshooting
- **Login failed:** Double-check your Gmail address, app-specific password, and ensure IMAP is enabled.
- **No emails found:** Verify the subject filter and date range. Ensure emails exist in the specified range.
- **Failed to parse tables:** Some tables may have complex structures that cannot be parsed. Check the output log for more details.
- **No data saved to Excel:** Ensure the correct delimiter is used in "Log:" entries, and that emails contain extractable data.
