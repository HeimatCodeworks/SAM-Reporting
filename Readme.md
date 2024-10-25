# SAM Email Reporter

## Overview
The `SAM Email Reporter` is a tool that extracts emails from Gmail and compiles the data into a single Excel sheet.

## Setup
1. **Place all files in the same folder:**
   - `SAM Email Reporter.exe`
   - `README.md`
   - An `output` folder (created automatically on the first run).

2. **Run the Program**
   - Double-click `SAM Email Reporter.exe` or run it from the command line.
   - Enter the following information in the console when prompted:
     - Your Gmail address
     - Your app-specific password for Gmail
     - Email subject filter
     - Start date and end date (format: DD-MMM-YYYY)

3. **Output**
   - The extracted data will be saved in `output/output.xlsx`.

## Requirements
- You must create an [app-specific password](https://support.google.com/accounts/answer/185833?hl=en) for your Gmail account to use this tool.

## Troubleshooting
- If you encounter issues, ensure that you enter the correct credentials and parameters in the console.
- The `output/output.xlsx` file will contain the extracted data if successful.

