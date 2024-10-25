# SAM Report Extractor

## Overview
This Python script connects to your Gmail account, searches for emails by subject and date range, extracts tables from the email body, and saves them to an Excel file.

## Setup

### 1. Enable IMAP in Gmail
- Go to **Gmail Settings** → **Forwarding and POP/IMAP**.
- Under "IMAP access," select "Enable IMAP."
- Click **Save Changes**.

### 2. Generate an App-Specific Password
- Go to [Google Account Security Settings](https://myaccount.google.com/security).
- Under "Signing in to Google," select **App passwords**.
- Generate an app-specific password for Gmail.

### 3. Set Up the Virtual Environment
- If you haven’t already created a virtual environment:
  ```bash
  python -m venv .venv
