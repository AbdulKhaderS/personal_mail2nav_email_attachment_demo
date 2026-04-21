# Mail2Nav – Email Attachment Automation (Demo)

This is a **demo** version of my real Mail2Nav tool.  
It reads Outlook emails related to item pricing/creation, saves attached Excel files, and helps me reply after processing them in Navision.  
All email addresses here are example/demo addresses (no real company data).

## What this tool does

- Connects to Outlook and scans a specific folder for new emails.
- Filters emails by subject / type (for example NEW ITEMS, CHANGE PRICE, etc.).
- Saves Excel attachments to a local folder with clean file names.
- Tracks each email as a "job" (sender, subject, file, type).
- Provides a Tkinter control panel:
  - Process new emails.
  - Open a reply window for completed jobs (Done — Send Reply).
  - - Fits into an ERP workflow (Navision) for new item creation and price changes.

## Tech stack

- Python
- Outlook automation (`win32com.client`, `pythoncom`)
- GUI with `tkinter`
- Standard libraries: `os`, `json`, `datetime`, `threading`, etc.
- Runs on Windows with Outlook desktop installed (COM automation)

## Important note

- This repository is for **personal learning and portfolio**.
- All email addresses use `@example.com` and do not belong to any real company.
- The private production version at work uses real company emails and secure configuration files.

## How to use (demo)

1. Clone this repository.
2. Open `email_attachment_demo.py` in PyCharm or your editor.
3. Update these placeholders to match your environment:
   - Outlook folder name.
   - Local folder path for saving Excel files.
   - Test email addresses (if you want to send to your own test account).
4. Run the script.
5. Use the Mail2Nav control panel to:
   - Click **Process New Emails** to load new jobs.
   - After you process them in your ERP / Navision, click **Done — Send Reply** to open the Outlook reply window.

## Possible improvements

- Move configuration (folders, subjects, recipients) to a JSON or `.env` file.
- Add logging to a text file for all processed messages.
- Add error handling for missing attachments or invalid Excel files.
- Package it as an `.exe` for easier use by non-technical users.
