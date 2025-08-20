Swing Stock Alert Script
This Python script automates the process of identifying swing trading opportunities in stocks listed on the National Stock Exchange (NSE) and sends a daily report to specified recipients via email. The report includes data from a Google Sheet named swing_stock, which is made available as a downloadable Excel file containing only the swing_stock sheet. The script uses Google Sheets API, yfinance for stock data, and SMTP for email delivery.
Table of Contents

Overview
Requirements
Setup Instructions
Usage
Script Components
Ensuring Single-Sheet Download
Error Handling and Fallback
Troubleshooting

Overview
The script performs the following tasks:

Fetches stock symbols from a Google Sheet (Sheet1, Column B, starting from row 4).
Retrieves historical stock data using the yfinance library and calculates technical indicators (e.g., RSI, ADX, Linear Regression).
Updates a Calculation sheet with analysis results.
Waits until 9:25 AM IST to refresh the Google Sheet via a Google Apps Script macro.
Retrieves data from the swing_stock sheet and recipient emails from the credential sheet.
Creates or updates a dedicated Google Spreadsheet (Swing_Stock_Export) containing only the swing_stock sheet.
Sends emails to recipients with an HTML report, including a link to download the swing_stock sheet as an Excel file or a CSV attachment as a fallback.
Refreshes the Google Sheet and logs the process.

The script ensures that the downloadable Excel file contains only the swing_stock sheet, addressing the requirement to prevent access to other sheets (e.g., Sheet1, Calculation, credential).
Requirements
Functional Requirements

Stock Data Analysis:
Fetch stock symbols from Sheet1 (Column B, starting from row 4).
Calculate technical indicators: Latest Close, 200-day Linear Regression, ADX(14), RSI(14), Volume > 1.5 * SMA_Vol20.
Identify stocks meeting all conditions (Close > Regression, Recent Upside Break or Very Close to Break, ADX > 25, RSI > 40, Volume > 1.5 * SMA_Vol20).


Google Sheets Integration:
Authenticate with Google Sheets API using service account credentials.
Update the Calculation sheet with analysis results.
Fetch data from the swing_stock sheet for reporting.
Execute a Google Apps Script macro to refresh the spreadsheet.


Single-Sheet Download:
Provide a downloadable Excel file containing only the swing_stock sheet.
Ensure other sheets (Sheet1, Calculation, credential) are not included in the download.


Email Notifications:
Send HTML emails to recipients listed in the credential sheet (Column D, starting from row 3).
Include a table of swing_stock data, a list of stocks meeting all conditions, and a download link for the swing_stock sheet.
Provide a CSV attachment as a fallback if the Google Sheets export fails.


Scheduling:
Wait until 9:25 AM IST to refresh the Google Sheet before sending emails.


Error Handling:
Handle Google API errors, yfinance connectivity issues, and SMTP failures.
Log all operations for debugging.



Technical Requirements

Python Version: Python 3.8 or higher.
Dependencies: Listed in requirements.txt (see below).
Environment Variables:
GOOGLE_CREDENTIALS: JSON string of Google service account credentials with Google Sheets and Drive API access.
SMTP_USERNAME: Gmail address for sending emails.
SMTP_PASSWORD: Gmail App Password for SMTP authentication.
SWING_STOCK_EXPORT_SHEET_ID: (Optional) ID of the dedicated Swing_Stock_Export spreadsheet.


Google Sheets Setup:
A Google Spreadsheet with ID 1TETP6W6Ee5J7GCMHbSZMLykZlFJOy-qk5JFWN0DYgI4.
Sheets: Sheet1 (stock symbols), Calculation (analysis results), swing_stock (report data), credential (recipient emails).
A Google Apps Script macro (URL in script) to refresh the spreadsheet.


Google Drive Storage: Sufficient storage quota for the dedicated export spreadsheet.

Setup Instructions

Install Python Dependencies:
Run pip install -r requirements.txt to install required libraries.


Set Up Environment Variables:
Create a .env file in the project directory with:GOOGLE_CREDENTIALS='{"type": "service_account", ...}'
SMTP_USERNAME='your_gmail@gmail.com'
SMTP_PASSWORD='your_app_password'
SWING_STOCK_EXPORT_SHEET_ID='optional_spreadsheet_id'


Obtain Google service account credentials from the Google Cloud Console with Sheets and Drive API scopes.
Generate a Gmail App Password for SMTP authentication.


Configure Google Sheets:
Ensure the spreadsheet (1TETP6W6Ee5J7GCMHbSZMLykZlFJOy-qk5JFWN0DYgI4) has the required sheets.
Populate Sheet1 with stock symbols (Column B, row 4 onward).
Populate credential with recipient emails (Column D, row 3 onward).
Deploy a Google Apps Script macro and update macro_url in the script if needed.


Google Drive Storage:
Check the Google Drive account for available storage (delete unused files if needed).
If SWING_STOCK_EXPORT_SHEET_ID is not set, the script creates a new spreadsheet and logs its ID.



Usage

Run the Script:
Execute python stock_swing_trading.py.
The script runs daily, waiting until 9:25 AM IST to refresh the Google Sheet and send emails.


Output:
Logs are printed to the console and saved (based on logging configuration).
Emails are sent to recipients with:
An HTML table of swing_stock data.
A list of stocks meeting all conditions.
A download link for an Excel file containing only the swing_stock sheet or a CSV attachment.




Verify Results:
Check the Calculation sheet for updated stock analysis.
Download the Excel file from the email link to confirm it contains only the swing_stock sheet.
If the CSV fallback is used, verify the attachment contains swing_stock data.



Script Components

Stock Data Fetching:
Uses yfinance to fetch 1-year daily stock data for each symbol (appended with .NS for NSE).
Calculates indicators: Latest Close, 200-day Linear Regression, ADX(14), RSI(14), Volume > 1.5 * SMA_Vol20.
Detects recent upside breaks or closeness to break using peak detection (scipy.signal.argrelextrema).


Google Sheets Operations:
Authenticates using gspread and service account credentials.
Updates the Calculation sheet with analysis results (columns B2:L2 for headers, data from row 4).
Fetches swing_stock data and recipient emails from credential.
Executes a Google Apps Script macro twice with a 4-second pause to refresh data.


Export Spreadsheet Management:
Reuses a dedicated spreadsheet (Swing_Stock_Export) identified by SWING_STOCK_EXPORT_SHEET_ID.
Creates a new spreadsheet if the ID is not set, storing the ID for future runs.
Ensures only one sheet (swing_stock) exists by deleting others and updates it with swing_stock data.
Shares the spreadsheet with “Anyone with the link” for download access.


Email Sending:
Uses smtplib to send HTML emails via Gmail’s SMTP server.
Includes a table of swing_stock data, a list of qualifying stocks, and a download link.
Falls back to a CSV attachment if the export spreadsheet fails.


Scheduling and Timing:
Waits until 9:25 AM IST using datetime and pytz for accurate time zone handling.
Includes a 15-second delay before the final Google Sheet refresh.



Ensuring Single-Sheet Download
To meet the requirement that only the swing_stock sheet is downloadable:

Dedicated Spreadsheet: The script uses a single spreadsheet (Swing_Stock_Export) with only the swing_stock sheet, created or updated each run.
Sheet Management: Deletes any additional sheets in the export spreadsheet to ensure only swing_stock is present.
Export URL: Generates a URL (https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx) that downloads an Excel file containing only the swing_stock sheet.
Verification: The downloaded Excel file is tested to confirm it contains a single sheet with swing_stock data, preventing access to other sheets (Sheet1, Calculation, credential).
Permissions: The spreadsheet is shared with “Anyone with the link” as a reader, ensuring recipients can download without seeing


-Deveshpuri Goswami