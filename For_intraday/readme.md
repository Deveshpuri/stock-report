Stock Report Email Generator
Created by: Devesh Puri GoswamiDate: August 19, 2025
This project automates the generation and emailing of stock reports using data from Google Sheets. It fetches stock data at specific times (9:15 AM, 9:20 AM, and 9:23 AM IST), updates a Google Sheet with stock prices and calculations, and sends formatted HTML emails with the data to recipients. A second email round at 10:55 AM IST includes additional orb_dhan data.
Features

herethe link of sheet for vetter understanding
https://docs.google.com/spreadsheets/d/19JfiLy5VPfLLMMMl_k4TYF6ylXYkKQdgZd5t0boi7TA/edit?usp=sharing

Executes Google Apps Script macros to refresh Google Sheet data at 9:15 AM, 9:20 AM, and 9:23 AM IST.
Fetches stock data from high_break_trade and onetime_five_open sheets.
Updates the compare and orb_dhan sheets with stock names and price formulas.
Sends two rounds of emails with formatted tables containing stock data from high_break_trade, onetime_five_open, and orb_dhan sheets.
Uses environment variables for secure credential management.
Includes retry logic for macro execution and parallel email sending with ThreadPoolExecutor.

Google Sheet Structure and Updates
The script interacts with a Google Sheet identified by the ID 1ZYa5e92hmTc27KYSLKJyx_zvYydR5b1jMk7dYIqFsss. The sheet contains the following worksheets, which are modified as part of the process:

high_break_trade: Stores initial stock data refreshed by the macro. Used to extract stock names for the compare sheet.
onetime_five_open: Contains stock data with a dynamic last row (based on column B). Used for email reports.
compare: Updated with stock names and formulas:
At 9:20 AM IST: Clears ranges B4:E and G4:J, writes stock names from high_break_trade to B4 downward, and adds formulas in C4:E (e.g., =CONCAT("NSE:",B4), =GOOGLEFINANCE(C4,"priceopen"), =GOOGLEFINANCE(C4,"price")).
At 9:23 AM IST: Writes updated stock names from high_break_trade to G4 downward and adds formulas in H4:J.


orb_dhan: Updated with stock names and formulas:
Clears ranges B4:F and H4:L.
Writes 9:20 AM stock names from compare column L to B4 downward, with formulas in C4:E (e.g., =CONCAT("NSE:",B4), =GOOGLEFINANCE(C4,"priceopen"), =GOOGLEFINANCE(C4,"price")) and percentage change in F4 (=(E4-D4)/E4*100).
Writes 9:23 AM stock names from compare column M to H4 downward, with formulas in I4:K and percentage change in L4.


credential: Stores recipient email addresses in column D (starting from row 3), used to populate the email list.

Steps in Google Sheet

9:15 AM IST: Executes the first macro to refresh high_break_trade and onetime_five_open.
9:20 AM IST: Executes the second macro, clears and updates compare sheet with 9:20 AM data, and sends the first email round.
9:23 AM IST: Executes the third macro, updates compare with 9:23 AM data, clears and updates orb_dhan sheet, and prepares for the second email.
10:55 AM IST: Sends the second email round with updated orb_dhan data.

Prerequisites

Python 3.8+
Required Python packages: requests, gspread, google-auth, python-dotenv, smtplib
Google Cloud service account credentials (stored in .env)
Gmail account with an App Password for SMTP
Google Sheets API enabled
Google Apps Script macro deployed with the provided URL

Setup

Clone the repository:git clone https://github.com/Deveshpuri/stock-report.git
cd stock-report


Install dependencies:pip install -r requirements.txt


Create a .env file in the root directory or For_intraday subdirectory with the following:SMTP_USERNAME=your_email@gmail.com
SMTP_PASSWORD=your_app_password
GOOGLE_CREDENTIALS='{"type": "service_account", ...}'  # Replace with your JSON credentials


Replace your_email@gmail.com and your_app_password with your Gmail address and App Password.
Replace the GOOGLE_CREDENTIALS value with the JSON content from your Google Cloud service account key file.


Ensure the Google Sheet ID (1ZYa5e92hmTc27KYSLKJyx_zvYydR5b1jMk7dYIqFsss) and macro URL (https://script.google.com/macros/s/AKfycbxGVZ_OJviGVLR4_-JXEoJjTwXobvm6wX0uV_HXOVLwNQkoY49JWphXhrOrcI0ySiYw/exec) in stock_report_email.py match your setup.
Configure the credential sheet with recipient email addresses in column D (starting from row 3).

Usage
Run the script:
python stock_report_email.py

The script will:

Wait until 9:15 AM IST to execute the first macro and refresh the Google Sheet.
Refresh the Google Sheet again at 9:20 AM IST.
Update the compare sheet with stock data and formulas at 9:20 AM IST.
Send the first round of emails with high_break_trade and onetime_five_open data at 9:20 AM IST.
Wait until 9:23 AM IST for a third refresh.
Update the orb_dhan sheet with stock data and formulas at 9:23 AM IST.
Send a second round of emails with high_break_trade, onetime_five_open, and orb_dhan data at 10:55 AM IST.

Running on Your Own System

Install Python: Download and install Python 3.8+ from python.org if not already installed. Verify with:python --version


Set Up Virtual Environment (Optional):
Create a virtual environment:python -m venv venv


Activate it:
Windows: venv\Scripts\activate


Install dependencies within the virtual environment:pip install requests gspread google-auth python-dotenv




Configure .env: Place the .env file in the same directory as stock_report_email.py (e.g., For_intraday or root) with your credentials.
Adjust for Testing: Since it’s 2:16 PM IST on August 19, 2025, modify stock_report_email.py temporarily:
Open the file and change target_time_ist_1 to a future time, e.g.:target_time_ist_1 = current_time_ist.replace(hour=14, minute=30, second=0, microsecond=0)


Adjust target_time_ist_2 (add 5 minutes) and the final wait (add 3 minutes, or comment out time.sleep(180) for immediate execution).


Run the Script:
Navigate to the project directory:cd C:\Users\DEVESHPURI GOSWAMI\Desktop\Projects\stock-report1


Execute:python stock_report_email.py




Restore Original Times: After testing, revert target_time_ist_1, target_time_ist_2, and the 3-minute wait to 9:15 AM, 9:20 AM, 9:23 AM, and 10:55 AM IST.

Expected Terminal Output
The script provides real-time feedback. Example output during a test run (adjusted for 2:30 PM IST):
Starting to fetch and send stock report emails...
Waiting 840 seconds until 2:30 PM IST to refresh Google Sheet...
Macro executed successfully on attempt 1
First Google Sheet refresh completed at 2:30 PM IST.
Waiting 300 seconds until 2:35 PM IST for second refresh...
Macro executed successfully on attempt 1
Second Google Sheet refresh completed at 2:35 PM IST.
Cleared existing data from B4:E and G4:J in compare sheet at 2:35 PM IST.
Wrote 10 stocks to column B at 2:35 PM IST.
Email sent to user1@example.com
Email sent to user2@example.com
Sending emails at 2:35 PM IST...
Sent 2 emails successfully.

Waiting until 2:38 PM IST to refresh Google Sheet again...

Macro executed successfully on attempt 1
Third Google Sheet refresh completed at 2:38 PM IST.
Wrote 10 stocks to column G at 2:38 PM IST.
Cleared existing data from B4:B and H4:H in orb_dhan sheet.
Wrote 10 stocks to column B in orb_dhan sheet.
Wrote 10 stocks to column H in orb_dhan sheet.
Updated formulas in orb_dhan sheet for 9:20 AM and 9:23 AM data.
Sending second emails with high_break_trade, onetime_five_open, and orb_dhan data at 2:55 PM IST...
Email sent to user1@example.com
Email sent to user2@example.com
Sent 2 second emails successfully with high_break_trade, onetime_five_open, and orb_dhan data.

Sent 2 emails successfully. Process completed in 1500.34 seconds.


Output varies based on the number of recipients, macro success, and email delivery.

Notes

The script uses IST (UTC+5:30) time zone. For production, keep the original times.
The .env file is excluded from version control via .gitignore for security.
Ensure the Google Sheets API and Gmail SMTP settings are correctly configured.
The script uses ThreadPoolExecutor for parallel email sending with a maximum of 15 workers.

License
MIT License