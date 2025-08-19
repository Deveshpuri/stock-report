import json
import requests
import time



import smtplib
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from datetime import timezone
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Record start time
start_time = time.time()
print("Starting to fetch and send stock report emails...")

# Function to execute macro with retry logic
def execute_macro_with_retry(url, retries=3, delay=5):
    for attempt in range(retries):
        try:
            response = requests.get(url)
            response.raise_for_status()
            print(f"Macro executed successfully on attempt {attempt + 1}")
            return True
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < retries - 1:
                time.sleep(delay)
    print("All retry attempts failed for macro execution")
    return False

# Step 1: Wait until 9:15 AM IST to execute the first Google Apps Script macro
current_time_ist = datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)
target_time_ist_1 = current_time_ist.replace(hour=9, minute=15, second=0, microsecond=0)
if current_time_ist < target_time_ist_1:
    wait_seconds = (target_time_ist_1 - current_time_ist).total_seconds()
    print(f"Waiting {wait_seconds:.0f} seconds until 9:15 AM IST to refresh Google Sheet...")
    time.sleep(wait_seconds)

# Execute the Google Apps Script macro (first refresh)
macro_url = "https://script.google.com/macros/s/AKfycbxGVZ_OJviGVLR4_-JXEoJjTwXobvm6wX0uV_HXOVLwNQkoY49JWphXhrOrcI0ySiYw/exec"
execute_macro_with_retry(macro_url)
print("First Google Sheet refresh completed at 9:15 AM IST.")

# Step 2: Wait 5 minutes (until 9:20 AM IST) for the second refresh
target_time_ist_2 = target_time_ist_1 + timedelta(minutes=5)
wait_seconds = (target_time_ist_2 - (datetime.now(timezone.utc) + timedelta(hours=5, minutes=30))).total_seconds()
if wait_seconds > 0:
    print(f"Waiting {wait_seconds:.0f} seconds until 9:20 AM IST for second refresh...")
    time.sleep(wait_seconds)

# Execute the Google Apps Script macro again (second refresh)
execute_macro_with_retry(macro_url)
print("Second Google Sheet refresh completed at 9:20 AM IST.")

# Step 3: Access the Google Sheet data
try:
    credentials_info = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
    creds = Credentials.from_service_account_info(credentials_info, scopes=[
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets'
    ])
    client = gspread.authorize(creds)
except Exception as e:
    print(f"Error loading credentials: {e}")
    exit(1)

sheet_id = '1ZYa5e92hmTc27KYSLKJyx_zvYydR5b1jMk7dYIqFsss'
workbook = client.open_by_key(sheet_id)

# Fetch data from both sheets
high_break_trade_sheet = workbook.worksheet('high_break_trade')
onetime_five_open_sheet = workbook.worksheet('onetime_five_open')

gid_high_break = high_break_trade_sheet.id
data_high_break = high_break_trade_sheet.get_all_values()

column_b_values = onetime_five_open_sheet.col_values(2)
last_row_b = max(i for i, val in enumerate(column_b_values, 1) if val) if any(column_b_values) else 3
data_onetime_five = onetime_five_open_sheet.get_all_values()[:last_row_b + 1]

# Fetch emails from "credential" sheet
credential_sheet = workbook.worksheet('credential')
email_column = credential_sheet.col_values(4)
recipients = [email for email in email_column[2:] if email]

# Write stock names and formulas to "compare" sheet at 9:20 AM
compare_sheet = workbook.worksheet('compare')
stock_names_920 = [row[1] for row in data_high_break[3:] if row[1]]  # Fetch stocks from column B, skipping headers
num_stocks_920 = len(stock_names_920)

# Flush (clear) existing data in ranges B4:E and G4:J
compare_sheet.batch_clear(['B4:E', 'G4:J'])
print("Cleared existing data from B4:E and G4:J in compare sheet at 9:20 AM IST.")

# Write new stock names to column B
compare_sheet.update(range_name='B4', values=[[stock] for stock in stock_names_920])  # Write to B4 downward
print(f"Wrote {num_stocks_920} stocks to column B at 9:20 AM IST.")

# Generate and write formulas row-by-row for 9:20 AM data
c_formulas = [f'=CONCAT("NSE:",B{row+4})' for row in range(num_stocks_920)]
d_formulas = [f'=GOOGLEFINANCE(C{row+4},"priceopen")' for row in range(num_stocks_920)]
e_formulas = [f'=GOOGLEFINANCE(C{row+4},"price")' for row in range(num_stocks_920)]
compare_sheet.update(range_name='C4:C' + str(3 + num_stocks_920), values=[[formula] for formula in c_formulas], value_input_option='USER_ENTERED')
compare_sheet.update(range_name='D4:D' + str(3 + num_stocks_920), values=[[formula] for formula in d_formulas], value_input_option='USER_ENTERED')
compare_sheet.update(range_name='E4:E' + str(3 + num_stocks_920), values=[[formula] for formula in e_formulas], value_input_option='USER_ENTERED')

time.sleep(10)  # Increased sleep to allow GOOGLEFINANCE to fetch data

# Build HTML table/grid
high_break_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_high_break[0])
high_break_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #dee2e6; background-color: {"#ffffff" if i % 2 == 0 else "#f8f9fa"}; color: #495057;">{val}</td>' for val in row) + '</tr>'
    for i, row in enumerate(data_high_break[1:])
)
high_break_stocks = ''.join(
    f'<li style="padding: 10px; background-color: #fff; margin: 8px 0; border-left: 4px solid #dc3545; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">{stock}</li>'
    for stock in [row[1] for row in data_high_break[3:] if row[1]]
)

onetime_five_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_onetime_five[0])
onetime_five_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #dee2e6; background-color: {"#ffffff" if i % 2 == 0 else "#f8f9fa"}; color: #495057;">{val}</td>' for val in row) + '</tr>'
    for i, row in enumerate(data_onetime_five[1:last_row_b + 1])
)

export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_high_break}"
html_body_template = """
<html>
  <head>
    <style>
      body {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #333; line-height: 1.6; background-color: #f4f4f4; margin: 0; padding: 20px; }}
      .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
      h3 {{ color: #2c3e50; margin-top: 30px; border-bottom: 2px solid #007bff; padding-bottom: 8px; }}
      p {{ margin: 10px 0; }}
      table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
      ul {{ list-style: none; padding: 0; }}
      .download-btn {{ display: inline-block; padding: 12px 24px; background-color: #007bff; color: #fff; text-decoration: none; border-radius: 5px; font-weight: 500; transition: background-color 0.3s, transform 0.2s; }}
      .download-btn:hover {{ background-color: #0056b3; transform: translateY(-2px); }}
      .footer {{ margin-top: 30px; text-align: center; color: #6c757d; font-size: 0.9em; }}
      @media (max-width: 600px) {{ .container {{ padding: 15px; }} table {{ font-size: 0.9em; }} }}
    </style>
  </head>
  <body>
    <div class="container">
      <p>Dear {recipient_name},</p>
      <p>Generated on: {generated_time}</p>
      <h3>High Break Trade Data</h3>
      <table>{high_break_headers}{high_break_rows}</table>
      <h3>Onetime Five Open Data</h3>
      <table>{onetime_five_headers}{onetime_five_rows}</table>
      <h3>Stock Names (Column B)</h3>
      <ul>{high_break_stocks}</ul>
      <p><a href="{export_url}" class="download-btn">Download Excel</a></p>
      <div class="footer">This is an automated report from HighBreak Alert. For support, contact our team.</div>
    </div>
  </body>
</html>
"""

# SMTP setup
smtp_server = 'smtp.gmail.com'
smtp_port = 587
username = os.getenv('SMTP_USERNAME')
password = os.getenv('SMTP_PASSWORD')
sender = '"HighBreak Alert" <' + os.getenv('SMTP_USERNAME') + '>'
generated_time = (datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)).strftime('%Y-%m-%d %H:%M:%S') + ' IST'

# Send email function
def send_email(recipient_email, html_body):
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as s:
            s.starttls()
            s.login(username, password)
            msg = MIMEText(html_body, 'html')
            msg['Subject'] = 'Stock Data Report from high_break_trade and onetime_five_open'
            msg['From'] = sender
            msg['To'] = recipient_email
            s.sendmail(sender, [recipient_email], msg.as_string())
            print(f"Email sent to {recipient_email}")
        return True
    except Exception as e:
        print(f"Error sending to {recipient_email}: {e}")
        return False

# Prepare personalized messages
messages_to_send = []
for recipient_email in recipients:
    recipient_name = recipient_email.split('@')[0].replace('.', ' ').replace('_', ' ').title()
    html_body = html_body_template.format(
        recipient_name=recipient_name,
        generated_time=generated_time,
        high_break_headers=high_break_headers,
        high_break_rows=high_break_rows,
        high_break_stocks=high_break_stocks,
        onetime_five_headers=onetime_five_headers,
        onetime_five_rows=onetime_five_rows,
        export_url=export_url
    )
    messages_to_send.append((recipient_email, html_body))

# Send emails immediately after second refresh
print("Sending emails at 9:20 AM IST...")
emails_sent = 0
with ThreadPoolExecutor(max_workers=15) as executor:
    results = executor.map(lambda args: send_email(*args), messages_to_send)
    emails_sent = sum(1 for success in results if success)

# Wait 3 minutes until 9:23 AM IST
print("\nWaiting until 9:23 AM IST to refresh Google Sheet again...\n")
time.sleep(180)

# Execute the Google Apps Script macro again (third refresh at 9:23 AM)
execute_macro_with_retry(macro_url)
print("Third Google Sheet refresh completed at 9:23 AM IST.")
data_high_break_923 = high_break_trade_sheet.get_all_values()
stock_names_923 = [row[1] for row in data_high_break_923[3:] if row[1]]  # Fetch updated stocks
num_stocks_923 = len(stock_names_923)
compare_sheet.update(range_name='G4', values=[[stock] for stock in stock_names_923])  # Write to G4 downward
print(f"Wrote {num_stocks_923} stocks to column G at 9:23 AM IST.")
# Generate and write formulas row-by-row for 9:23 AM data
h_formulas = [f'=CONCAT("NSE:",G{row+4})' for row in range(num_stocks_923)]
i_formulas = [f'=GOOGLEFINANCE(H{row+4},"priceopen")' for row in range(num_stocks_923)]
j_formulas = [f'=GOOGLEFINANCE(H{row+4},"price")' for row in range(num_stocks_923)]
compare_sheet.update(range_name='H4:H' + str(3 + num_stocks_923), values=[[formula] for formula in h_formulas], value_input_option='USER_ENTERED')
compare_sheet.update(range_name='I4:I' + str(3 + num_stocks_923), values=[[formula] for formula in i_formulas], value_input_option='USER_ENTERED')
compare_sheet.update(range_name='J4:J' + str(3 + num_stocks_923), values=[[formula] for formula in j_formulas], value_input_option='USER_ENTERED')

time.sleep(10)  # Increased sleep to allow GOOGLEFINANCE to fetch data

# New section for orb_dhan sheet updates
orb_dhan_sheet = workbook.worksheet('orb_dhan')

# Flush (clear) existing data in columns B4 downward and H4 downward
orb_dhan_sheet.batch_clear(['B4:F', 'H4:L'])
print("\nCleared existing data from B4:B and H4:H in orb_dhan sheet.")

# Fetch stock names for 9:20 AM data from compare sheet column L (L4 downward)
stock_names_920 = compare_sheet.col_values(12)[3:]  # Column L is index 12 (1-based), skip rows 1-3
stock_names_920 = [stock for stock in stock_names_920 if stock]  # Filter out empty values
num_stocks_920 = len(stock_names_920)
orb_dhan_sheet.update(range_name='B4', values=[[stock] for stock in stock_names_920])
print(f"Wrote {num_stocks_920} stocks to column B in orb_dhan sheet.")

# Fetch stock names for 9:23 AM data from compare sheet column M (M4 downward)
stock_names_923 = compare_sheet.col_values(13)[3:]  # Column M is index 13 (1-based), skip rows 1-3
stock_names_923 = [stock for stock in stock_names_923 if stock]  # Filter out empty values
num_stocks_923 = len(stock_names_923)
orb_dhan_sheet.update(range_name='H4', values=[[stock] for stock in stock_names_923])
print(f"Wrote {num_stocks_923} stocks to column H in orb_dhan sheet.")

# Update formulas for 9:20 AM data (columns C, D, E) based on column B
orb_dhan_sheet.update(range_name='C4:C' + str(3 + num_stocks_920), values=[[f'=CONCAT("NSE:",B{row+4})'] for row in range(num_stocks_920)], value_input_option='USER_ENTERED')
orb_dhan_sheet.update(range_name='D4:D' + str(3 + num_stocks_920), values=[[f'=GOOGLEFINANCE(C{row+4},"priceopen")'] for row in range(num_stocks_920)], value_input_option='USER_ENTERED')
orb_dhan_sheet.update(range_name='E4:E' + str(3 + num_stocks_920), values=[[f'=GOOGLEFINANCE(C{row+4},"price")'] for row in range(num_stocks_920)], value_input_option='USER_ENTERED')
# Add percentage change formula for 9:20 AM data in column F
orb_dhan_sheet.update(range_name='F4:F' + str(3 + num_stocks_920), values=[[f'=(E{row+4}-D{row+4})/E{row+4}*100'] for row in range(num_stocks_920)], value_input_option='USER_ENTERED')

# Update formulas for 9:23 AM data (columns I, J, K) based on column H
orb_dhan_sheet.update(range_name='I4:I' + str(3 + num_stocks_923), values=[[f'=CONCAT("NSE:",H{row+4})'] for row in range(num_stocks_923)], value_input_option='USER_ENTERED')
orb_dhan_sheet.update(range_name='J4:J' + str(3 + num_stocks_923), values=[[f'=GOOGLEFINANCE(I{row+4},"priceopen")'] for row in range(num_stocks_923)], value_input_option='USER_ENTERED')
orb_dhan_sheet.update(range_name='K4:K' + str(3 + num_stocks_923), values=[[f'=GOOGLEFINANCE(I{row+4},"price")'] for row in range(num_stocks_923)], value_input_option='USER_ENTERED')
# Add percentage change formula for 9:23 AM data in column L
orb_dhan_sheet.update(range_name='L4:L' + str(3 + num_stocks_923), values=[[f'=(K{row+4}-J{row+4})/K{row+4}*100'] for row in range(num_stocks_923)], value_input_option='USER_ENTERED')

print(f"Updated formulas in orb_dhan sheet for 9:20 AM and 9:23 AM data.")

# New section to send email again with high_break_trade, onetime_five_open, and orb_dhan data
data_orb_dhan = orb_dhan_sheet.get_all_values()
data_high_break = high_break_trade_sheet.get_all_values()
data_onetime_five = onetime_five_open_sheet.get_all_values()[:last_row_b + 1]

# Build HTML table/grid for orb_dhan data
orb_dhan_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_orb_dhan[0][:6]) + ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_orb_dhan[0][7:13])
orb_dhan_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #dee2e6; background-color: {"#ffffff" if i % 2 == 0 else "#f8f9fa"}; color: #495057;">{val}</td>' for val in row[:6] + row[7:13]) + '</tr>'
    for i, row in enumerate(data_orb_dhan[1:])
)

# Build HTML table/grid for high_break_trade data
high_break_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_high_break[0])
high_break_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #dee2e6; background-color: {"#ffffff" if i % 2 == 0 else "#f8f9fa"}; color: #495057;">{val}</td>' for val in row) + '</tr>'
    for i, row in enumerate(data_high_break[1:])
)
high_break_stocks = ''.join(
    f'<li style="padding: 10px; background-color: #fff; margin: 8px 0; border-left: 4px solid #dc3545; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">{stock} (9:20 AM)</li>'
    for stock in stock_names_920
) + ''.join(
    f'<li style="padding: 10px; background-color: #fff; margin: 8px 0; border-left: 4px solid #28a745; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">{stock} (9:23 AM)</li>'
    for stock in stock_names_923
)

# Build HTML table/grid for onetime_five_open data
onetime_five_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #e9ecef; border: 1px solid #dee2e6; color: #343a40;">{col}</th>' for col in data_onetime_five[0])
onetime_five_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #dee2e6; background-color: {"#ffffff" if i % 2 == 0 else "#f8f9fa"}; color: #495057;">{val}</td>' for val in row) + '</tr>'
    for i, row in enumerate(data_onetime_five[1:last_row_b + 1])
)

export_url_orb_dhan = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
generated_time = (datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)).strftime('%Y-%m-%d %H:%M:%S') + ' IST'

# Updated HTML template to include orb_dhan section
html_body_template = """
<html>
  <head>
    <style>
      body {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #333; line-height: 1.6; background-color: #f4f4f4; margin: 0; padding: 20px; }}
      .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
      h3 {{ color: #2c3e50; margin-top: 30px; border-bottom: 2px solid #007bff; padding-bottom: 8px; }}
      p {{ margin: 10px 0; }}
      table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
      ul {{ list-style: none; padding: 0; }}
      .download-btn {{ display: inline-block; padding: 12px 24px; background-color: #007bff; color: #fff; text-decoration: none; border-radius: 5px; font-weight: 500; transition: background-color 0.3s, transform 0.2s; }}
      .download-btn:hover {{ background-color: #0056b3; transform: translateY(-2px); }}
      .footer {{ margin-top: 30px; text-align: center; color: #6c757d; font-size: 0.9em; }}
      @media (max-width: 600px) {{ .container {{ padding: 15px; }} table {{ font-size: 0.9em; }} }}
    </style>
  </head>
  <body>
    <div class="container">
      <p>Dear {recipient_name},</p>
      <p>Generated on: {generated_time}</p>
      <h3>High Break Trade Data</h3>
      <table>{high_break_headers}{high_break_rows}</table>
      <h3>Onetime Five Open Data</h3>
      <table>{onetime_five_headers}{onetime_five_rows}</table>
      <h3>Orb Dhan Data</h3>
      <table>{orb_dhan_headers}{orb_dhan_rows}</table>
      <h3>Stock Names (Column B & Column B for 9:23 )</h3>
      <ul>{high_break_stocks}</ul>
      <p><a href="{export_url}" class="download-btn">Download Excel</a></p>
      <div class="footer">This is an automated report from HighBreak Alert. For support, contact our team.</div>
    </div>
  </body>
</html>
"""

# Prepare personalized messages for second email
messages_to_send_orb_dhan = []
for recipient_email in recipients:
    recipient_name = recipient_email.split('@')[0].replace('.', ' ').replace('_', ' ').title()
    html_body = html_body_template.format(
        recipient_name=recipient_name,
        generated_time=generated_time,
        high_break_headers=high_break_headers,
        high_break_rows=high_break_rows,
        high_break_stocks=high_break_stocks,
        onetime_five_headers=onetime_five_headers,
        onetime_five_rows=onetime_five_rows,
        orb_dhan_headers=orb_dhan_headers,
        orb_dhan_rows=orb_dhan_rows,
        export_url=export_url_orb_dhan
    )
    messages_to_send_orb_dhan.append((recipient_email, html_body))

# Send emails with high_break_trade, onetime_five_open, and orb_dhan data
print("Sending second emails with high_break_trade, onetime_five_open, and orb_dhan data at 10:55 AM IST...")
emails_sent_orb_dhan = 0
with ThreadPoolExecutor(max_workers=15) as executor:
    results = executor.map(lambda args: send_email(*args), messages_to_send_orb_dhan)
    emails_sent_orb_dhan = sum(1 for success in results if success)

print(f"Sent {emails_sent_orb_dhan} second emails successfully with high_break_trade, onetime_five_open, and orb_dhan data.")

# Summary
elapsed_time = time.time() - start_time
print(f"\nSent {emails_sent} emails successfully. Process completed in {elapsed_time:.2f} seconds.\n")