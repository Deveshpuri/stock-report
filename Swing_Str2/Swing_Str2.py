from dotenv import load_dotenv
import os
import json
import requests
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import gspread
from google.oauth2.service_account import Credentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
import pytz
import logging
import yfinance as yf
import pandas as pd
import numpy as np
from scipy.signal import argrelextrema
import io
import csv

# Load environment variables from .env file
load_dotenv()

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Set IST timezone
ist = pytz.timezone('Asia/Kolkata')

# Record start time
start_time = time.time()
logger.info("Starting swing trading stock finder script with new strategy...")

# Step 1: Execute Google Apps Script macro twice with a 4-second pause
macro_url = "https://script.google.com/macros/s/AKfycbykFjLRDhZ9tcu20L0F-7aTirVffIvo6tn811Pn6ONsa06cGV0JnpUsXiYJ_o_kDQ/exec"
logger.info("Executing first Google Apps Script macro...")
try:
    response = requests.get(macro_url)
    response.raise_for_status()
    logger.info("First Google Apps Script macro executed successfully.")
except Exception as e:
    logger.error(f"Error executing first macro: {e}")

logger.info("Pausing for 4 seconds...")
time.sleep(4)

logger.info("Executing second Google Apps Script macro...")
try:
    response = requests.get(macro_url)
    response.raise_for_status()
    logger.info("Second Google Apps Script macro executed successfully.")
except Exception as e:
    logger.error(f"Error executing second macro: {e}")

# Step 2: Google Sheets authentication
try:
    credentials_info = json.loads(os.environ.get('GOOGLE_CREDENTIALS_JSON'))
    creds = Credentials.from_service_account_info(credentials_info, scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ])
    client = gspread.authorize(creds)
    logger.info("Google Sheets authentication successful")
except Exception as e:
    logger.error(f"Error loading credentials: {e}")
    exit(1)

# Open the Google Spreadsheet
sheet_id = '1TETP6W6Ee5J7GCMHbSZMLykZlFJOy-qk5JFWN0DYgI4'
try:
    workbook = client.open_by_key(sheet_id)
    logger.info("Successfully opened Google Spreadsheet")
except Exception as e:
    logger.error(f"Error opening spreadsheet: {e}")
    exit(1)

# Get stock symbols from Sheet1 (Column B, starting from row 4)
try:
    sheet1 = workbook.worksheet('Sheet1')
    stock_symbols = [s for s in sheet1.col_values(2)[3:] if s.strip()]
    logger.info(f"Retrieved {len(stock_symbols)} stock symbols from Sheet1: {stock_symbols[:5]}...")
except Exception as e:
    logger.error(f"Error reading stock symbols from Sheet1: {e}")
    exit(1)

if not stock_symbols:
    logger.error("No stock symbols found in Sheet1, Column B, starting from row 4")
    exit(1)

# Headers for Calculation sheet (updated to start at B2)
headers = [
    "Stock", "NSE Symbol", "Latest Close", "Regression(200)", "Close > Regression",
    "Recent Upside Break", "Very Close to Break", "ADX(14)", "RSI(14)",
    "Volume > 1.5 * SMA_Vol20", "All Conditions Met"
]

# Create or open Calculation sheet and set headers
try:
    calc_sheet = workbook.worksheet('Calculation')
    logger.info("Opened Calculation sheet")
except gspread.WorksheetNotFound:
    calc_sheet = workbook.add_worksheet(title='Calculation', rows=100, cols=12)  # Adjusted to 12 columns
    logger.info("Created Calculation sheet")

# Clear and set headers in B2:L2 every time
calc_sheet.update(values=[[""] * 11], range_name='B1:P2000')  # Empty row 1
calc_sheet.update(values=[headers], range_name='B2:L2')    # Headers from B2
logger.info("Set headers in Calculation sheet at B2:L2")

# Manual RSI calculation
def calculate_rsi(close, period=14):
    delta = close.diff(1)
    gain = np.where(delta > 0, delta, 0)
    loss = np.where(delta < 0, -delta, 0)
    avg_gain = pd.Series(gain).rolling(window=period, min_periods=1).mean()
    avg_loss = pd.Series(loss).rolling(window=period, min_periods=1).mean()
    rs = avg_gain / avg_loss.replace(0, np.finfo(float).eps)
    return float(100 - (100 / (1 + rs.iloc[-1])))

# Manual ADX calculation
def calculate_adx(high, low, close, period=14):
    tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
    dm_plus = (high - high.shift(1)).where((high - high.shift(1) > low.shift(1) - low) & (high - high.shift(1) > 0), 0)
    dm_minus = (low.shift(1) - low).where((low.shift(1) - low > high - high.shift(1)) & (low.shift(1) - low > 0), 0)
    atr = tr.ewm(span=period, adjust=False).mean()
    di_plus = 100 * (dm_plus.ewm(span=period, adjust=False).mean() / atr.replace(0, np.finfo(float).eps))
    di_minus = 100 * (dm_minus.ewm(span=period, adjust=False).mean() / atr.replace(0, np.finfo(float).eps))
    dx = 100 * abs(di_plus - di_minus) / (di_plus + di_minus).replace(0, np.finfo(float).eps)
    return float(dx.ewm(span=period, adjust=False).mean().iloc[-1])

# Manual ATR calculation
def calculate_atr(high, low, close, period=14):
    close_shift = close.shift(1)
    tr1 = high - low
    tr2 = abs(high - close_shift)
    tr3 = abs(low - close_shift)
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    return tr.rolling(window=period, min_periods=1).mean().iloc[-1]

# Calculate Linear Regression value
def calculate_regression(close, length=200):
    if len(close) < length:
        return np.nan
    x = np.arange(length)
    y = close[-length:].values
    slope, intercept = np.polyfit(x, y, 1)
    return slope * (length - 1) + intercept

# Detect recent upside break or closeness
def detect_upside_break(high, low, close, lookback=24, slope_factor=1.0):
    order = lookback // 2
    ph_idx = argrelextrema(high.values, np.greater, order=order)[0]
    if len(ph_idx) < 2:
        return False, False
    i1, i2 = ph_idx[-2], ph_idx[-1]
    ph1, ph2 = high.iloc[i1], high.iloc[i2]
    delta = i2 - i1
    slope = (ph2 - ph1) / delta
    if slope >= 0:
        return False, False
    atr = calculate_atr(high, low, close)
    min_slope_threshold = -slope_factor * atr / delta
    if slope > min_slope_threshold:
        return False, False
    n = len(high) - 1
    trend_value = ph2 + slope * (n - i2)
    recent_break = False
    for k in range(1, 21):
        m = n - k + 1
        trend_m = ph2 + slope * (m - i2)
        if close.iloc[m-1] < trend_m and close.iloc[m] > trend_m:
            recent_break = True
            break
    very_close = not recent_break and close.iloc[-1] < trend_value and (trend_value - close.iloc[-1]) < 0.5 * atr
    return recent_break, very_close

# Function to fetch and calculate stock data
def get_stock_data(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="1y", interval="1d", auto_adjust=False, prepost=False)
        
        if hist.empty or len(hist) < 200:
            return None, symbol
        
        latest_close = float(hist['Close'].iloc[-1])
        regression_200 = calculate_regression(hist['Close'])
        adx_14 = calculate_adx(hist['High'], hist['Low'], hist['Close'])
        rsi_14 = calculate_rsi(hist['Close'])
        latest_volume = int(hist['Volume'].iloc[-1])
        sma_vol_20 = float(hist['Volume'].rolling(window=20).mean().iloc[-1])
        recent_break, very_close = detect_upside_break(hist['High'], hist['Low'], hist['Close'])
        
        if pd.isna(regression_200) or pd.isna(sma_vol_20):
            logger.warning(f"Missing data for {symbol}")
            return None, symbol

        close_gt_regression = "Yes" if latest_close > regression_200 else "No"
        upside_break = "Yes" if recent_break else "No"
        close_to_break = "Yes" if very_close else "No"
        vol_gt_1_5_sma = "Yes" if latest_volume > 1.5 * sma_vol_20 else "No"
        all_conditions_met = "Yes" if all([
            close_gt_regression == "Yes",
            (upside_break == "Yes" or close_to_break == "Yes"),
            adx_14 > 25,
            rsi_14 > 40,
            vol_gt_1_5_sma == "Yes"
        ]) else "No"

        return [symbol[:-3], f"NSE:{symbol[:-3]}", latest_close, regression_200, close_gt_regression,
                upside_break, close_to_break, adx_14, rsi_14, vol_gt_1_5_sma, all_conditions_met], symbol
    except Exception as e:
        logger.error(f"Error fetching data for {symbol}: {e}")
        return None, symbol

# Test with INTERARCH.NS
logger.info("Testing data fetch with INTERARCH.NS...")
test_data, _ = get_stock_data("INTERARCH.NS")
if test_data:
    logger.info(f"Test fetch successful: {test_data}")
else:
    logger.error("Test fetch failed for INTERARCH.NS. Check yfinance connectivity or API status.")

# Batch update data for Calculation sheet
batch_data = []
start_row = 4

with ThreadPoolExecutor(max_workers=12) as executor:
    future_to_symbol = {}
    for symbol in stock_symbols:
        future = executor.submit(get_stock_data, f"{symbol}.NS")
        future_to_symbol[future] = symbol
    for future in as_completed(future_to_symbol):
        try:
            data, symbol = future.result()
            if data:
                batch_data.append(data)
            else:
                print(f"No valid data returned for {symbol}")
        except Exception as e:
            print(f"Error processing result for {future_to_symbol[future]}: {e}")

if batch_data:
    try:
        calc_sheet.update(values=batch_data, range_name=f'B{start_row}:L{start_row + len(batch_data) - 1}')
        logger.info(f"Updated Calculation sheet with {len(batch_data)} rows")
        time.sleep(3)
    except Exception as e:
        logger.error(f"Error updating Calculation sheet: {e}")
else:
    logger.warning("No data to update in Calculation sheet")

# Wait until 9:25 IST to refresh Google Sheet
current_time_ist = datetime.now(ist)
target_time_ist_1 = current_time_ist.replace(hour=9, minute=25, second=0, microsecond=0)
if current_time_ist < target_time_ist_1:
    wait_seconds = (target_time_ist_1 - current_time_ist).total_seconds()
    print(f"Waiting {wait_seconds:.0f} seconds until 9:25 AM IST to refresh Google Sheet...")
    for i in range(int(wait_seconds), 0, -1):
        print(f"Waiting... {i} seconds remaining", end='\r')
        time.sleep(1)

# Fetch data from swing_stock just before sending email
try:
    swing_stock_sheet = workbook.worksheet('swing_stock')
    swing_stock_data = swing_stock_sheet.get_all_values()
    credential_sheet = workbook.worksheet('credential')
    
    # Get all values from column D (index 3)
    email_column = credential_sheet.col_values(4)  # Fourth column (index 3)
    
    # Filter out empty strings and get emails starting from row 3 (index 2)
    recipients = []
    for i, email in enumerate(email_column):
        if i >= 2 and email and email.strip():  # Start from row 3 (index 2)
            recipients.append(email.strip())
    
    if not recipients:
        logger.warning("No valid recipient emails found in credential sheet after filtering")
    
    met_stocks = []
    for row in swing_stock_data[3:]:
        if len(row) > 10 and row[10] == "Yes":  # Check column K (index 10) for "Yes"
            met_stocks.append(row[1])  # Using column B (index 1) for stock names
    
    logger.info(f"Retrieved {len(met_stocks)} stocks meeting all conditions from swing_stock: {met_stocks[:5]}...")
    logger.info(f"Retrieved {len(recipients)} recipient emails from credential sheet: {recipients}")
    
except gspread.exceptions.WorksheetNotFound:
    logger.error("Credential sheet not found")
    exit(1)
except Exception as e:
    logger.error(f"Error fetching sheet data for email: {e}")
    logger.error(f"Email column content: {email_column}")
    exit(1)

# Manage a dedicated spreadsheet for swing_stock exports
export_spreadsheet_id = os.environ.get('SWING_STOCK_EXPORT_SHEET_ID')
export_spreadsheet = None

try:
    if export_spreadsheet_id:
        # Try to open existing spreadsheet
        export_spreadsheet = client.open_by_key(export_spreadsheet_id)
        logger.info(f"Opened existing swing_stock export spreadsheet: {export_spreadsheet.title}")
    else:
        # Create new spreadsheet if ID not set
        export_spreadsheet = client.create("Swing_Stock_Export")
        os.environ['SWING_STOCK_EXPORT_SHEET_ID'] = export_spreadsheet.id
        logger.info(f"Created new swing_stock export spreadsheet: {export_spreadsheet.title}")
        
    # Ensure only one sheet exists
    temp_sheet = export_spreadsheet.sheet1
    temp_sheet.update_title('swing_stock')
    for ws in export_spreadsheet.worksheets():
        if ws.title != 'swing_stock':
            export_spreadsheet.del_worksheet(ws)
    
    # Update with swing_stock data
    temp_sheet.update(swing_stock_data)
    logger.info("Updated swing_stock sheet with latest data")
    
    # Set sharing permissions to "Anyone with the link"
    export_spreadsheet.share('', perm_type='anyone', role='reader')
    
    # Generate export URL
    export_url = f"https://docs.google.com/spreadsheets/d/{export_spreadsheet.id}/export?format=xlsx"
    logger.info(f"Export URL for swing_stock sheet: {export_url}")
    
    # Prepare CSV data as fallback
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerows(swing_stock_data)
    csv_data = csv_buffer.getvalue()
    csv_buffer.close()
    
except Exception as e:
    logger.error(f"Error managing swing_stock export spreadsheet: {e}")
    logger.info("Falling back to CSV attachment")
    export_url = None  # Disable export URL
    csv_data = io.StringIO()
    writer = csv.writer(csv_data)
    writer.writerows(swing_stock_data)
    csv_data = csv_data.getvalue()

# Create HTML content for swing_stock sheet
swing_stock_headers = ''.join(f'<th style="padding: 12px; text-align: left; font-weight: 600; background-color: #2c3e50; color: #ffffff; border: 1px solid #34495e;">{col}</th>' for col in swing_stock_data[0])
swing_stock_rows = ''.join(
    '<tr>' + ''.join(f'<td style="padding: 12px; border: 1px solid #ecf0f1; background-color: {"#ffffff" if i % 2 == 0 else "#f9fbfd"}; color: #2c3e50;">{val}</td>' for val in row) + '</tr>'
    for i, row in enumerate(swing_stock_data[1:], 1)
)

# Check if we have any stocks meeting conditions
if met_stocks:
    met_stocks_html = ''.join(
        f'<li style="padding: 12px; background-color: #ffffff; margin: 8px 0; border-left: 5px solid #e74c3c; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); color: #2c3e50; font-weight: 500;">{stock}</li>'
        for stock in met_stocks
    )
else:
    met_stocks_html = '<li style="padding: 12px; background-color: #f8f9fa; margin: 8px 0; border-left: 5px solid #6c757d; border-radius: 6px; color: #6c757d; font-style: italic;">No stocks met all conditions today</li>'

generated_time = datetime.now(ist).strftime('%H:%M:%S') + ' IST'
html_body_template = """
<html>
  <head>
    <style>
      body {{ font-family: 'Arial', sans-serif; color: #2c3e50; line-height: 1.6; background-color: #ecf0f1; margin: 0; padding: 20px; }}
      .container {{ max-width: 900px; margin: 0 auto; background: #ffffff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
      .header {{ background-color: #3498db; color: #ffffff; padding: 15px; text-align: center; border-radius: 8px 8px 0 0; }}
      h3 {{ color: #2c3e50; margin-top: 25px; border-bottom: 3px solid #3498db; padding-bottom: 10px; font-size: 1.5em; }}
      p {{ margin: 15px 0; font-size: 1.1em; }}
      table {{ border-collapse: collapse; width: 100%; margin-bottom: 25px; }}
      th, td {{ padding: 12px; text-align: left; }}
      ul {{ list-style: none; padding: 0; }}
      .download-btn {{ display: inline-block; padding: 14px 28px; background-color: #e74c3c; color: #ffffff; text-decoration: none; border-radius: 8px; font-weight: 600; transition: all 0.3s ease; }}
      .download-btn:hover {{ background-color: #c0392b; transform: scale(1.05); }}
      .footer {{ margin-top: 30px; text-align: center; color: #7f8c8d; font-size: 0.9em; border-top: 1px solid #ecf0f1; padding-top: 15px; }}
      @media (max-width: 600px) {{ .container {{ padding: 15px; }} table {{ font-size: 0.9em; }} .download-btn {{ padding: 10px 20px; }} }}
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">Swing Stock Alert Report</div>
      <p>Dear {recipient_name},</p>
      <p><strong>Generated at:</strong> {generated_time}</p>
      <h3>Swing Stock Data</h3>
      <table>{swing_stock_headers}{swing_stock_rows}</table>
      <h3>Stocks Meeting All Conditions</h3>
      <ul>{met_stocks_html}</ul>
      {download_link}
      <div class="footer">This is an automated report from Swing Stock Alert. For support, contact our team.</div>
    </div>
  </body>
</html>
"""

# SMTP Configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587
username = os.environ.get('SMTP_USERNAME')
password = os.environ.get('SMTP_PASSWORD')

# Validate SMTP credentials
if not username or not password:
    logger.error("SMTP credentials not found in environment variables")
    exit(1)
    
sender = '"SwingScan Insights" <' + username + '>'

def send_email(recipient_email, html_body, csv_data=None):
    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'Swing Stock Alert Report'
        msg['From'] = sender
        msg['To'] = recipient_email
        msg.attach(MIMEText(html_body, 'html'))
        
        if csv_data:
            part = MIMEBase('text', 'csv')
            part.set_payload(csv_data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="swing_stock.csv"')
            msg.attach(part)
        
        with smtplib.SMTP(smtp_server, smtp_port) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(username, password)
            s.send_message(msg)
            logger.info(f"Email sent to {recipient_email}")
            return True
    except Exception as e:
        logger.error(f"Error sending to {recipient_email}: {e}")
        return False

# Prepare personalized messages
messages_to_send = []
for recipient_email in recipients:
    recipient_name = recipient_email.split('@')[0].replace('.', ' ').replace('_', ' ').title()
    download_link = f'<p><a href="{export_url}" class="download-btn">Download swing_stock Sheet (Excel)</a></p>' if export_url else '<p>Swing stock data attached as CSV.</p>'
    html_body = html_body_template.format(
        recipient_name=recipient_name,
        generated_time=generated_time,
        swing_stock_headers=swing_stock_headers,
        swing_stock_rows=swing_stock_rows,
        met_stocks_html=met_stocks_html,
        download_link=download_link
    )
    messages_to_send.append((recipient_email, html_body, csv_data if not export_url else None))

# Send emails with swing_stock data using ThreadPoolExecutor
logger.info(f"Attempting to send {len(messages_to_send)} emails...")
emails_sent = 0

with ThreadPoolExecutor(max_workers=6) as executor:
    futures = []
    for recipient_email, html_body, csv_attachment in messages_to_send:
        future = executor.submit(send_email, recipient_email, html_body, csv_attachment)
        futures.append(future)
    
    for future in as_completed(futures):
        try:
            if future.result():
                emails_sent += 1
        except Exception as e:
            logger.error(f"Error in email sending future: {e}")

logger.info(f"Stocks meeting all conditions: {', '.join(met_stocks)}")
print(f"Stocks meeting all conditions: {', '.join(met_stocks)}")

# Refresh the Google Sheet
try:
    response = requests.get(macro_url)
    response.raise_for_status()
    print("Google Sheet refreshed successfully.")
except Exception as e:
    logger.error(f"Error refreshing Google Sheet: {e}")

elapsed_time = time.time() - start_time
logger.info(f"Sent {emails_sent} emails successfully. Process completed in {elapsed_time:.2f} seconds.")
print(f"Sent {emails_sent} emails successfully. Process completed in {elapsed_time:.2f} seconds.")

print("Time lag: 15 seconds")
time.sleep(15)

# Final refresh
try:
    requests.get(macro_url)
    print("Final Google Sheet refresh completed.")
except Exception as e:
    logger.error(f"Error with final refresh: {e}")