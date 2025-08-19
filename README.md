📊 Stock Report Automation

Hi, I’m Deveshpuri Goswami 👋

This project is an automated stock reporting system that runs daily using GitHub Actions. It fetches stock data, performs calculations, updates Google Sheets, and sends professional HTML email reports.

🚀 Features

⏳ Automated Schedule – Runs daily at 9:00 AM IST via GitHub Actions.

📈 Stock Analysis – Calculates SMA(200), ADX(14), RSI(14), Volume SMA, and trading conditions.

📊 Google Sheets Integration – Reads stock symbols, updates calculation sheet, and refreshes via Apps Script.

📧 Email Reports – Sends a styled HTML report with stock data and download link to all recipients from the sheet.

⚡ Parallel Processing – Uses multi-threading for faster stock data fetch.

🛠️ Tech Stack

Python (3.12)

yfinance, pandas, numpy – Stock data & indicators

gspread, Google Sheets API – Spreadsheet integration

smtplib – Email sending via Gmail

GitHub Actions – CI/CD automation

📂 How It Works

Refreshes Google Sheet using Google Apps Script.

Fetches stock symbols from Sheet1.

Computes indicators (SMA200, RSI14, ADX14, etc.).

Updates results in Calculation sheet.

Collects swing trade data from other sheets.

Builds a professional HTML email report.

Sends emails to all recipients in credential sheet.

⚙️ Setup

Clone this repo.

Add required secrets in GitHub:

GOOGLE_CREDENTIALS_JSON

SMTP_USERNAME

SMTP_PASSWORD

Push changes – GitHub Actions will automatically run the workflow.

📅 Schedule

The workflow runs every day at:
🕘 09:00 AM IST (03:30 AM UTC)

✨ Author

👨‍💻 Deveshpuri Goswami
Automating stock reports with Python + Google Sheets + GitHub Actions.