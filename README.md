# 📊 Stock Report Automation

**Deveshpuri Goswami** 👋  

Automate daily stock reporting with Python, Google Sheets, and GitHub Actions.  
This project fetches stock data, computes technical indicators, updates a Google Sheet, and emails sleek HTML reports—all on autopilot.

---

## 🚀 Features

- ⏰ **Daily Automation**  
  Runs every day at **9:00 AM IST / 03:30 UTC**, orchestrated via GitHub Actions.

- 📈 **Technical Calculations**  
  - SMA(200)  
  - RSI(14)  
  - ADX(14)  
  - Volume SMA(20)  
  - Smart trading signals (e.g., Close > SMA, RSI > 40)

- 📊 **Google Sheets Integration**  
  Reads ticker symbols, updates a “Calculation” sheet, and refreshes using Apps Script.

- 📧 **HTML Email Reports**  
  Stylishly formatted report including:  
  - Swing stock data  
  - High-break alerts  
  - Download link for Excel export  

- ⚡ **Efficient Workflow**  
  Stock data fetched in parallel for speed, plus robust logging and error handling.

---

## 🛠️ Tech Stack

- **Python 3.12**  
- Core libraries:  
  - `yfinance`, `pandas`, `numpy` – financial data & indicator computations  
  - `gspread`, `google-auth` – interfacing with Google Sheets  
  - `smtplib`, `email` – sending emails via Gmail  
- **GitHub Actions** for automation and CI/CD  

---

## 📂 Workflow Overview

1. Trigger Google Apps Script to refresh data  
2. Load stock symbols from Google Sheets  
3. Fetch 1-year historical data via Yahoo Finance  
4. Calculate indicators (SMA, RSI, ADX, volume, etc.)  
5. Update “Calculation” tab in the Sheet  
6. Gather swing trade and high-break data from other sheets  
7. Craft an HTML email template  
8. Send emails to addresses listed in the sheet  

---

## ⚙️ Setup Instructions

1. **Clone the repo**  
   ```bash
   git clone https://github.com/Deveshpuri/stock-report.git
   cd stock-report
   ```

2. **Add GitHub Secrets**  
   - `GOOGLE_CREDENTIALS_JSON` – your service account JSON  
   - `SMTP_USERNAME` – Gmail username  
   - `SMTP_PASSWORD` – Gmail app-specific password  

3. **Install Dependencies**  
   ```bash
   pip install -r requirements.txt
   ```

4. **Run Locally (Optional)**  
   ```bash
   python test1.py
   ```

---

## 📅 Schedule

This workflow runs daily at:  
**09:00 AM IST / 03:30 AM UTC** — configured via GitHub Actions cron schedule.

---

## ✨ Author

**Deveshpuri Goswami**  
Automating stock insights daily with Python, Google Sheets, and GitHub Actions 🚀
