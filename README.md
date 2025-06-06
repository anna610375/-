# PrintOffice Automation

This project automates fetching payments from FinTablo and posting them to PrintOffice24 using Selenium.

## Prerequisites
- Python 3.10+
- Google Chrome with a saved user profile
- `chromedriver` handled by `webdriver_manager`

Install dependencies:
```bash
pip install -r requirements.txt
```

## Environment Variables
Set these variables before running scripts:
- `PRINTOFFICE_USER` – login for PrintOffice24
- `PRINTOFFICE_PASS` – password for PrintOffice24
- `FIN_TABLO_PATH` – path to the FinTablo export Excel file

You can create a `.env` file and use `dotenv` to load it if desired.

## Running the Payment Script
```bash
python apply_payments/parse_fin_tablo_and_apply_payments.py
```
This script reads yesterday's payments from FinTablo and applies them to matching deals in PrintOffice24. A summary is written to `check_payments.xlsx`.
