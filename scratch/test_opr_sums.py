import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    ws_opr = sh.worksheet("OPR")
    raw_values = ws_opr.get_all_values()
    headers = raw_values[0]
    
    df = pd.DataFrame(raw_values[1:], columns=headers)
    
    # Clean numeric columns
    df['Don_ltc'] = pd.to_numeric(df['Don_ltc'], errors='coerce').fillna(0)
    df['Don_ontime'] = pd.to_numeric(df['Don_ontime'], errors='coerce').fillna(0)
    
    print(f"Total rows in OPR sheet: {len(df)}")
    print(f"Sum of Don_ltc globally: {df['Don_ltc'].sum()}")
    print(f"Sum of Don_ontime globally: {df['Don_ontime'].sum()}")
    
    print("\nSums by date:")
    date_sums = df.groupby('NgayLTC')[['Don_ltc', 'Don_ontime']].sum().reset_index()
    print(date_sums.to_string())

if __name__ == "__main__":
    main()
