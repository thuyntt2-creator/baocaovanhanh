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
    df_blank = df[df['NgayLTC'] == '']
    
    print(f"Number of rows where NgayLTC is empty: {len(df_blank)}")
    
    # Check if there are rows that have non-empty columns but empty NgayLTC
    # We check columns: AM, Don_ltc, KhoLay, SellerName
    non_empty_rows = df_blank[(df_blank['AM'] != '') | (df_blank['Don_ltc'] != '') | (df_blank['KhoLay'] != '')]
    print(f"Number of rows with empty NgayLTC but non-empty data: {len(non_empty_rows)}")
    
    if len(non_empty_rows) > 0:
        print("\nFirst 10 rows with empty NgayLTC but non-empty data:")
        print(non_empty_rows.head(10).to_string())
    else:
        print("\nAll empty NgayLTC rows are completely empty.")

if __name__ == "__main__":
    main()
