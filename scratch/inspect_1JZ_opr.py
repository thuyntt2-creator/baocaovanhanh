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
SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    print(f"Opened spreadsheet: {sh.title}")
    
    ws_opr = sh.worksheet("OPR")
    raw_values = ws_opr.get_all_values()
    print(f"Row count of 'OPR' in 1JZ: {len(raw_values)}")
    if len(raw_values) < 2:
        print("Empty worksheet.")
        return
        
    df = pd.DataFrame(raw_values[1:], columns=raw_values[0])
    print("\nColumns in 1JZ OPR:")
    print(list(df.columns))
    
    if 'NgayLTC' in df.columns:
        print("\nUnique values of 'NgayLTC':")
        print(df['NgayLTC'].value_counts())
        
        df_26 = df[df['NgayLTC'] == '2026-06-26']
        print(f"\nNumber of rows for 2026-06-26 in 1JZ: {len(df_26)}")
        if len(df_26) > 0:
            print("\nFirst 5 rows for 2026-06-26:")
            print(df_26.head(5).to_string())
    else:
        print("\nNgayLTC not found in 1JZ OPR!")

if __name__ == "__main__":
    main()
