import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Load data thô
    ws_raw = sh.worksheet("data thô")
    raw_vals = ws_raw.get_all_values()
    df_raw = pd.DataFrame(raw_vals[1:], columns=raw_vals[0])
    df_raw['Volume'] = pd.to_numeric(df_raw['Volume'], errors='coerce').fillna(0)
    df_raw['WarehouseID'] = df_raw['WarehouseID'].str.strip()
    
    # 2. Check specific IDs on 2026-06-11
    target_ids = ['20144000', '20633000']
    df_11 = df_raw[df_raw['Ngay'] == '2026-06-11']
    
    print("\n--- Volume in 'data thô' on 2026-06-11 ---")
    for tid in target_ids:
        sub = df_11[df_11['WarehouseID'] == tid]
        print(f"WarehouseID {tid}: {len(sub)} rows, Total Volume: {sub['Volume'].sum()}")
        
    # Let's inspect unique WarehouseIDs on 2026-06-11
    print("\nTotal unique WarehouseIDs on 2026-06-11:", df_11['WarehouseID'].nunique())

if __name__ == "__main__":
    main()
