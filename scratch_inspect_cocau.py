import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding for Windows Command Prompt / Tasks Scheduler
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
    
    # 1. Inspect 'Cơ cấu' worksheet
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_data = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_data[1:], columns=cocau_data[0])
    print(f"Cơ cấu row count: {len(df_cocau)}")
    print("Cơ cấu columns:", df_cocau.columns.tolist())
    print("Cơ cấu first 10 rows:")
    print(df_cocau.head(10).to_string())
    
    # 2. Inspect 'data thô' worksheet
    ws_raw = sh.worksheet("data thô")
    # Let's get only the headers and first 100 rows to speed up and inspect
    raw_headers = ws_raw.get_values("A1:Z1")[0]
    raw_sample = ws_raw.get_values("A2:Z100")
    df_raw = pd.DataFrame(raw_sample, columns=raw_headers[:len(raw_sample[0])])
    print(f"\ndata thô columns: {raw_headers}")
    print("data thô first 5 rows:")
    print(df_raw.head(5).to_string())
    
    # Let's count matching WarehouseID
    ws_raw_all = ws_raw.get_all_values()
    df_raw_all = pd.DataFrame(ws_raw_all[1:], columns=ws_raw_all[0])
    
    # Strip whitespace and convert to string
    df_raw_all['WarehouseID'] = df_raw_all['WarehouseID'].str.strip()
    df_cocau['Mã bưu cục'] = df_cocau['Mã bưu cục'].str.strip()
    
    raw_wids = set(df_raw_all['WarehouseID'].unique())
    cocau_codes = set(df_cocau['Mã bưu cục'].unique())
    
    print(f"\nUnique WarehouseIDs in data thô: {len(raw_wids)}")
    print(f"Unique Mã bưu cục in Cơ cấu: {len(cocau_codes)}")
    
    matched = raw_wids.intersection(cocau_codes)
    print(f"Matched WarehouseIDs: {len(matched)}")
    
    unmatched = raw_wids - cocau_codes
    print(f"Unmatched WarehouseIDs count: {len(unmatched)}")
    print("Some unmatched WarehouseIDs:", list(unmatched)[:30])
    
    # Let's check volume sum
    df_raw_all['Volume'] = pd.to_numeric(df_raw_all['Volume'], errors='coerce').fillna(0)
    total_raw_vol = df_raw_all['Volume'].sum()
    matched_vol = df_raw_all[df_raw_all['WarehouseID'].isin(cocau_codes)]['Volume'].sum()
    print(f"\nTotal raw Volume: {total_raw_vol}")
    print(f"Volume of matched WarehouseIDs: {matched_vol}")

if __name__ == "__main__":
    main()
