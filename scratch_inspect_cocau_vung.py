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
    
    # 1. Read 'CoCauVung'
    ws_vung = sh.worksheet("CoCauVung")
    vung_data = ws_vung.get_all_values()
    df_vung = pd.DataFrame(vung_data[1:], columns=vung_data[0])
    # Strip column names
    df_vung.columns = [c.strip() for c in df_vung.columns]
    print(f"CoCauVung row count: {len(df_vung)}")
    print("CoCauVung columns:", df_vung.columns.tolist())
    print("CoCauVung first 10 rows:")
    print(df_vung.head(10).to_string())
    
    # 2. Check overlap between data thô and CoCauVung
    ws_raw = sh.worksheet("data thô")
    raw_vals = ws_raw.get_all_values()
    df_raw = pd.DataFrame(raw_vals[1:], columns=raw_vals[0])
    
    raw_wids = set(df_raw['WarehouseID'].str.strip().unique())
    vung_wids = set(df_vung['warehouse_id'].str.strip().unique())
    
    print(f"\nUnique WarehouseIDs in data thô: {len(raw_wids)}")
    print(f"Unique warehouse_ids in CoCauVung: {len(vung_wids)}")
    
    matched = raw_wids.intersection(vung_wids)
    print(f"Matched WarehouseIDs: {len(matched)}")
    print("Unmatched WarehouseIDs count:", len(raw_wids - vung_wids))
    print("Some unmatched WarehouseIDs:", list(raw_wids - vung_wids)[:20])
    
    # Calculate volume of matched WarehouseIDs
    df_raw['Volume'] = pd.to_numeric(df_raw['Volume'], errors='coerce').fillna(0)
    total_vol = df_raw['Volume'].sum()
    matched_vol = df_raw[df_raw['WarehouseID'].str.strip().isin(vung_wids)]['Volume'].sum()
    print(f"\nTotal raw Volume: {total_vol}")
    print(f"Volume of matched WarehouseIDs in CoCauVung: {matched_vol}")

if __name__ == "__main__":
    main()
