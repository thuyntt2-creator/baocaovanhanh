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
    print("Loading data thô...")
    ws_raw = sh.worksheet("data thô")
    raw_vals = ws_raw.get_all_values()
    df_raw = pd.DataFrame(raw_vals[1:], columns=raw_vals[0])
    df_raw['Volume'] = pd.to_numeric(df_raw['Volume'], errors='coerce').fillna(0)
    df_raw['WarehouseID'] = df_raw['WarehouseID'].str.strip()
    
    # Filter for 2026-06-25
    df_25 = df_raw[df_raw['Ngay'] == '2026-06-25']
    total_vol_25 = df_25['Volume'].sum()
    print(f"\n--- 2026-06-25 Stats ---")
    print(f"Total rows on 2026-06-25: {len(df_25)}")
    print(f"Total volume on 2026-06-25: {total_vol_25}")
    
    # Top warehouses on 2026-06-25 by Volume
    top_wh = df_25.groupby('WarehouseID')['Volume'].sum().reset_index()
    top_wh = top_wh.sort_values(by='Volume', ascending=False)
    print("\nTop 20 Warehouses on 2026-06-25 by Volume:")
    print(top_wh.head(20).to_string(index=False))
    
    # 2. Check if they are in 'Cơ cấu' or 'CoCauVung'
    ws_cocau = sh.worksheet("Cơ cấu")
    cc_vals = ws_cocau.get_all_values()
    df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
    df_cc['Mã bưu cục'] = df_cc['Mã bưu cục'].str.strip()
    cc_map = df_cc.set_index('Mã bưu cục')['Bưu cục'].to_dict()
    
    ws_vung = sh.worksheet("CoCauVung")
    vung_vals = ws_vung.get_all_values()
    df_vung = pd.DataFrame(vung_vals[1:], columns=vung_vals[0])
    df_vung.columns = [c.strip() for c in df_vung.columns]
    df_vung['warehouse_id'] = df_vung['warehouse_id'].str.strip()
    vung_map = df_vung.set_index('warehouse_id')['Bưu cục'].to_dict()
    
    print("\nChecking top 20 warehouses in mappings:")
    top_20_ids = top_wh.head(20)['WarehouseID'].tolist()
    for wid in top_20_ids:
        in_cc = cc_map.get(wid, "Not Found")
        in_vung = vung_map.get(wid, "Not Found")
        print(f"  WarehouseID {wid:12s} -> In Cơ cấu: {in_cc:35s} | In CoCauVung: {in_vung}")

if __name__ == "__main__":
    main()
