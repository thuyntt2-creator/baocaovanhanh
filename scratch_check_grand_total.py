import os
import sys
import io
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

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
    # 1. Load Excel DB sheet
    excel_path = os.path.join(BASE_DIR, 'manual_raw.xlsx')
    if not os.path.exists(excel_path):
        print("manual_raw.xlsx not found.")
        return
    df_excel_db = pd.read_excel(excel_path, sheet_name='DB')
    
    # 2. Connect to Google Sheets to load data thô
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_raw = sh.worksheet("data thô")
    raw_vals = ws_raw.get_all_values()
    df_raw = pd.DataFrame(raw_vals[1:], columns=raw_vals[0])
    df_raw['Volume'] = pd.to_numeric(df_raw['Volume'], errors='coerce').fillna(0)
    
    # Let's inspect some dates, e.g. 2026-06-11, 2026-06-12, 2026-06-13
    dates_to_check = ['2026-06-11', '2026-06-12', '2026-06-13']
    for d in dates_to_check:
        print(f"\n--- Checking date: {d} ---")
        # 1. Total Volume in data thô (all company)
        total_raw_vol = df_raw[df_raw['Ngay'] == d]['Volume'].sum()
        print(f"Total raw volume in 'data thô' (all company): {total_raw_vol}")
        
        # 2. Sum of matched NTB warehouses in data thô (using Mã bưu cục from Cơ cấu in Google Sheet)
        ws_cc = sh.worksheet("Cơ cấu")
        cc_vals = ws_cc.get_all_values()
        df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
        cc_codes = set(df_cc['Mã bưu cục'].str.strip().unique())
        
        ntb_raw_vol = df_raw[(df_raw['Ngay'] == d) & (df_raw['WarehouseID'].str.strip().isin(cc_codes))]['Volume'].sum()
        print(f"NTB raw volume in 'data thô' (matched to Cơ cấu): {ntb_raw_vol}")
        
        # 3. Volume in Excel DB sheet (sum of all NTB rows for this date)
        # In Excel DB, Time is formatted as "YYYY-MM-DD - Thứ X"
        excel_rows = df_excel_db[df_excel_db['Time'].astype(str).str.startswith(d)]
        # Sum of non-Grand Total rows
        ntb_excel_rows = excel_rows[excel_rows['Chi tiết'] != 'Grand Total']
        ntb_excel_vol = ntb_excel_rows['Volume'].sum()
        print(f"Sum of NTB rows in Excel DB sheet: {ntb_excel_vol}")
        
        # 4. Grand Total row in Excel DB sheet
        gt_row = excel_rows[excel_rows['Chi tiết'] == 'Grand Total']
        if not gt_row.empty:
            print(f"Grand Total row volume in Excel DB sheet: {gt_row['Volume'].sum()}")
            
if __name__ == "__main__":
    main()
