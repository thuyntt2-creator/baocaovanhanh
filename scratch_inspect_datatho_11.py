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
    # 1. Connect to Google Sheets to load data thô
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_raw = sh.worksheet("data thô")
    raw_vals = ws_raw.get_all_values()
    df_raw = pd.DataFrame(raw_vals[1:], columns=raw_vals[0])
    df_raw['Volume'] = pd.to_numeric(df_raw['Volume'], errors='coerce').fillna(0)
    df_raw['WarehouseID'] = df_raw['WarehouseID'].str.strip()
    
    # Top 20 in data thô on 2026-06-11
    df_raw_11 = df_raw[df_raw['Ngay'] == '2026-06-11']
    top_raw = df_raw_11.groupby('WarehouseID')['Volume'].sum().reset_index()
    top_raw = top_raw.sort_values(by='Volume', ascending=False)
    
    # 2. Load Excel DB
    excel_path = os.path.join(BASE_DIR, 'manual_raw.xlsx')
    df_excel_db = pd.read_excel(excel_path, sheet_name='DB')
    df_excel_11 = df_excel_db[df_excel_db['Time'].astype(str).str.startswith('2026-06-11')]
    df_excel_11 = df_excel_11[df_excel_11['Chi tiết'] != 'Grand Total']
    top_excel = df_excel_11.groupby('Chi tiết')['Volume'].sum().reset_index()
    top_excel = top_excel.sort_values(by='Volume', ascending=False)
    
    print("\n--- Top 20 in 'data thô' on 2026-06-11 ---")
    print(top_raw.head(20).to_string(index=False))
    
    print("\n--- Top 20 in Excel 'DB' on 2026-06-11 ---")
    print(top_excel.head(20).to_string(index=False))

if __name__ == "__main__":
    main()
