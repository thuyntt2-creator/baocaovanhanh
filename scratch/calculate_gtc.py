import os
import io
import sys
import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def clean_and_int(val):
    if not val or pd.isna(val):
        return 0
    val_str = str(val).strip().replace('.', '').replace(',', '')
    try:
        return int(val_str)
    except ValueError:
        return 0

def clean_and_float(val):
    if not val or pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('%', '')
    if ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def run():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    ws = sh.worksheet("Data")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Let's see the unique dates in sorted order
    all_dates = sorted(df['Time'].unique())
    last_7_dates = all_dates[-7:]
    print("Latest 7 dates being analyzed:")
    for d in last_7_dates:
        print(f"  - {d}")
        
    # Filter df for only these 7 dates
    df_7d = df[df['Time'].isin(last_7_dates)].copy()
    
    # Convert Volume and Sản Lượng Giao Thành Công to integers
    df_7d['Volume_int'] = df_7d['Volume'].apply(clean_and_int)
    df_7d['GTC_int'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_and_int)
    
    print(f"\nFiltered data shape: {df_7d.shape}")
    print(f"Total Volume in 7 days: {df_7d['Volume_int'].sum()}")
    print(f"Total GTC in 7 days: {df_7d['GTC_int'].sum()}")
    
    # Aggregate by post office ('Chi tiết')
    grouped = df_7d.groupby('Chi tiết').agg({
        'Volume_int': 'sum',
        'GTC_int': 'sum'
    }).reset_index()
    
    grouped['GTC_rate'] = grouped['GTC_int'] / grouped['Volume_int']
    
    # Filter for GTC rate < 50%
    low_gtc = grouped[grouped['GTC_rate'] < 0.50].sort_values(by='GTC_rate')
    
    print(f"\nNumber of post offices with GTC < 50% in last 7 days: {len(low_gtc)}")
    for idx, row in low_gtc.iterrows():
        rate_pct = row['GTC_rate'] * 100
        print(f"{row['Chi tiết']}: GTC = {row['GTC_int']:,} / Vol = {row['Volume_int']:,} ({rate_pct:.2f}%)")

if __name__ == "__main__":
    run()
