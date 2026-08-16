import os
import io
import sys
import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def clean_pct(val):
    if not val:
        return 0.0
    val_str = str(val).strip().replace('%', '')
    val_str = val_str.replace(',', '.')
    try:
        num = float(val_str)
        if '%' in str(val) or num > 1.0:
            return num / 100.0
        return num
    except ValueError:
        return 0.0

def clean_volume(val):
    if not val:
        return 0
    val_str = str(val).strip().replace('.', '').replace(',', '')
    try:
        return int(val_str)
    except ValueError:
        try:
            return int(float(val_str))
        except ValueError:
            return 0

def check_gtc_data():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sheet_key = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
    sh = gc_client.open_by_key(sheet_key)
    
    # 1. Inspect 'gtc' tab
    print("--- Reading 'gtc' tab ---")
    ws_gtc = sh.worksheet("gtc")
    df_gtc = pd.DataFrame(ws_gtc.get_all_values())
    df_gtc.columns = df_gtc.iloc[0]
    df_gtc = df_gtc.iloc[1:].copy()
    df_gtc.columns = [c.strip() for c in df_gtc.columns]
    
    # Parse Time to extract date
    # Format of Time is "YYYY-MM-DD - Thứ X"
    df_gtc['date_str'] = df_gtc['Time'].apply(lambda x: str(x).split(' ')[0])
    
    all_dates_gtc = sorted(df_gtc['date_str'].unique())
    print(f"gtc dates range from {all_dates_gtc[0]} to {all_dates_gtc[-1]}")
    print(f"Total unique dates in gtc: {len(all_dates_gtc)}")
    print(f"Latest 7 dates in gtc: {all_dates_gtc[-7:]}")
    
    # 2. Inspect 'raw' tab
    print("\n--- Reading 'raw' tab ---")
    try:
        ws_raw = sh.worksheet("raw")
        df_raw = pd.DataFrame(ws_raw.get_all_values())
        df_raw.columns = df_raw.iloc[0]
        df_raw = df_raw.iloc[1:].copy()
        df_raw.columns = [c.strip() for c in df_raw.columns]
        df_raw['date_str'] = df_raw['Time'].apply(lambda x: str(x).split(' ')[0])
        all_dates_raw = sorted(df_raw['date_str'].unique())
        print(f"raw dates range from {all_dates_raw[0]} to {all_dates_raw[-1]}")
        print(f"Total unique dates in raw: {len(all_dates_raw)}")
        print(f"Latest 7 dates in raw: {all_dates_raw[-7:]}")
    except Exception as e:
        print(f"Error checking 'raw': {e}")
        df_raw = pd.DataFrame()

if __name__ == "__main__":
    check_gtc_data()
