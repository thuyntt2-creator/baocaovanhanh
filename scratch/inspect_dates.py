import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_dates():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sheet_key = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
    sh = gc_client.open_by_key(sheet_key)
    
    # Check 'gtc' tab
    print("=== 'gtc' TAB ===")
    ws_gtc = sh.worksheet("gtc")
    df_gtc = pd.DataFrame(ws_gtc.get_all_values())
    if not df_gtc.empty:
        df_gtc.columns = df_gtc.iloc[0]
        df_gtc = df_gtc.iloc[1:]
        print("Columns:", df_gtc.columns.tolist())
        if 'Time' in df_gtc.columns:
            dates = df_gtc['Time'].unique()
            print(f"Number of unique dates: {len(dates)}")
            print("Sorted unique dates (sample):", sorted(list(dates))[-15:])
            
    # Check 'raw' tab
    print("\n=== 'raw' TAB ===")
    try:
        ws_raw = sh.worksheet("raw")
        df_raw = pd.DataFrame(ws_raw.get_all_values())
        if not df_raw.empty:
            df_raw.columns = df_raw.iloc[0]
            df_raw = df_raw.iloc[1:]
            print("Columns:", df_raw.columns.tolist())
            if 'Time' in df_raw.columns:
                dates = df_raw['Time'].unique()
                print(f"Number of unique dates: {len(dates)}")
                print("Sorted unique dates (sample):", sorted(list(dates))[-15:])
    except Exception as e:
        print(f"Error checking 'raw': {e}")

if __name__ == "__main__":
    inspect_dates()
