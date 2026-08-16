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
    
    for sname in ['DB', 'Data', 'dbtts', 'raw']:
        try:
            ws = sh.worksheet(sname)
            vals = ws.get_all_values()
            df = pd.DataFrame(vals[1:], columns=vals[0])
            print(f"\n=== Sheet '{sname}' ===")
            print(f"  Rows: {len(df)}")
            if len(df) > 0:
                # Unique dates/Time
                time_col = 'Time' if 'Time' in df.columns else df.columns[3]
                unique_times = df[time_col].dropna().unique()
                print(f"  Unique times count: {len(unique_times)}")
                print(f"  Sample times: {list(unique_times)[:5]}")
                print(f"  Latest times: {list(unique_times)[-5:]}")
                # Sum of volume
                vol_col = 'Volume' if 'Volume' in df.columns else df.columns[4]
                df[vol_col] = pd.to_numeric(df[vol_col], errors='coerce').fillna(0)
                print(f"  Total Volume: {df[vol_col].sum()}")
                
                # Check for 2026-06-25
                df_25 = df[df[time_col].str.startswith('2026-06-25', na=False)]
                print(f"  Rows for 2026-06-25: {len(df_25)}")
                print(f"  Volume for 2026-06-25: {df_25[vol_col].sum()}")
        except Exception as e:
            print(f"  Error reading sheet '{sname}': {e}")

if __name__ == "__main__":
    main()
