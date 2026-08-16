import os
import io
import sys
import gspread
import pandas as pd
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

def check_dates():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    
    for tab_name in ["gtc", "raw", "Data"]:
        ws = sh.worksheet(tab_name)
        data = ws.get_all_values()
        df = pd.DataFrame(data[1:], columns=data[0])
        dates = sorted(df['Time'].unique())
        print(f"Tab {tab_name}: count={len(df)}, min_date={dates[0] if dates else 'None'}, max_date={dates[-1] if dates else 'None'}")
        print(f"  Latest 3 dates: {dates[-3:] if dates else []}")

if __name__ == "__main__":
    check_dates()
