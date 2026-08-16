import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    
    ws_gtc = sh.worksheet("gtc")
    rows = ws_gtc.get_all_values()
    print(f"Total rows in sheet tab: {len(rows)}")
    print("First 3 rows:")
    for r in rows[:3]:
        print(r)
    print("\nRows matching Cam Linh on 2026-07-04:")
    for r in rows:
        if "Cam Linh" in str(r[1]) and "2026-07-04" in str(r[3]):
            print([cell for cell in r])
            print([type(cell) for cell in r])

if __name__ == "__main__":
    inspect()
