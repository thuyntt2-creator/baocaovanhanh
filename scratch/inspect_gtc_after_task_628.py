import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("gtc")
    
    rows = ws.get_all_values()
    print("Total rows:", len(rows))
    
    print("\nRow 2 (Đức Lập):")
    r2_unformatted = ws.row_values(2, value_render_option='UNFORMATTED_VALUE')
    print("Values:", r2_unformatted)
    print("Types :", [type(x) for x in r2_unformatted])
    
    # Find Cam Linh
    for idx, r in enumerate(rows):
        if "Cam Linh" in r[1] and "2026-07-04" in r[3] and "Hàng Mới Ca 1" in r[2]:
            row_idx = idx + 1
            print(f"\nCam Linh Row {row_idx}:")
            r_unformatted = ws.row_values(row_idx, value_render_option='UNFORMATTED_VALUE')
            print("Values:", r_unformatted)
            print("Types :", [type(x) for x in r_unformatted])
            break

if __name__ == "__main__":
    inspect()
