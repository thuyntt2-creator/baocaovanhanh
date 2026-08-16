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
    
    ws_cc = sh.worksheet("Cơ cấu")
    cc_vals = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
    
    target_id = '20144000'
    match = df_cc[df_cc['Mã bưu cục'].str.strip() == target_id]
    print(f"\nSearching for '{target_id}' in 'Cơ cấu' Mã bưu cục:")
    if len(match) > 0:
        print(match.to_string())
    else:
        print("Not found.")
        
    # Let's also print all rows in Cơ cấu that match Ninh Thuận or Bình Thuận to see what they look like
    print("\nSample rows in Cơ cấu for Bình Thuận:")
    print(df_cc[df_cc['Tỉnh'] == 'Bình Thuận'].head(10).to_string())

if __name__ == "__main__":
    main()
