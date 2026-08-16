import sys
import io
import os
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    new_key = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
    print(f"--- Checking AM sheet in {new_key} ---")
    try:
        sh = gc_client.open_by_key(new_key)
        ws = sh.worksheet("Lê Minh Đại")
        rows = ws.get_all_values()
        print("Total rows in Lê Minh Đại sheet:", len(rows))
        if len(rows) > 1:
            header = rows[0]
            df = pd.DataFrame(rows[1:], columns=header)
            print("Unique 'Nhóm BL' values in Lê Minh Đại sheet:")
            print(df['Nhóm BL'].value_counts())
            print("\nFirst 5 rows:")
            print(df[['bc', 'order_code', 'Aging', 'Nhóm BL']].head())
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
