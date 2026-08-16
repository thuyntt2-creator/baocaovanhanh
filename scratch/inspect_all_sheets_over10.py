import sys
import io
import os
import gspread
import unicodedata
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
    print(f"--- Inspecting all worksheets in {new_key} ---")
    try:
        sh = gc_client.open_by_key(new_key)
        worksheets = sh.worksheets()
        print(f"Total worksheets found: {len(worksheets)}")
        
        for ws in worksheets:
            title = ws.title
            nfc_title = unicodedata.normalize('NFC', title)
            nfd_title = unicodedata.normalize('NFD', title)
            is_nfc_equal_nfd = (nfc_title == nfd_title)
            
            # Read all rows
            rows = ws.get_all_values()
            row_count = len(rows)
            
            print(f"\nWorksheet: '{title}' (ID: {ws.id})")
            print(f"  NFC equal NFD: {is_nfc_equal_nfd} | NFC hex: {title.encode('utf-8').hex()}")
            print(f"  Total rows: {row_count}")
            
            if row_count > 1:
                header = rows[0]
                df = pd.DataFrame(rows[1:], columns=header)
                if 'Nhóm BL' in df.columns:
                    print("  Unique 'Nhóm BL':", df['Nhóm BL'].value_counts().to_dict())
                else:
                    print("  'Nhóm BL' column not found! Available columns:", list(df.columns)[:5])
            elif row_count == 1:
                print("  Only header row exists.")
            else:
                print("  Worksheet is empty.")
                
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
