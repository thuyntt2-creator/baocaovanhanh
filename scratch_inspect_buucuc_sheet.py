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

def inspect_sheet(sh, name):
    print(f"\n======================================")
    print(f"Inspecting sheet: '{name}'")
    print(f"======================================")
    try:
        ws = sh.worksheet(name)
        vals = ws.get_values("A1:Z30") # Get first 30 rows and 26 columns
        if not vals:
            print("Empty sheet.")
            return
        
        # Display first 15 rows
        for i, row in enumerate(vals[:15]):
            # Filter out trailing empty strings to make it clean
            clean_row = [x for x in row if x != '']
            if clean_row:
                print(f"Row {i+1:2d}: {clean_row[:15]}")
            else:
                print(f"Row {i+1:2d}: [Empty]")
    except Exception as e:
        print(f"Error reading '{name}': {e}")

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    inspect_sheet(sh, 'Bưu cục')
    inspect_sheet(sh, 'Trang tính12')
    inspect_sheet(sh, 'Trang tính13')

if __name__ == "__main__":
    main()
