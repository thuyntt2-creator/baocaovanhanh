import os
import sys
import io
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # Check 'Data Khác', 'Data Shopee', 'Data TTS' worksheets for "Chưa gán AM" or empty/errors
    categories = ['Khác', 'Shopee', 'TTS']
    for cat in categories:
        ws = sh.worksheet(f"Data {cat}")
        rows = ws.get_all_values()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        
        # Print column headers
        print(f"\nTab 'Data {cat}': columns are {list(df.columns)}")
        
        # Find if any cell in any row has "Chưa gán AM"
        # Let's inspect column 'AM' (if it exists) or search all cells
        unassigned_cells = ws.findall("Chưa gán AM")
        if unassigned_cells:
            print(f"  ❌ Found {len(unassigned_cells)} occurrences of 'Chưa gán AM':")
            for c in unassigned_cells[:10]:
                print(f"    - Row {c.row}, Col {c.col} (value: {c.value})")
                row_val = ws.row_values(c.row)
                print(f"      Full row: {row_val}")
        else:
            print("  ✅ No 'Chưa gán AM' found.")
            
        # Check for #N/A cells
        na_cells = ws.findall("#N/A")
        if na_cells:
            print(f"  ⚠️ Found {len(na_cells)} occurrences of '#N/A':")
            for c in na_cells[:10]:
                print(f"    - Row {c.row}, Col {c.col} (value: {c.value})")
        else:
            print("  ✅ No '#N/A' found.")

if __name__ == "__main__":
    main()
