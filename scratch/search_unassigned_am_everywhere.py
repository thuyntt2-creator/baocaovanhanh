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
    
    print("Checking worksheets...")
    for ws in sh.worksheets():
        title = ws.title
        # Skip raw data sheet if it has too many rows
        if title == "data rớt LC":
            continue
        try:
            cells = ws.findall("Chưa gán AM")
            if cells:
                print(f"Tab '{title}': Found 'Chưa gán AM' at cells:")
                for cell in cells[:10]:
                    print(f"  - Row {cell.row}, Col {cell.col} (value: {cell.value})")
                if len(cells) > 10:
                    print(f"  ... and {len(cells) - 10} more occurrences.")
        except Exception as e:
            print(f"Error checking tab '{title}': {e}")

if __name__ == "__main__":
    main()
