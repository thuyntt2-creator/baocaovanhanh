import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
SHEET_KEY = '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    print(f"Spreadsheet Title: {sh.title}")
    worksheets = sh.worksheets()
    print("Worksheets:")
    for ws in worksheets:
        print(f"  - {ws.title}")
        
    for ws in worksheets:
        if ws.title in ['raw_data', 'Đơn giao aging trên 5 ngày']:
            print(f"\n--- Checking worksheet: {ws.title} ---")
            try:
                row1 = ws.get_values('A1:Z1', value_render_option='FORMULA')
                print(f"  Row 1 (formulas): {row1}")
                row2 = ws.get_values('A2:Z2', value_render_option='FORMULA')
                print(f"  Row 2 (formulas): {row2}")
                row2_val = ws.get_values('A2:Z2')
                print(f"  Row 2 (values): {row2_val}")
            except Exception as e:
                print(f"  Error reading {ws.title}: {e}")

if __name__ == "__main__":
    main()
