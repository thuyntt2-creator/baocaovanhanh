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
    
    ws = sh.worksheet('data thô')
    print("Fetching 'Ngay' values...")
    data = ws.get_all_values()
    headers = data[0]
    ngay_idx = headers.index('Ngay')
    
    dates = [row[ngay_idx] for row in data[1:] if row[ngay_idx]]
    print(f"Total non-empty dates: {len(dates)}")
    print(f"Sample dates (first 10): {dates[:10]}")
    print(f"Sample dates (last 10): {dates[-10:]}")
    
    # Check unique date formats
    formats = {}
    for d in dates:
        fmt = 'other'
        if '-' in d:
            parts = d.split('-')
            if len(parts) == 3:
                fmt = 'YYYY-MM-DD'
        elif '/' in d:
            fmt = 'slash'
        elif d.isdigit():
            fmt = 'digit'
        formats[fmt] = formats.get(fmt, 0) + 1
        
    print(f"Date formats found: {formats}")

if __name__ == "__main__":
    main()
