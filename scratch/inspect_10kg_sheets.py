import os
import sys
import io
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
SHEET_KEY = '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    for title in ['trên10kg', 'SL > 10kg']:
        ws = sh.worksheet(title)
        data = ws.get_all_values()
        print(f"\n================ Worksheet: {title} ================")
        print(f"Total Rows: {len(data)}, Total Cols: {len(data[0]) if data else 0}")
        if data:
            print("First 10 rows:")
            for i, r in enumerate(data[:10]):
                print(f"Row {i+1}: {r[:15]}")

if __name__ == "__main__":
    main()
