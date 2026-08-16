import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    ws_rawopr = sh.worksheet("rawopr")
    raw_values = ws_rawopr.get_all_values()
    print(f"Row count of 'rawopr': {len(raw_values)}")
    if len(raw_values) < 2:
        print("Empty worksheet.")
        return
        
    df = pd.DataFrame(raw_values[1:], columns=raw_values[0])
    print("Columns in 'rawopr':")
    print(list(df.columns))
    
    date_col = next((c for c in df.columns if 'ngay' in c.lower() or 'date' in c.lower()), None)
    if date_col:
        print(f"\nUnique values of '{date_col}' in 'rawopr':")
        print(df[date_col].value_counts())
    else:
        print("\nNo date column found in 'rawopr'!")
        print(df.iloc[:5, :5])

if __name__ == "__main__":
    main()
