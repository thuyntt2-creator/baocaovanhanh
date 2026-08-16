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
    
    print(f"Opened spreadsheet: {sh.title}")
    
    ws_opr = sh.worksheet("OPR")
    raw_values = ws_opr.get_all_values()
    print(f"Row count of 'OPR': {len(raw_values)}")
    if len(raw_values) < 2:
        print("Empty worksheet.")
        return
        
    df = pd.DataFrame(raw_values[1:], columns=raw_values[0])
    
    print("Columns:")
    print(list(df.columns))
    
    if 'NgayLTC' in df.columns:
        print("\nUnique values of 'NgayLTC':")
        print(df['NgayLTC'].value_counts())
    else:
        print("\n'NgayLTC' column not found!")
        # Print first few columns and rows
        print(df.iloc[:5, :5])
        
    if 'Khung giờ tạo' in df.columns:
        print("\nUnique values of 'Khung giờ tạo':")
        print(df['Khung giờ tạo'].value_counts())

    if 'khung_gio_tao' in df.columns:
        print("\nUnique values of 'khung_gio_tao':")
        print(df['khung_gio_tao'].value_counts())

if __name__ == "__main__":
    main()
