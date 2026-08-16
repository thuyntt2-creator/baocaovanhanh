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
    
    ws_opr = sh.worksheet("OPR")
    raw_values = ws_opr.get_all_values()
    headers = raw_values[0]
    
    df = pd.DataFrame(raw_values[1:], columns=headers)
    df_26 = df[df['NgayLTC'] == '2026-06-26']
    
    print(f"Number of rows for 2026-06-26: {len(df_26)}")
    if len(df_26) > 0:
        print("\nFirst 5 rows of 2026-06-26:")
        print(df_26.head(5).to_string())
        
        print("\nChecking for nulls/empty strings in each column for 2026-06-26:")
        for col in df_26.columns:
            empty_count = (df_26[col] == '').sum()
            null_count = df_26[col].isnull().sum()
            print(f"Column '{col}': {empty_count} empty, {null_count} nulls")
            
        print("\nDistinct values in 'khung_gio_tao' for 2026-06-26:")
        print(df_26['khung_gio_tao'].value_counts())
        
        print("\nDistinct values in 'Khung giờ tạo' for 2026-06-26:")
        print(df_26['Khung giờ tạo'].value_counts())
        
        print("\nDistinct values in 'AM' for 2026-06-26:")
        print(df_26['AM'].value_counts())

if __name__ == "__main__":
    main()
