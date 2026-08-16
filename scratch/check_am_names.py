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
    
    ws = sh.worksheet("Cơ cấu")
    rows = ws.get_all_values()
    df = pd.DataFrame(rows[1:], columns=rows[0])
    
    print("AM values in 'Cơ cấu' tab:")
    print(df['AM'].unique())
    print("\nEmpty or unassigned rows in 'Cơ cấu' tab:")
    empty_rows = df[df['AM'].str.strip() == '']
    print(f"Empty rows count: {len(empty_rows)}")
    for idx, r in empty_rows.iterrows():
        print(f"Row {idx+2}: {dict(r)}")
        
if __name__ == "__main__":
    main()
