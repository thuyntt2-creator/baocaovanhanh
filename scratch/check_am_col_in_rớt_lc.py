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
    
    categories = ['Khác', 'Shopee', 'TTS']
    for cat in categories:
        ws = sh.worksheet(f"Data {cat}")
        rows = ws.get_all_values()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        
        # Check column 'Quản lý'
        if 'Quản lý' in df.columns:
            unassigned_rows = df[df['Quản lý'].str.contains("Chưa gán|chua gan|không|khong|N/A", case=False, na=True) | (df['Quản lý'].str.strip() == "")]
            print(f"\nTab 'Data {cat}': found {len(unassigned_rows)} unassigned/empty rows in 'Quản lý' column. Unique values:")
            print(df['Quản lý'].unique())
            if len(unassigned_rows) > 0:
                print("First 10 unassigned rows:")
                for idx, r in unassigned_rows.head(10).iterrows():
                    print(f"  - Row {idx+2}: {dict(r)}")
        else:
            print(f"\nTab 'Data {cat}' does not have 'Quản lý' column!")

if __name__ == "__main__":
    main()
