import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
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
    
    # Read Data
    ws_data = sh.worksheet("Data")
    data_vals = ws_data.get_all_values()
    df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])
    
    print(f"Data row count: {len(df_data)}")
    print("Data columns:", df_data.columns.tolist())
    print("\nUnique Cấp Quản Lý in Data:")
    print(df_data['Cấp Quản Lý'].value_counts())
    
    print("\nUnique Chi tiết count in Data:", df_data['Chi tiết'].nunique())
    print("Sample unique Chi tiết in Data:")
    print(df_data['Chi tiết'].unique()[:30])
    
    df_data['Volume'] = pd.to_numeric(df_data['Volume'], errors='coerce').fillna(0)
    print("\nTotal Volume in Data:", df_data['Volume'].sum())
    
    # Check if there is data for 2026-06-25 in Data
    # The date is stored in 'Time' column, let's see its format
    print("\nSample Time format in Data:")
    print(df_data['Time'].unique()[:10])
    
    # Filter for Time starting with 2026-06-25
    df_data_25 = df_data[df_data['Time'].str.startswith('2026-06-25')]
    print(f"Rows for 2026-06-25 in Data: {len(df_data_25)}")
    print(f"Total volume for 2026-06-25 in Data: {df_data_25['Volume'].sum()}")

if __name__ == "__main__":
    main()
