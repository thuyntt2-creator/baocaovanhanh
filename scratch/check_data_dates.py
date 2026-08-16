import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    
    # We want to check sheet "Data"
    ws = sh.worksheet("Data")
    print(f"Reading all values from {ws.title}...")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    print(f"Loaded {len(df)} rows and columns: {df.columns.tolist()}")
    
    # Print distinct values in 'Time'
    print("\nDistinct dates in 'Time' column:")
    dates = df['Time'].unique()
    print(sorted(dates))
    
    # Print distinct values in 'Loại Hàng'
    print("\nDistinct values in 'Loại Hàng' column:")
    print(df['Loại Hàng'].unique())

if __name__ == "__main__":
    inspect()
