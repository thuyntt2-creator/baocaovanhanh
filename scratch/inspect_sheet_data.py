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

# We find credentials.json in the parent folder
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_sheet():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    print(f"Spreadsheet: {sh.title}")
    
    target_ws = None
    for ws in sh.worksheets():
        print(f"Tab: {ws.title} (ID: {ws.id})")
        if str(ws.id) == '324939847':
            target_ws = ws
            
    if not target_ws:
        print("Tab with ID 324939847 not found, checking by title or using first tab")
        target_ws = sh.get_worksheet(0)
        
    print(f"Using worksheet: {target_ws.title}")
    data = target_ws.get_all_values()
    df = pd.DataFrame(data)
    print("DataFrame shape:", df.shape)
    print("First 20 rows:")
    print(df.head(20))
    
    # Save a CSV of the data to scratch directory for further investigation if needed
    csv_path = os.path.join(SCRIPT_DIR, 'gtc_inspect.csv')
    df.to_csv(csv_path, index=False, encoding='utf-8')
    print(f"Saved first 100 rows to {csv_path}")

if __name__ == "__main__":
    inspect_sheet()
