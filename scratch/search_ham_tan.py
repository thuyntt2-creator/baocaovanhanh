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

def run():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    ws = sh.worksheet("Bưu cục")
    data = ws.get_all_values()
    df = pd.DataFrame(data)
    
    print("Searching for 'Hàm Tân' in worksheet 'Bưu cục':")
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row):
            if 'Hàm Tân' in str(val):
                print(f"Row {r_idx}, Col {c_idx}: {val}")
                # Print neighboring cells in the same row
                # Let's print cells around it
                start_c = max(0, c_idx - 2)
                end_c = min(len(row), c_idx + 10)
                print(f"  Row cells {start_c} to {end_c-1}: {list(row[start_c:end_c])}")

if __name__ == "__main__":
    run()
