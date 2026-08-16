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
    ws = sh.worksheet("Data")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Filter for row where Volume has dot or comma
    print("Example rows where Volume or GTC might have separators:")
    count = 0
    for idx, row in df.iterrows():
        vol = str(row['Volume'])
        gtc = str(row['Sản Lượng Giao Thành Công'])
        if '.' in vol or ',' in vol or '.' in gtc or ',' in gtc:
            print(f"Row {idx}: Vol={vol}, GTC={gtc}, %GTC={row['% GTC']}, Time={row['Time']}, Detail={row['Chi tiết']}")
            count += 1
            if count >= 10:
                break

if __name__ == "__main__":
    run()
