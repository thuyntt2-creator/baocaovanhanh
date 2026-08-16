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
    
    # Print headers (row 1)
    print("Row 0 (Header):")
    row0 = df.iloc[0].tolist()
    for i, h in enumerate(row0):
        if h != '':
            print(f"  Col {i}: {h}")
            
    print("\nRow 1 (Subheader):")
    row1 = df.iloc[1].tolist()
    for i, h in enumerate(row1):
        if h != '':
            print(f"  Col {i}: {h}")

    # Inspect Row 11 (which starts index 0)
    print("\nRow 11 Detail:")
    r11 = df.iloc[11].tolist()
    for i, val in enumerate(r11):
        print(f"  Col {i:2d}: {val}")

    # Inspect Row 74 (which starts index 0)
    print("\nRow 74 Detail:")
    r74 = df.iloc[74].tolist()
    for i, val in enumerate(r74):
        print(f"  Col {i:2d}: {val}")

if __name__ == "__main__":
    run()
