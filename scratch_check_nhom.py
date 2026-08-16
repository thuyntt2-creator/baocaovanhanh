import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
    
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()
    aging_header = aging_data[0]
    print("Full aging headers:", aging_header)
    
    if "Nhóm BL" in aging_header:
        idx = aging_header.index("Nhóm BL")
        values = [row[idx] for row in aging_data[1:] if len(row) > idx]
        unique_vals = set(values)
        print("Unique values in 'Nhóm BL':", unique_vals)
        print("First 10 values in 'Nhóm BL':", values[:10])
    else:
        print("Column 'Nhóm BL' not found in headers!")

if __name__ == '__main__':
    main()
