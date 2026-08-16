import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')

keys = {
    '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4': '2026 NTB - BÁO CÁO VẬN HÀNH',
    '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk': 'NTB - BÁO CÁO VẬN HÀNH',
    '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU': 'Aging >5 ngày',
    '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M': 'Aging >5 ngày - follow gán',
    '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg': 'NTB- FOLLOW OFF Tuyến',
    '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ': 'Dash Board'
}

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key, name in keys.items():
        try:
            sh = gc_client.open_by_key(key)
            titles = [ws.title for ws in sh.worksheets()]
            if 'raw_data' in titles:
                print(f"✅ Found 'raw_data' in spreadsheet '{sh.title}' ({key})")
            if 'raw' in titles:
                print(f"✅ Found 'raw' in spreadsheet '{sh.title}' ({key})")
        except Exception as e:
            print(f"Error opening {key}: {e}")

if __name__ == "__main__":
    main()
