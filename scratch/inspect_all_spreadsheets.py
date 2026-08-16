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
    '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk': 'report_tongquan key',
    '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU': '1WCz key',
    '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M': '1l2j key',
    '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg': '1Pjz key',
    '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ': '1JZ (Dash Board)'
}

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key, name in keys.items():
        print(f"\n=========================================")
        print(f"Key: {key} ({name})")
        try:
            sh = gc_client.open_by_key(key)
            print(f"Spreadsheet Title: {sh.title}")
            metadata = sh.fetch_sheet_metadata()
            properties = metadata.get('properties', {})
            print(f"Locale: {properties.get('locale')}")
            worksheets = sh.worksheets()
            print("Worksheets:")
            for ws in worksheets:
                print(f"  - {ws.title}")
        except Exception as e:
            print(f"Error opening sheet: {e}")

if __name__ == "__main__":
    main()
