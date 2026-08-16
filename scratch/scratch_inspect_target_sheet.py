import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SHEET_KEY = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'

def list_tabs():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    try:
        sh = gc_client.open_by_key(SHEET_KEY)
        print(f"Spreadsheet '{sh.title}' ({SHEET_KEY}):")
        for ws in sh.worksheets():
            print(f"  - {ws.title} (ID: {ws.id})")
    except Exception as e:
        print(f"Lỗi mở key {SHEET_KEY}: {e}")

if __name__ == "__main__":
    list_tabs()
