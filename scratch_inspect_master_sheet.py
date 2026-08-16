import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding cho Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
MASTER_SHEET_KEY = '1PIyzade3_ml9Zq8OwTD5WGrc-paZJId7DfB06qSWJpg'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_master():
    print(f"📖 Connecting to master sheet: {MASTER_SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    try:
        sh = gc_client.open_by_key(MASTER_SHEET_KEY)
        print("Success!")
        print("Worksheets:")
        for ws in sh.worksheets():
            print(f"  - '{ws.title}' (ID: {ws.id})")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    inspect_master()
