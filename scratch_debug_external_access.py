import os
import sys
import io
import traceback
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
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
EXTERNAL_SHEET_KEY = '1rhBD3-QSxn2c3yYNnyuj-iaW26H6WIH4G5FKb5Iv394'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print(f"Connecting to external Google Sheet: {EXTERNAL_SHEET_KEY}...")
    try:
        creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
        gc_client = gspread.authorize(creds)
        sh = gc_client.open_by_key(EXTERNAL_SHEET_KEY)
        print(f"Successfully opened sheet: '{sh.title}'")
    except Exception as e:
        print(f"Failed to access external sheet. Exception class: {e.__class__.__name__}")
        print("Traceback:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
