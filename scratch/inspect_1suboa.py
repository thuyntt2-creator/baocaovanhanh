import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'
    sh = gc_client.open_by_key(key)
    print(f"Spreadsheet '{sh.title}':")
    for ws in sh.worksheets():
        values = ws.get_values("A1:J2")
        print(f"  - Tab: '{ws.title}' (ID: {ws.id})")
        print(f"    Header: {values[0] if len(values) > 0 else 'EMPTY'}")
        if len(values) > 1:
            print(f"    Row 1: {values[1]}")

if __name__ == "__main__":
    inspect()
