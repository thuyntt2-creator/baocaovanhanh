import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
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
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("Cơ cấu")
    rows = ws.get_all_values()
    print("Search results in Cơ cấu for 'AM Chi':")
    for r in rows:
        if any("AM Chi" in str(cell) for cell in r):
            print(r)
    print("Search results in Cơ cấu for 'AM Phước':")
    for r in rows:
        if any("AM Phước" in str(cell) for cell in r):
            print(r)

if __name__ == "__main__":
    inspect()
