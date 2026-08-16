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

JSON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def list_tabs():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key in ['1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M', '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU']:
        try:
            sh = gc_client.open_by_key(key)
            print(f"Spreadsheet '{sh.title}' ({key}):")
            for ws in sh.worksheets():
                print(f"  - {ws.title} (ID: {ws.id})")
        except Exception as e:
            print(f"Loi mo key {key}: {e}")

if __name__ == "__main__":
    list_tabs()
