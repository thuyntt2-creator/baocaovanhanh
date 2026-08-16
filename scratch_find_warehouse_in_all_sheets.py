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
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    target_val = '21624000'
    print(f"Searching for '{target_val}' in all worksheets...")
    
    for ws in sh.worksheets():
        try:
            # We can search by fetching cell values
            cells = ws.findall(target_val)
            if cells:
                print(f"Found in sheet '{ws.title}': {len(cells)} times")
                for cell in cells[:5]:
                    print(f"  Row {cell.row}, Col {cell.col}")
            else:
                pass
        except Exception as e:
            print(f"Error searching in sheet '{ws.title}': {e}")

if __name__ == "__main__":
    main()
