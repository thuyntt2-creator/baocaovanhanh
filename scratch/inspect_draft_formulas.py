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
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws = sh.worksheet('DRAFT')
    print("Worksheet: DRAFT")
    
    # Read headers
    headers = ws.get_values('A1:Z1')[0]
    print(f"Headers: {headers}")
    
    # Read Row 2 values
    row2_val = ws.get_values('A2:Z2')[0]
    print(f"Row 2 values: {row2_val}")
    
    # Read Row 2 formulas
    row2_form = ws.get_values('A2:Z2', value_render_option='FORMULA')[0]
    print(f"Row 2 formulas: {row2_form}")
    
    for idx in range(len(row2_val)):
        header = headers[idx] if idx < len(headers) else f"Col {idx+1}"
        val = row2_val[idx]
        form = row2_form[idx] if idx < len(row2_form) else ""
        print(f"Col {idx+1} ({header}): Value='{val}', Formula='{form}'")

if __name__ == "__main__":
    main()
