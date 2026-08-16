import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    ws_opr = sh.worksheet("OPR")
    
    # Check formulas in row 2
    row2_vals = ws_opr.get_values("A2:L2", value_render_option="FORMULA")
    print(f"Row 2 formulas/values: {row2_vals}")
    
    # Check formulas in row 100
    row100_vals = ws_opr.get_values("A100:L100", value_render_option="FORMULA")
    print(f"Row 100 formulas/values: {row100_vals}")
    
    # Check cell A2 specifically
    a2_cell = ws_opr.cell(2, 1, value_render_option="FORMULA")
    print(f"Cell A2 value/formula: {a2_cell.value}")

if __name__ == "__main__":
    main()
