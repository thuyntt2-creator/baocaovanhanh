import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_data = sh.worksheet("Data")
    # Get values with formulas
    formulas = ws_data.get_all_values(value_render_option='FORMULA')
    print("Formulas in Data sheet (first 15 rows):")
    for idx, row in enumerate(formulas[:15]):
        print(f"{idx}: {row}")

if __name__ == "__main__":
    main()
