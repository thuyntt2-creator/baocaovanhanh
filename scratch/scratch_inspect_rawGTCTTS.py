import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SHEET_KEY = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws = sh.worksheet("rawGTCTTS")
    print("--- rawGTCTTS First 3 rows ---")
    data = ws.get_values("A1:M4")
    for r in data:
        print(r)
        
    print("\n--- Cocau First 3 rows ---")
    ws_cc = sh.worksheet("Cơ cấu")
    cc_data = ws_cc.get_values("A1:G4")
    for r in cc_data:
        print(r)

if __name__ == "__main__":
    inspect()
