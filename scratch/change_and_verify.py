import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def test():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("gtc")
    
    original_val = ws.acell('A2').value
    print(f"Original value of A2: '{original_val}'")
    
    print("Writing 'TEST_WRITE' to A2...")
    ws.update_acell('A2', 'TEST_WRITE')
    
    new_val = ws.acell('A2').value
    print(f"New value of A2: '{new_val}'")
    
    # Restore original value
    print("Restoring original value...")
    ws.update_acell('A2', original_val)

if __name__ == "__main__":
    test()
