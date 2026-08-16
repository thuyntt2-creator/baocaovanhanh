import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(MAIN_SPREADSHEET_ID)
    
    ws_fd = sh.worksheet("FD ")
    fd_rows = ws_fd.get_all_values()
    
    print("List of all raw values in 'Bưu Cục' column of the FD sheet:")
    for idx, r in enumerate(fd_rows[3:]):
        if r and len(r) > 0:
            print(f"Row {idx+4}: '{r[0]}'")

if __name__ == "__main__":
    main()
