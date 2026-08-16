import os
import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_PATH = r'C:\Users\lap4all\Documents\Auto report\credentials.json'
SPREADSHEET_ID   = '15Z-aMM6OFfiWUXd2Zwz6BFNq_Y0KWwHiVDqxkioHufM'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet('CoCauVung')
    data = ws.get_all_values()
    print("Number of rows:", len(data))
    if len(data) > 0:
        print("Headers:", [str(c) for c in data[0]])
    if len(data) > 1:
        print("Row 1:", [str(c) for c in data[1]])
        print("Row 2:", [str(c) for c in data[2]])

if __name__ == '__main__':
    main()
