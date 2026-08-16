import sys
sys.stdout.reconfigure(encoding='utf-8')

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

CREDENTIALS_PATH = r'credentials.json'
SPREADSHEET_ID   = '15Z-aMM6OFfiWUXd2Zwz6BFNq_Y0KWwHiVDqxkioHufM'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
gc = gspread.authorize(creds)
spreadsheet = gc.open_by_key(SPREADSHEET_ID)

try:
    ws = spreadsheet.worksheet('CoCauVung')
    data = ws.get_all_values()
    if len(data) >= 1:
        df = pd.DataFrame(data[1:], columns=data[0])
        print("CoCauVung columns:")
        print(df.columns.tolist())
        print("First 5 rows of CoCauVung:")
        print(df.head(5).to_string())
    else:
        print("CoCauVung is empty.")
except Exception as e:
    print(f"Error reading CoCauVung: {e}")
