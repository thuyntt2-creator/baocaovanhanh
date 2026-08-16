import os
import sys
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials as UserCredentials

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = r"c:\Users\lap4all\Documents\Auto report"
GOOGLE_SHEET_KEY = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
auth_file = os.path.join(ROOT_DIR, 'authorized_user.json')
creds = UserCredentials.from_authorized_user_file(auth_file, scopes=scopes)
gc = gspread.authorize(creds)
sh = gc.open_by_key(GOOGLE_SHEET_KEY)

print("✅ Connecting using authorized_user.json succeeded!")
print("Sheet title:", sh.title)

ws_gtc = sh.worksheet("gtc")
ws_data = sh.worksheet("Data")

print("gtc rows count:", ws_gtc.row_count)
print("Data rows count:", ws_data.row_count)
