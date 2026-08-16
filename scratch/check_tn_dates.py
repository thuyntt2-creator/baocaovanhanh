import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SHEET_ID)

tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

print("Unique dates in column 'Thời gian':", df_tn['Thời gian'].unique().tolist())
print(df_tn.head(10).to_string())
