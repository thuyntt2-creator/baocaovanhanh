import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

print("=== SHEET: data ===")
ws_data = spreadsheet.worksheet("data")
data_vals = ws_data.get_all_values()
df_data = pd.DataFrame(data_vals)
print("Data shape:", df_data.shape)
print("First 15 rows of 'data':")
print(df_data.head(15).to_string())

print("\n=== SHEET: thu nhập ===")
ws_tn = spreadsheet.worksheet("thu nhập")
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals)
print("Thu nhập shape:", df_tn.shape)
print("First 15 rows of 'thu nhập':")
print(df_tn.head(15).to_string())

