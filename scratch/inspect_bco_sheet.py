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

worksheets = spreadsheet.worksheets()
print("Worksheets:", [w.title for w in worksheets])

for ws in worksheets:
    print(f"\n--- Worksheet: {ws.title} (rows: {ws.row_count}, cols: {ws.col_count}) ---")
    data = ws.get_all_values()
    if not data:
        print("Empty sheet")
        continue
    df = pd.DataFrame(data)
    print("Shape:", df.shape)
    print("First 5 rows:")
    print(df.head(5).to_string())
