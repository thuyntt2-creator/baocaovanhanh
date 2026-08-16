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

print("Worksheets:")
for ws in spreadsheet.worksheets():
    print(f"Sheet: {ws.title} | id: {ws.id}")

print("\nSearching for productivity / năng suất sheet...")
for ws in spreadsheet.worksheets():
    t_lower = ws.title.lower()
    if 'năng suất' in t_lower or 'nang suat' in t_lower or 'năng' in t_lower or 'suat' in t_lower:
        print(f"\n=== INSPECTING SHEET: {ws.title} ===")
        vals = ws.get_all_values()
        df = pd.DataFrame(vals)
        print("Shape:", df.shape)
        print("First 15 rows:")
        print(df.head(15).to_string())

