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

print("\nSearching for new recruitment sheet...")
for ws in spreadsheet.worksheets():
    title_lower = ws.title.lower()
    if 'tuyển' in title_lower or 'tuyen' in title_lower or 'recru' in title_lower or 'báo cáo' in title_lower or 'bao cao' in title_lower:
        print(f"\n=== INSPECTING SHEET: {ws.title} ===")
        vals = ws.get_all_values()
        df = pd.DataFrame(vals)
        print("Shape:", df.shape)
        print("First 15 rows:")
        print(df.head(15).to_string())

