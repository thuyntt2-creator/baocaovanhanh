import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

ws_rec = spreadsheet.worksheet("báo cáo tuyển dụng")

# Get values with formulas if any
formulas = ws_rec.get_all_values(value_render_option='FORMULA')

print("=== FORMULAS / HEADERS (Row 0, 1, 2) ===")
for r_idx in range(min(5, len(formulas))):
    print(f"\nRow {r_idx}:")
    for c_idx, val in enumerate(formulas[r_idx]):
        if val != "":
            print(f"  Col {c_idx}: {val}")

# Check specific row 10 (Di Linh) formulas
print("\n=== Di Linh (Row 10) ===")
for c_idx, val in enumerate(formulas[9]): # 0-indexed row 9 is row 10 in 1-indexed
    if val != "":
        print(f"  Col {c_idx}: {val}")

# Check specific row 34 (Cam Linh) formulas
print("\n=== Cam Linh (Row 34) ===")
for c_idx, val in enumerate(formulas[33]): # 0-indexed row 33 is row 34 in 1-indexed
    if val != "":
        print(f"  Col {c_idx}: {val}")

