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
vals = ws_rec.get_all_values()

row0 = vals[0]
row1 = vals[1]

print("=== COLUMN MAPPING IN 'báo cáo tuyển dụng' ===")
for c in range(max(len(row0), len(row1))):
    v0 = row0[c] if c < len(row0) else ""
    v1 = row1[c] if c < len(row1) else ""
    print(f"Col {c:2d}: Row0='{v0}' | Row1='{v1}'")

# Let's inspect rows for Di Linh & Cam Linh again alongside these column headers
for target_id, name in [('20942000', 'Di Linh'), ('22830000', 'Cam Linh')]:
    print(f"\n==================== DATA FOR {target_id} - {name} ====================")
    for idx, r in enumerate(vals):
        if target_id in " ".join(r):
            print(f"Row Index: {idx}")
            for c in range(len(r)):
                val = r[c]
                v0 = row0[c] if c < len(row0) else ""
                v1 = row1[c] if c < len(row1) else ""
                print(f"  Col {c:2d} ({v0} / {v1}): '{val}'")

