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

print("Total rows:", len(vals))

# Find row for Di Linh and Cam Linh by checking ID or name
di_linh_row = None
cam_linh_row = None

for idx, r in enumerate(vals):
    r_str = " ".join(r)
    if '20942000' in r_str:
        di_linh_row = (idx, r)
    if '22830000' in r_str:
        cam_linh_row = (idx, r)

print("\n--- HEADER ROWS (Row 0 & Row 1) ---")
row0 = vals[0]
row1 = vals[1]
for c in range(max(len(row0), len(row1))):
    v0 = row0[c] if c < len(row0) else ""
    v1 = row1[c] if c < len(row1) else ""
    print(f"Col {c:2d} | Row0: '{v0}' | Row1: '{v1}'")

if di_linh_row:
    print(f"\n--- DI LINH (Row {di_linh_row[0]}) ---")
    for c, val in enumerate(di_linh_row[1]):
        v0 = row0[c] if c < len(row0) else ""
        v1 = row1[c] if c < len(row1) else ""
        print(f"Col {c:2d} | [{v0} / {v1}] => '{val}'")

if cam_linh_row:
    print(f"\n--- CAM LINH (Row {cam_linh_row[0]}) ---")
    for c, val in enumerate(cam_linh_row[1]):
        v0 = row0[c] if c < len(row0) else ""
        v1 = row1[c] if c < len(row1) else ""
        print(f"Col {c:2d} | [{v0} / {v1}] => '{val}'")

