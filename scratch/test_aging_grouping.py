import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1MjLW8NbD5ZjoOdd90myGv0i1NGAtlvScxebfAXMM1j8'

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key(SHEET_KEY)
ws = sh.worksheet('stuck')
vals = ws.get_all_values()

headers = vals[0]

def parse_ton_hours(val):
    s = str(val).strip()
    if not s:
        return 0.0
    try:
        return float(s.split("_")[0])
    except ValueError:
        return 0.0

def get_lc_aging_group(ton_str, ton_hours):
    s = str(ton_str).strip().lower()
    if '192' in s and '120' not in s:
        return "192h+"
    elif '120_192' in s or (120 <= ton_hours < 192):
        return "120 - 192h"
    elif '72_96' in s or '96_120' in s or (72 <= ton_hours < 120):
        return "72 - 120h"
    elif '36_48' in s or '48_72' in s or (36 <= ton_hours < 72):
        return "36 - 72h"
    elif ton_hours < 36:
        return "Dưới 36h"
    return "Khác"

am_orders = {}

for r in vals[1:]:
    if len(r) >= 9:
        ma_don = r[1].strip()
        st = r[4].strip()
        ton_raw = r[5].strip()
        bc = r[6].strip()
        am = r[8].strip()

        if st == "Đã đóng kiện" or not ma_don:
            continue

        ton_h = parse_ton_hours(ton_raw)
        grp = get_lc_aging_group(ton_raw, ton_h)

        am_orders.setdefault(am, []).append({
            "code": ma_don,
            "bc": bc,
            "ton_raw": ton_raw,
            "ton_h": ton_h,
            "group": grp
        })

print(f"Total AMs: {len(am_orders)}")
for am, ords in list(am_orders.items())[:3]:
    print(f"\nAM: {am} ({len(ords)} total orders)")
    by_grp = {}
    for o in ords:
        by_grp.setdefault(o["group"], []).append(o)
    for g, g_ords in by_grp.items():
        print(f"  Group '{g}': {len(g_ords)} orders")
