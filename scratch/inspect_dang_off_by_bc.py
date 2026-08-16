import os
import sys
import json
import gspread
import unicodedata
from google.oauth2.service_account import Credentials

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
CRED_FILE = r'c:\Users\lap4all\Documents\Auto report\credentials.json'

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(CRED_FILE, scopes=scopes)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_KEY)

ws = sh.worksheet("Đang OFF")
rows = ws.get_all_values()

# Group by Bưu cục (Col E / index 4)
bc_data = {}
for i in range(1, len(rows)):
    row = rows[i]
    if not any(row): continue
    
    province = row[0].strip()
    district = row[1].strip()
    ward = row[2].strip()
    ward_id = row[3].strip()
    bc = row[4].strip()
    result = row[6].strip()
    cap_down = row[7].strip()
    off_from = row[8].strip()
    off_to = row[9].strip()
    am = unicodedata.normalize("NFC", row[10].strip()) if len(row) > 10 else ""
    
    if bc not in bc_data:
        bc_data[bc] = {
            "am": am,
            "province": province,
            "district": district,
            "result": result,
            "cap_down": cap_down,
            "off_from": off_from,
            "off_to": off_to,
            "wards": []
        }
    bc_data[bc]["wards"].append({
        "ward": ward,
        "ward_id": ward_id,
        "district": district,
        "province": province
    })

print(f"Total Bưu Cục: {len(bc_data)}")
for bc, info in bc_data.items():
    print(f"\n📦 BƯU CỤC: {bc} | AM: {info['am']} | {len(info['wards'])} xã/phường | Tắt: {info['off_from']} -> Mở: {info['off_to']}")
    for w in info['wards']:
        print(f"   • {w['ward']} (ID: {w['ward_id']}) - {w['district']}, {w['province']}")
