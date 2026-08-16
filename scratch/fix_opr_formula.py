import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    ws_opr = sh.worksheet("OPR")
    
    # 1. Clear cells L3 to L50500
    print("Clearing individual formulas in range L3:L50500...")
    ws_opr.batch_clear(["L3:L50500"])
    print("Successfully cleared!")
    
    # 2. Update L2 formula
    formula_l2 = '=ArrayFormula(IF(B2:B="", "", IF((B2:B="1.Tạo trước 9h") + (B2:B="3.Tạo sau 19h"), "2.Tạo từ 19h-9h", "1.Tạo từ 9h-19h")))'
    print("Updating L2 formula...")
    ws_opr.update_acell("L2", formula_l2)
    print("Successfully updated formula in L2!")
    
    # Verify L2:L5
    val = ws_opr.get_values("L2:L5", value_render_option="FORMULA")
    print(f"Verified L2:L5 formulas/values: {val}")
    
    # Verify evaluated values
    raw_vals = ws_opr.get_values("A2:L6")
    print("\nFirst few rows with calculated Khung giờ tạo:")
    for idx, r in enumerate(raw_vals):
        # Pad row to at least 12 elements if it's truncated
        if len(r) < 12:
            r = r + [""] * (12 - len(r))
        print(f"Row {idx+2}: {r[0]} | khung_gio_tao={r[1]} | AM={r[10]} | KhungGioTao={r[11]}")

if __name__ == "__main__":
    main()
