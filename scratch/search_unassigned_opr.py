import os
import sys
import io
import gspread
import pandas as pd
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
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    print("Checking OPR sheet specifically...")
    ws = sh.worksheet("OPR")
    
    # Check for #N/A or error cells in OPR sheet
    cells_na = ws.findall("#N/A")
    if cells_na:
        print(f"Tab 'OPR': Found {len(cells_na)} '#N/A' cells. First 15:")
        for cell in cells_na[:15]:
            print(f"  - Row {cell.row}, Col {cell.col} (value: {cell.value})")
            # Let's read the formula for this row to see what's happening
            val_formula = ws.get_values(f"A{cell.row}:L{cell.row}", value_render_option="FORMULA")
            print(f"    Formula/Data in row {cell.row}: {val_formula}")
            
    # Also find if there are any AMs that are literally "Chưa gán AM" or empty
    raw_vals = ws.get_all_values()
    headers = raw_vals[0]
    print(f"OPR headers: {headers}")
    
    # Find AM column index (usually column 10, i.e. K)
    # Let's scan all rows to see if any AM column contains "Chưa gán AM", empty, or "#N/A"
    col_am_idx = 10 # index 10 is K
    for idx, h in enumerate(headers):
        if h.lower() == "am":
            col_am_idx = idx
            break
            
    print(f"AM column index is {col_am_idx} (Column {chr(65+col_am_idx)})")
    
    unassigned_rows = []
    for row_idx, r in enumerate(raw_vals[1:], start=2):
        if len(r) > col_am_idx:
            am_val = r[col_am_idx]
            if am_val == "Chưa gán AM" or am_val == "#N/A" or not am_val.strip() or am_val == "Không xác định":
                unassigned_rows.append((row_idx, am_val, r[5] if len(r) > 5 else "")) # IDKhoLay/KhoLay
                
    if unassigned_rows:
        print(f"\nFound {len(unassigned_rows)} rows in 'OPR' tab with unassigned AM. First 15:")
        for r_idx, val, kho in unassigned_rows[:15]:
            print(f"  - Row {r_idx}: AM='{val}', KhoLay='{kho}'")
    else:
        print("\n✅ No rows with unassigned AM found in 'OPR' tab!")

if __name__ == "__main__":
    main()
