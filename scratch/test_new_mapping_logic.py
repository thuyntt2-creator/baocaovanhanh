import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import unicodedata

# Fix encoding
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def normalize_key(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip().lower())

def test_map():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    
    # Read Cơ cấu
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    
    bc_to_am = {}
    bc_to_mapped_bc = {}
    for row in cocau_rows[1:]:
        if len(row) >= 5:
            bc_name = normalize_key(row[1]) # Column B - Bưu cục (Looker name)
            mapped_bc = row[2].strip()       # Column C - BC (Mapped name)
            am_name = row[4].strip()         # Column E - Am (AM name)
            if bc_name:
                if am_name:
                    bc_to_am[bc_name] = am_name
                if mapped_bc:
                    bc_to_mapped_bc[bc_name] = mapped_bc
                    
    # Read gtc
    ws_gtc = sh.worksheet("gtc")
    gtc_rows = ws_gtc.get_all_values()
    df_gtc = pd.DataFrame(gtc_rows[1:], columns=gtc_rows[0])
    
    # Filter out total rows
    col_detail = df_gtc.columns[1]
    df_gtc = df_gtc[
        ~df_gtc[col_detail].astype(str).str.lower().str.contains('total|tổng', na=False)
    ]
    df_gtc = df_gtc[df_gtc[col_detail].astype(str).str.strip() != ""]
    
    # Apply maps
    df_gtc['Mapped_AM'] = df_gtc[col_detail].apply(lambda x: bc_to_am.get(normalize_key(x), "NOT FOUND"))
    df_gtc['Mapped_BC'] = df_gtc[col_detail].apply(lambda x: bc_to_mapped_bc.get(normalize_key(x), "NOT FOUND"))
    
    # Print sample
    print(df_gtc[['Cấp Quản Lý', col_detail, 'Mapped_AM', 'Mapped_BC']].head(20))
    
    # Check if there are differences
    diff_bc = df_gtc[df_gtc[col_detail] != df_gtc['Mapped_BC']]
    print(f"\nTotal rows: {len(df_gtc)}")
    print(f"Rows where Looker BC != Mapped BC: {len(diff_bc)}")
    if len(diff_bc) > 0:
        print("Sample of different BC names:")
        print(diff_bc[[col_detail, 'Mapped_BC']].drop_duplicates().head(20))

if __name__ == "__main__":
    test_map()
