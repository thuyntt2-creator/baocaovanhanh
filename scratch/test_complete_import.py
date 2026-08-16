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

def parse_pct(val):
    val = str(val).replace('%', '').strip()
    val = val.replace(',', '.')
    try:
        return float(val) / 100.0
    except ValueError:
        return 0.0

def parse_float(val):
    val = str(val).strip()
    val = val.replace(',', '.')
    try:
        return float(val)
    except ValueError:
        return 0.0

def parse_int(val):
    val = str(val).strip()
    # Remove dots and commas (e.g., 1,054 or 1.054 -> 1054)
    val = val.replace(',', '').replace('.', '')
    try:
        return int(val)
    except ValueError:
        return 0

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
    
    # Apply maps and parse values
    df_filtered = df_gtc.iloc[:, :9].copy()
    
    # Map Cấp Quản Lý and Chi tiết
    df_filtered.iloc[:, 0] = df_filtered[col_detail].apply(lambda x: bc_to_am.get(normalize_key(x), "Không xác định"))
    df_filtered.iloc[:, 1] = df_filtered[col_detail].apply(lambda x: bc_to_mapped_bc.get(normalize_key(x), x))
    
    # Parse numbers
    df_filtered.iloc[:, 4] = df_filtered.iloc[:, 4].apply(parse_int) # Volume
    df_filtered.iloc[:, 5] = df_filtered.iloc[:, 5].apply(parse_pct) # % Gán
    df_filtered.iloc[:, 6] = df_filtered.iloc[:, 6].apply(parse_pct) # % GTC
    df_filtered.iloc[:, 7] = df_filtered.iloc[:, 7].apply(parse_pct) # % Chuyển trả
    df_filtered.iloc[:, 8] = df_filtered.iloc[:, 8].apply(parse_float) # Leadtime
    
    # Print sample
    print("Sample rows after mapping and parsing:")
    print(df_filtered.head(10))

if __name__ == "__main__":
    test_map()
