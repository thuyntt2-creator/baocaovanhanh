import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # Read Data
    ws_data = sh.worksheet("Data")
    data_vals = ws_data.get_all_values()
    df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])
    data_po_names = set(df_data['Chi tiết'].str.strip().unique())
    
    # Read Cơ cấu
    ws_cc = sh.worksheet("Cơ cấu")
    cc_vals = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
    
    cc_buucuc = set(df_cc['Bưu cục'].str.strip().unique())
    cc_bc = set(df_cc['BC'].str.strip().unique())
    cc_codes = set(df_cc['Mã bưu cục'].str.strip().unique())
    
    print(f"Unique 'Chi tiết' in Data: {len(data_po_names)}")
    print(f"Unique 'Bưu cục' in Cơ cấu: {len(cc_buucuc)}")
    print(f"Unique 'BC' in Cơ cấu: {len(cc_bc)}")
    print(f"Unique 'Mã bưu cục' in Cơ cấu: {len(cc_codes)}")
    
    # Check intersections
    int_buucuc = data_po_names.intersection(cc_buucuc)
    int_bc = data_po_names.intersection(cc_bc)
    
    print(f"\nIntersection with 'Bưu cục': {len(int_buucuc)}")
    print(f"Intersection with 'BC': {len(int_bc)}")
    
    unmatched = data_po_names - cc_buucuc - cc_bc
    print(f"Unmatched 'Chi tiết' count: {len(unmatched)}")
    print("Sample unmatched:")
    print(list(unmatched)[:15])
    
    # Check if there is some mapping between Mã bưu cục and Chi tiết in Data
    # Let's inspect some rows in Data to see if we can find how they mapped
    print("\nSample rows from Data:")
    print(df_data[['Cấp Quản Lý', 'Chi tiết', 'Time', 'Volume']].head(10).to_string())

if __name__ == "__main__":
    main()
