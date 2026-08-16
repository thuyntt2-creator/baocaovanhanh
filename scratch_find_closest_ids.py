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
    
    # 1. Load Cơ cấu
    ws_cc = sh.worksheet("Cơ cấu")
    cc_vals = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
    df_cc['Mã bưu cục_clean'] = df_cc['Mã bưu cục'].str.strip()
    
    # 2. Check screenshot IDs
    screenshot_ids = [
        '21624000', '22428000', '21880000', '20165000', '20725000', 
        '20492000', '22858000', '21530000', '20782000', '1352', 
        '22955000', '20665000', '22409000', '20998000', '21608000', 
        '21993000', '22958000'
    ]
    
    print("\nComparing screenshot IDs with Mã bưu cục:")
    for sid in screenshot_ids:
        # Check if sid starts with any Mã bưu cục, or any Mã bưu cục starts with sid
        # Also check with trailing zeros stripped
        sid_stripped = sid.rstrip('0')
        
        matches = []
        for idx, row in df_cc.iterrows():
            code = row['Mã bưu cục_clean']
            code_stripped = code.rstrip('0')
            
            if code == sid or code_stripped == sid_stripped or code.startswith(sid_stripped) or sid.startswith(code_stripped):
                matches.append(row.to_dict())
                
        if matches:
            print(f"WarehouseID {sid} matched:")
            for m in matches:
                print(f"  -> Mã bưu cục: {m['Mã bưu cục']} | Bưu cục: {m['Bưu cục']} | Tỉnh: {m['Tỉnh']} | AM: {m['Am']}")
        else:
            print(f"WarehouseID {sid} -> NO MATCH")

if __name__ == "__main__":
    main()
