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
    
    # Load Cơ cấu
    ws_cc = sh.worksheet("Cơ cấu")
    cc_vals = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_vals[1:], columns=cc_vals[0])
    
    # Specific IDs from screenshot
    ids_to_check = [
        '21624000', '22428000', '21880000', '20165000', '20725000', 
        '20492000', '22858000', '21530000', '20782000', '1352', 
        '22955000', '20665000', '22409000', '20998000', '21608000', 
        '21993000', '22958000'
    ]
    
    print("\nChecking if IDs from screenshot are in 'Cơ cấu' (by 'Mã bưu cục' or other columns):")
    for column in df_cc.columns:
        # Find matches in this column
        matches = df_cc[df_cc[column].str.strip().isin(ids_to_check)]
        print(f"Column '{column}' matches: {len(matches)}")
        if len(matches) > 0:
            print(matches[[column] + [c for c in df_cc.columns if c != column]].to_string(index=False))

if __name__ == "__main__":
    main()
