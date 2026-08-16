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
    
    ws_vung = sh.worksheet("CoCauVung")
    vung_vals = ws_vung.get_all_values()
    df_vung = pd.DataFrame(vung_vals[1:], columns=vung_vals[0])
    df_vung.columns = [c.strip() for c in df_vung.columns]
    
    search_terms = ['Lê Hồng Phong', 'Huỳnh Thúc Kháng', 'Phú Tài', 'Phú Thủy', 'Mũi Né']
    
    print("\nSearching in CoCauVung:")
    for term in search_terms:
        matches = df_vung[df_vung['Bưu cục'].str.contains(term, case=False, na=False)]
        print(f"\nTerm '{term}' matches in CoCauVung: {len(matches)}")
        if len(matches) > 0:
            print(matches.to_string())
            
    # Also let's inspect unique warehouse_ids in CoCauVung
    print("\nUnique warehouse_ids in CoCauVung:")
    print(df_vung['warehouse_id'].unique())

if __name__ == "__main__":
    main()
