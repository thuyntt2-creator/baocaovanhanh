import os
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1705_0rKkgXBpsCbgK10EDr_mzSGhJOAcCa1WZsrWrU4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_KEY)
    
    ws_odr = sh.worksheet("ODR")
    odr_rows = ws_odr.get_all_values()
    df_odr = pd.DataFrame(odr_rows[1:], columns=odr_rows[0])
    
    print(f"Total rows in ODR: {len(df_odr)}")
    print("\nUnique 'Quản lý' in full sheet:")
    print(df_odr["Quản lý"].value_counts())
    
    print("\nUnique 'Chi tiết' in full sheet (First 30):")
    print(df_odr["Chi tiết"].value_counts().head(30))
    
    print("\nUnique 'Time' dates in full sheet:")
    print(df_odr["Time"].value_counts().sort_index())
    
    # Also read Cocau for mapping analysis
    ws_cocau = sh.worksheet("Cocau")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    # Check matching
    unmapped = []
    for po in df_odr["Chi tiết"].unique():
        match = df_cocau[df_cocau["Bưu cục"].str.lower() == po.lower()]
        if match.empty:
            unmapped.append(po)
            
    print(f"\nUnmapped POs count in full sheet: {len(unmapped)}")
    if unmapped:
        print(f"Unmapped examples: {unmapped[:10]}")

if __name__ == "__main__":
    main()
