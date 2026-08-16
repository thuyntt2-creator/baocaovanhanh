import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    
    # Try to find the worksheet "Trang tính26" or "Trang tính 26"
    target_ws = None
    for ws in sh.worksheets():
        if "trang tính" in ws.title.lower():
            print(f"Found worksheet: '{ws.title}'")
            if "26" in ws.title:
                target_ws = ws
                
    if not target_ws:
        # Fallback to the first "trang tính" sheet
        for ws in sh.worksheets():
            if "trang tính" in ws.title.lower():
                target_ws = ws
                break
                
    if not target_ws:
        print("❌ Không tìm thấy tab 'Trang tính 26'.")
        return
        
    print(f"Reading from worksheet '{target_ws.title}'...")
    rows = target_ws.get_all_values()
    df = pd.DataFrame(rows[1:], columns=rows[0])
    
    print("\nColumns in sheet:")
    print(list(df.columns))
    
    print("\nRows for 'Cam Linh' in sheet:")
    for idx, r in df.iterrows():
        if any("Cam Linh" in str(cell) for cell in r):
            print(list(r))

if __name__ == "__main__":
    inspect()
