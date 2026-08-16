import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("gtc")
    
    # Get unformatted values of the first row and Cam Linh row
    print("Unformatted value of A1:", ws.get('A1', value_render_option='UNFORMATTED_VALUE'))
    
    # Search for Cam Linh row index
    rows = ws.get_all_values()
    cam_linh_idx = -1
    for idx, r in enumerate(rows):
        if "Cam Linh" in r[1] and "2026-07-04" in r[3] and "Hàng Mới Ca 1" in r[2]:
            cam_linh_idx = idx + 1 # 1-indexed
            break
            
    if cam_linh_idx != -1:
        print(f"Cam Linh row index in sheet: {cam_linh_idx}")
        # Get unformatted row values
        row_vals = ws.row_values(cam_linh_idx, value_render_option='UNFORMATTED_VALUE')
        print("Unformatted row values:", row_vals)
        print("Row values types       :", [type(x) for x in row_vals])
    else:
        print("❌ Not found Cam Linh row.")

if __name__ == "__main__":
    inspect()
