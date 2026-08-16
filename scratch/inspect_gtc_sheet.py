import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Set encoding for Windows output
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_sheet():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sheet_key = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
    
    sh = gc_client.open_by_key(sheet_key)
    print(f"Spreadsheet: {sh.title}")
    
    target_ws = None
    target_gid = 324939847
    for ws in sh.worksheets():
        if ws.id == target_gid:
            target_ws = ws
            break
            
    if not target_ws:
        print(f"Không tìm thấy tab với gid {target_gid}. Các tab hiện có:")
        for ws in sh.worksheets():
            print(f"  - {ws.title} (ID: {ws.id})")
        return
        
    print(f"Found sheet tab: '{target_ws.title}' (ID: {target_ws.id})")
    
    # Get all values
    all_values = target_ws.get_all_values()
    if not all_values:
        print("Bảng tính trống!")
        return
        
    print(f"Tổng số dòng: {len(all_values)}")
    # Print the first 10 rows to inspect headers and data format
    for idx, row in enumerate(all_values[:15]):
        print(f"Row {idx}: {row[:15]}")

if __name__ == "__main__":
    inspect_sheet()
