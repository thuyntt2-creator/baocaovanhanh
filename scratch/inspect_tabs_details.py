import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_tabs():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sheet_key = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
    sh = gc_client.open_by_key(sheet_key)
    
    # 1. Bưu cục
    print("=== INSPECTING 'Bưu cục' TAB ===")
    ws_bc = sh.worksheet("Bưu cục")
    vals_bc = ws_bc.get_all_values()
    print(f"Bưu cục rows: {len(vals_bc)}")
    if len(vals_bc) > 0:
        for idx in range(min(5, len(vals_bc))):
            print(f"  Row {idx}: {vals_bc[idx][:18]}")
            
    # 2. gtc
    print("\n=== INSPECTING 'gtc' TAB ===")
    ws_gtc = sh.worksheet("gtc")
    vals_gtc = ws_gtc.get_all_values()
    print(f"gtc rows: {len(vals_gtc)}")
    if len(vals_gtc) > 0:
        for idx in range(min(5, len(vals_gtc))):
            print(f"  Row {idx}: {vals_gtc[idx][:10]}")
            
    # 3. raw
    print("\n=== INSPECTING 'raw' TAB ===")
    try:
        ws_raw = sh.worksheet("raw")
        vals_raw = ws_raw.get_all_values()
        print(f"raw rows: {len(vals_raw)}")
        if len(vals_raw) > 0:
            for idx in range(min(5, len(vals_raw))):
                print(f"  Row {idx}: {vals_raw[idx][:10]}")
    except Exception as e:
        print(f"Error reading 'raw': {e}")

if __name__ == "__main__":
    inspect_tabs()
