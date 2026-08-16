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

def analyze():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sheet_key = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
    
    sh = gc_client.open_by_key(sheet_key)
    ws = sh.worksheet("Bưu cục")
    
    all_values = ws.get_all_values()
    if not all_values:
        print("Trống")
        return
        
    df = pd.DataFrame(all_values)
    print("Header rows:")
    for idx in range(5):
        print(f"Row {idx}: {list(df.iloc[idx])}")
        
    # Let's find out if there are multiple dates
    # Assuming Row 1 is header (index 1)
    df.columns = df.iloc[1]
    # Keep rows after header
    df_data = df.iloc[2:].copy()
    
    # Trim column names
    df_data.columns = [str(c).strip() for c in df_data.columns]
    
    print("\nColumns found:")
    print(df_data.columns.tolist())
    
    if 'Time' in df_data.columns:
        print("\nUnique 'Time' values:")
        print(df_data['Time'].unique())
        
    if 'Chi tiết' in df_data.columns:
        print(f"\nTotal rows with 'Chi tiết': {df_data['Chi tiết'].nunique()}")

if __name__ == "__main__":
    analyze()
