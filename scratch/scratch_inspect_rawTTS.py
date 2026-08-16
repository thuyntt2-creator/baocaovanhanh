import os
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SHEET_KEY = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'

def run():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws = sh.worksheet("rawTTS")
    all_values = ws.get_all_values()
    df = pd.DataFrame(all_values)
    print("Shape:", df.shape)
    if len(df) > 0:
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        print("Columns:", list(df.columns))
        print("First 2 rows:")
        for idx, r in df.head(2).iterrows():
            print(r.to_dict())
            
if __name__ == "__main__":
    run()
