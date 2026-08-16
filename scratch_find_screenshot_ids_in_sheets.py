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
    
    screenshot_ids = [
        '21624000', '22428000', '21880000', '20165000', '20725000',
        '20492000', '22858000', '21530000', '20782000', '1352',
        '22955000', '20665000', '22409000', '20998000', '21608000',
        '21993000', '22958000', '21572000', '20499000', '21665000'
    ]
    
    for ws in sh.worksheets():
        title = ws.title
        if title == 'data thô':
            continue
            
        print(f"Searching in worksheet: '{title}'...")
        try:
            vals = ws.get_all_values()
            if not vals:
                continue
            df = pd.DataFrame(vals)
            for tid in screenshot_ids:
                mask = df.astype(str).apply(lambda col: col.str.strip() == tid)
                matches = df[mask.any(axis=1)]
                if len(matches) > 0:
                    print(f"  🎯 Found screenshot ID '{tid}' in sheet '{title}'! {len(matches)} matching rows:")
                    print(matches.to_string(header=False, index=False))
        except Exception as e:
            print(f"  Error reading '{title}': {e}")

if __name__ == "__main__":
    main()
