import os
import json
import gspread
from google.oauth2.service_account import Credentials

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
    
    result = {}
    for ws in sh.worksheets():
        try:
            all_rows = ws.get_all_values()
            if not all_rows:
                result[ws.title] = {"empty": True}
                continue
            headers = all_rows[0]
            rows_preview = all_rows[1:10]  # Let's get first 9 data rows to be safe
            result[ws.title] = {
                "headers": headers,
                "rows": rows_preview,
                "total_rows": len(all_rows)
            }
        except Exception as e:
            result[ws.title] = {"error": str(e)}
            
    out_path = os.path.join(BASE_DIR, "scratch", "sheet_data_inspect.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"Saved inspection data to {out_path}")

if __name__ == "__main__":
    main()
