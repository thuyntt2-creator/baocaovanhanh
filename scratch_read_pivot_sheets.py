import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

# Reconfigure stdout/stderr to use UTF-8
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    # Check 1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU
    print("--- Reading sheet 1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU ---")
    try:
        sh1 = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
        print("Title:", sh1.title)
        ws_pivot = sh1.worksheet("PIVOT")
        rows = ws_pivot.get_all_values()
        print("PIVOT rows (first 10):")
        for r in rows[:10]:
            print(r)
        print("PIVOT rows (last total row or top 5):")
        for r in rows[15:25]:
            print(r)
    except Exception as e:
        print("Error reading 1WCz:", e)

    # Check 1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M
    print("\n--- Reading sheet 1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M ---")
    try:
        sh2 = gc_client.open_by_key('1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M')
        print("Title:", sh2.title)
        print("Worksheets:")
        for ws in sh2.worksheets():
            print(f" - {ws.title}")
        if "PIVOT" in [ws.title for ws in sh2.worksheets()]:
            ws_pivot = sh2.worksheet("PIVOT")
            rows = ws_pivot.get_all_values()
            print("PIVOT rows (first 10):")
            for r in rows[:10]:
                print(r)
    except Exception as e:
        print("Error reading 1l2j:", e)

if __name__ == '__main__':
    main()
