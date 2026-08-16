import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_baton = sh.worksheet("Bất ổn")
    data = ws_baton.get_all_values()
    print(f"Total rows in 'Bất ổn': {len(data)}")
    if data:
        headers = data[0]
        print(f"Headers count: {len(headers)}")
        # Print column letters and headers
        for idx, h in enumerate(headers):
            col_letter = gspread.utils.rowcol_to_a1(1, idx+1)[:-1]
            print(f"Col {col_letter} (Index {idx}): {h}")
            
        # Check column U (index 20) values
        u_values = set()
        for row in data[1:]:
            if len(row) > 20:
                val = row[20].strip()
                if val:
                    u_values.add(val)
        print("\nUnique non-empty values in Column U (index 20):")
        print(list(u_values))
        
        # Count how many rows have 'bất ổn' in Column U
        baton_count = 0
        for idx, row in enumerate(data[1:]):
            if len(row) > 20 and "bất ổn" in row[20].lower():
                baton_count += 1
                if baton_count <= 5:
                    print(f"Row {idx+2}: {row[2]} - {row[4]} -> Col U: '{row[20]}'")
        print(f"Total hubs with 'bất ổn' in Col U: {baton_count}")

if __name__ == "__main__":
    main()
