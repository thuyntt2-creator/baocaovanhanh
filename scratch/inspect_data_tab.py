import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_top():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("Data")
    
    row1_formulas = ws.get("A1:V1", value_render_option="FORMULA")
    row2_formulas = ws.get("A2:V2", value_render_option="FORMULA")
    row1_values = ws.get("A1:V1", value_render_option="FORMATTED_VALUE")
    row2_values = ws.get("A2:V2", value_render_option="FORMATTED_VALUE")
    
    print("--- ROW 1 (Header) ---")
    r1_f = row1_formulas[0] if row1_formulas else []
    r1_v = row1_values[0] if row1_values else []
    for i in range(max(len(r1_f), len(r1_v))):
        val = r1_v[i] if i < len(r1_v) else ""
        form = r1_f[i] if i < len(r1_f) else ""
        print(f"Col {i+1:02d}: Value={repr(val)} | Formula={repr(form)}")
        
    print("\n--- ROW 2 ---")
    r2_f = row2_formulas[0] if row2_formulas else []
    r2_v = row2_values[0] if row2_values else []
    for i in range(max(len(r2_f), len(r2_v))):
        val = r2_v[i] if i < len(r2_v) else ""
        form = r2_f[i] if i < len(r2_f) else ""
        print(f"Col {i+1:02d}: Value={repr(val)} | Formula={repr(form)}")

if __name__ == "__main__":
    inspect_top()
