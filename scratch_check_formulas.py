import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding cho Windows
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

def check_formulas():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    ws = sh.worksheet("DRAFT")
    
    # Đọc công thức thay vì giá trị hiển thị bằng cách dùng value_render_option="FORMULA"
    print("📖 Đọc công thức trong 10 dòng đầu của tab DRAFT...")
    row_formulas = ws.get("A1:K10", value_render_option="FORMULA")
    for r_idx, row in enumerate(row_formulas):
        print(f"Dòng {r_idx + 1}: {row}")

if __name__ == "__main__":
    check_formulas()
