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

def apply_auto_formula():
    print(f"📖 Connecting to sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    ws = sh.worksheet("DRAFT")
    
    # Công thức tối ưu mới:
    # 1. Lấy dữ liệu thô từ IMPORTRANGE & QUERY từ dòng A2 (không lấy header)
    # 2. SORTN lấy unique theo cột 7 (ID phường/xã)
    # 3. FILTER loại bỏ các ID đã có trong tab 'Đang OFF' cột G
    # 4. Ghép dòng headers lên đầu
    new_formula = (
        '=LET(\n'
        '  raw_data, QUERY(IMPORTRANGE("1PIyzade3_ml9Zq8OwTD5WGrc-paZJId7DfB06qSWJpg", "Tổng hợp!A2:Ab"), "SELECT * WHERE Col3 = \'NTB\' AND Col26 > date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\'", 0),\n'
        '  headers, IMPORTRANGE("1PIyzade3_ml9Zq8OwTD5WGrc-paZJId7DfB06qSWJpg", "Tổng hợp!A1:Ab1"),\n'
        '  unique_data, SORTN(raw_data, 9^9, 2, 7, TRUE),\n'
        '  filtered_data, FILTER(unique_data, ISNA(MATCH(INDEX(unique_data,,7), \'Đang OFF\'!G2:G, 0))),\n'
        '  {headers; filtered_data}\n'
        ')'
    )
    
    print("✍️ Ghi công thức tự động lọc trùng vào ô A1...")
    ws.update_acell("A1", new_formula)
    print("✅ Đã cập nhật công thức thành công!")
    
if __name__ == "__main__":
    apply_auto_formula()
