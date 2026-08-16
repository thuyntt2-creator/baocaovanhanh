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
SHEET_KEY = '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_dates():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_push = sh.worksheet("PUSH REGION")
    push_data = ws_push.get_all_values()
    
    if len(push_data) < 2:
        print("Sheet trống")
        return
        
    headers = [h.strip() for h in push_data[0]]
    rows = push_data[1:]
    
    try:
        time_idx = headers.index("Cập nhật lúc")
    except ValueError:
        print("Không tìm thấy cột Cập nhật lúc")
        return
        
    timestamps = [r[time_idx].strip() for r in rows if len(r) > time_idx and r[time_idx].strip()]
    unique_timestamps = sorted(list(set(timestamps)))
    
    print(f"Tổng số dòng: {len(rows)}")
    print(f"Các mốc thời gian cập nhật khác nhau trong cột 'Cập nhật lúc':")
    for t in unique_timestamps[:10]:
        print(f"  - {t}")
    if len(unique_timestamps) > 10:
        print(f"  ... và {len(unique_timestamps) - 10} mốc thời gian khác.")

if __name__ == "__main__":
    inspect_dates()
