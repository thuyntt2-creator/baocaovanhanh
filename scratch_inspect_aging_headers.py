import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

JSON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_aging_headers():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
    ws = sh.worksheet("Đơn giao aging trên 5 ngày")
    data = ws.get_all_values()
    if data:
        print("Headers of 'Đơn giao aging trên 5 ngày':")
        print(data[0])
        print("\nVí dụ 1 dòng dữ liệu:")
        if len(data) > 1:
            print(data[1])
    else:
        print("Sheet trống")

if __name__ == "__main__":
    inspect_aging_headers()
