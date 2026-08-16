import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def test():
    # Find any excel file matching the pattern in Downloads
    DOWNLOADS_DIR = r"C:\Users\lap4all\Downloads"
    matched_file = None
    for f in os.listdir(DOWNLOADS_DIR):
        if "BÁO CÁO XỬ LÝ THÀNH CÔNG TRONG NGÀY" in f and f.endswith(".xlsx"):
            matched_file = os.path.join(DOWNLOADS_DIR, f)
            print(f"Found file: {f}")
            break
            
    if not matched_file:
        print("❌ Không tìm thấy file Excel nào trong Downloads.")
        return
        
    df = pd.read_excel(matched_file).fillna("")
    rename_map = {
        'Cấp quản lý': 'Cấp Quản Lý',
        'Loại hình': 'Loại Hàng',
        'Sản lượng': 'Volume',
        '% LTC': '% GTC',
        '% Đóng kiện': '% Chuyển trả',
        '% LC': 'Leadtime'
    }
    df.rename(columns=rename_map, inplace=True)
    
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    
    # Get or create worksheet 'test_gtc_upload'
    try:
        ws = sh.worksheet("test_gtc_upload")
        sh.del_worksheet(ws)
    except:
        pass
    ws = sh.add_worksheet(title="test_gtc_upload", rows=2000, cols=30)
    
    print("Uploading to 'test_gtc_upload'...")
    data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
    ws.update(data_to_upload)
    print("Upload completed successfully!")
    
    # Read it back and print Cam Linh row
    print("Reading it back from sheet...")
    rows = ws.get_all_values()
    print(f"Total rows in test_gtc_upload: {len(rows)}")
    
    for idx, r in enumerate(rows):
        if "Cam Linh" in r[1] and "2026-07-04" in r[3]:
            print("Row values from sheet:", r)
            unformatted = ws.row_values(idx + 1, value_render_option='UNFORMATTED_VALUE')
            print("Unformatted row values:", unformatted)
            print("Row values types       :", [type(x) for x in unformatted])

if __name__ == "__main__":
    test()
