import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

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
    sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
    
    # 1. Read aging orders
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()
    aging_header = aging_data[0]
    ag_order_idx = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
    
    target_orders = [row[ag_order_idx].strip() for row in aging_data[1:] if len(row) > ag_order_idx and row[ag_order_idx].strip()]
    print(f"Total target orders from 'Đơn giao aging trên 5 ngày': {len(target_orders)}")
    print("First 10 target orders:", target_orders[:10])

    # 2. Read data LM
    ws_lm = sh.worksheet("data LM")
    lm_data = ws_lm.get_all_values()
    lm_header = lm_data[1]
    order_col_idx = lm_header.index("Mã đơn hàng")
    status_col_idx = lm_header.index("Trạng thái")
    
    print("Header row in data LM:", lm_header)
    
    lm_status = {}
    for idx, row in enumerate(lm_data[2:]):
        if len(row) > max(order_col_idx, status_col_idx):
            m_don = row[order_col_idx].strip()
            t_thai = row[status_col_idx].strip()
            if m_don:
                lm_status[m_don] = t_thai

    print(f"Total unique orders found in 'data LM': {len(lm_status)}")
    
    # 3. Match
    found_count = 0
    statuses_found = {}
    for o in target_orders:
        if o in lm_status:
            found_count += 1
            st = lm_status[o]
            statuses_found[st] = statuses_found.get(st, 0) + 1
            
    print(f"Number of target orders found in 'data LM': {found_count}")
    print("Statuses of matched target orders:", statuses_found)

if __name__ == '__main__':
    main()
