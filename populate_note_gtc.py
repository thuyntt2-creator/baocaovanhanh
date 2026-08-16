import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Force UTF-8 encoding cho console trên Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.exists(os.path.join(BASE_DIR, 'credentials.json')):
    JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
else:
    JSON_FILE = os.path.join(os.path.dirname(BASE_DIR), 'credentials.json')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# 1. Sheet chứa tab Data vận hành (Data)
DATA_SHEET_KEY = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"

# 2. Sheet chứa tab NOTE và Cocau
NOTE_SHEET_KEY = "1-pnZL8EXqGYGxh3wLQR1Z3Ft_oEPs-p42G6ZC8hFKRw"

def clean_volume(val):
    if not val or pd.isna(val):
        return 0
    val_str = str(val).strip()
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        parts = val_str.split(',')
        if len(parts[-1]) == 3:
            val_str = val_str.replace(',', '')
        else:
            val_str = val_str.replace(',', '.')
    elif '.' in val_str:
        parts = val_str.split('.')
        if len(parts[-1]) == 3:
            val_str = val_str.replace('.', '')
    try:
        return int(round(float(val_str)))
    except ValueError:
        return 0

def clean_pct(val):
    if not val or pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('%', '')
    val_str = val_str.replace(',', '.')
    try:
        num = float(val_str)
        if '%' in str(val) or num > 1.0:
            return num / 100.0
        return num
    except ValueError:
        return 0.0

def process_map_and_populate_note():
    print("🔑 Đang xác thực với Google Sheets API...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    # -------------------------------------------------------------
    # BƯỚC 1: Mở Sheet NOTE để đọc tab 'Cocau' và tab 'NOTE'
    # -------------------------------------------------------------
    print(f"📖 Đang mở Bảng tính NOTE (Key: {NOTE_SHEET_KEY})...")
    sh_note = gc_client.open_by_key(NOTE_SHEET_KEY)
    
    # 1.1 Lấy mapping warehouse_id -> Bưu cục từ tab 'Cocau'
    cocau_ws = None
    for ws in sh_note.worksheets():
        if ws.title.strip().lower() in ['cocau', 'cơ cấu']:
            cocau_ws = ws
            break
            
    if not cocau_ws:
        print("❌ Không tìm thấy tab 'Cocau' trong Sheet NOTE!")
        return
        
    print(f"📋 Đang đọc mapping warehouse_id -> Tên bưu cục từ tab '{cocau_ws.title}'...")
    cocau_vals = cocau_ws.get_all_values()
    df_cocau = pd.DataFrame(cocau_vals[1:], columns=[c.strip() for c in cocau_vals[0]])
    
    id_to_name_map = {}
    for _, row in df_cocau.iterrows():
        b_id = str(row.get('warehouse_id', '') or row.iloc[0]).strip()
        b_name = str(row.get('Bưu cục', '') or row.iloc[1]).strip()
        if b_id:
            id_to_name_map[b_id] = b_name
            
    print(f"✅ Đã tải được {len(id_to_name_map)} bưu cục từ tab Cocau.")

    # -------------------------------------------------------------
    # BƯỚC 2: Mở Sheet Data vận hành để tính toán GTC 7 ngày
    # -------------------------------------------------------------
    print(f"\n📊 Đang mở Bảng tính Data vận hành (Key: {DATA_SHEET_KEY})...")
    sh_data = gc_client.open_by_key(DATA_SHEET_KEY)
    
    data_ws = None
    for ws in sh_data.worksheets():
        if ws.title.strip().lower() == 'data':
            data_ws = ws
            break
            
    if not data_ws:
        print("❌ Không tìm thấy tab 'Data'!")
        return
        
    print(f"📊 Đang đọc và tính toán tỷ lệ GTC 7 ngày từ tab '{data_ws.title}'...")
    data_vals = data_ws.get_all_values()
    df_data = pd.DataFrame(data_vals[1:], columns=[c.strip() for c in data_vals[0]])
    
    df_data['date_str'] = df_data['Time'].apply(lambda x: str(x).split(' ')[0])
    all_dates = sorted(df_data['date_str'].unique())
    last_7_dates = all_dates[-7:]
    print(f"📅 7 ngày gần nhất trong Data: {last_7_dates}")
    
    df_7d = df_data[df_data['date_str'].isin(last_7_dates)].copy()
    df_7d['Vol_Clean'] = df_7d['Volume'].apply(clean_volume)
    
    if 'Sản Lượng Giao Thành Công' in df_7d.columns:
        df_7d['GTC_Vol_Clean'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_volume)
    elif '% GTC' in df_7d.columns:
        df_7d['GTC_Pct_Clean'] = df_7d['% GTC'].apply(clean_pct)
        df_7d['GTC_Vol_Clean'] = df_7d['Vol_Clean'] * df_7d['GTC_Pct_Clean']
        
    grouped = df_7d.groupby('Chi tiết').agg(
        Total_Volume=('Vol_Clean', 'sum'),
        Total_GTC_Volume=('GTC_Vol_Clean', 'sum')
    ).reset_index()
    
    grouped['GTC_Rate_7d'] = grouped['Total_GTC_Volume'] / grouped['Total_Volume']
    
    gtc_by_name_map = {}
    for _, row in grouped.iterrows():
        b_name = str(row['Chi tiết']).strip()
        gtc_by_name_map[b_name] = row['GTC_Rate_7d']

    # -------------------------------------------------------------
    # BƯỚC 3: Đọc tab NOTE và Điền kết quả vào Cột B & C
    # -------------------------------------------------------------
    note_ws = None
    for ws in sh_note.worksheets():
        if 'note' in ws.title.strip().lower():
            note_ws = ws
            break
            
    if not note_ws:
        print("❌ Không tìm thấy tab NOTE!")
        return
        
    print(f"✏️ Đang xử lý điền thông tin vào tab '{note_ws.title}'...")
    note_vals = note_ws.get_all_values()
    
    if not note_vals:
        print("❌ Tab NOTE trống!")
        return
        
    updated_rows_B_C = []
    for row in note_vals[1:]:
        b_id = str(row[0]).strip() if len(row) > 0 else ""
        
        # 1. Map ID -> Tên Bưu cục từ dictionary của tab Cocau
        name_found = id_to_name_map.get(b_id, "")
        
        # 2. Map Tên Bưu cục -> % GTC 7 ngày
        if name_found and name_found in gtc_by_name_map:
            rate = gtc_by_name_map[name_found]
            gtc_val_str = f"{rate * 100:.2f}%" if not pd.isna(rate) else "0.00%"
        else:
            gtc_val_str = "N/A"
            
        updated_rows_B_C.append([name_found if name_found else b_id, gtc_val_str])
        
    # Ghi dữ liệu vào Cột B và C bắt đầu từ B2
    range_to_update = f"B2:C{1 + len(updated_rows_B_C)}"
    note_ws.update(values=updated_rows_B_C, range_name=range_to_update)
    
    print(f"🎉 ĐÃ CẬP NHẬT THÀNH CÔNG {len(updated_rows_B_C)} BƯU CỤC VÀO SHEET NOTE!")
    for idx, row in enumerate(updated_rows_B_C, 1):
        print(f"  {idx}. ID: {note_vals[idx][0]} -> Tên: {row[0]} | GTC 7D: {row[1]}")

if __name__ == "__main__":
    process_map_and_populate_note()
