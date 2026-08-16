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

# Cấu hình
JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SHEET_KEY = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
TARGET_SHEET_NAME = "GTC 7D dưới 50%"  # Tab kết quả sẽ đẩy dữ liệu lên

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

def analyze_and_push_gtc_7d():
    print("🔑 Đang xác thực với Google Sheets API...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    print(f"📖 Đang mở bảng tính (Key: {SHEET_KEY})...")
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Tìm tab 'Data'
    data_tab_name = None
    for ws in sh.worksheets():
        if ws.title.strip().lower() == "data":
            data_tab_name = ws.title
            break
            
    if not data_tab_name:
        print("❌ Không tìm thấy tab 'Data' trong bảng tính!")
        return

    print(f"\n📊 --- 1. ĐỌC VÀ XỬ LÝ DỮ LIỆU TỪ TAB '{data_tab_name}' ---")
    ws_data = sh.worksheet(data_tab_name)
    all_values = ws_data.get_all_values()
    
    if not all_values:
        print(f"❌ Tab '{data_tab_name}' trống!")
        return
        
    df = pd.DataFrame(all_values[1:], columns=[c.strip() for c in all_values[0]])
    
    # Trích xuất ngày từ cột 'Time'
    df['date_str'] = df['Time'].apply(lambda x: str(x).split(' ')[0])
    all_dates = sorted(df['date_str'].unique())
    last_7_dates = all_dates[-7:]
    print(f"📅 7 ngày gần nhất: {last_7_dates}")
    
    # Lọc dữ liệu trong 7 ngày gần nhất
    df_7d = df[df['date_str'].isin(last_7_dates)].copy()
    
    # Làm sạch số liệu
    df_7d['Vol_Clean'] = df_7d['Volume'].apply(clean_volume)
    if 'Sản Lượng Giao Thành Công' in df_7d.columns:
        df_7d['GTC_Vol_Clean'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_volume)
    elif '% GTC' in df_7d.columns:
        df_7d['GTC_Pct_Clean'] = df_7d['% GTC'].apply(clean_pct)
        df_7d['GTC_Vol_Clean'] = df_7d['Vol_Clean'] * df_7d['GTC_Pct_Clean']
        
    # Lấy thông tin Khu vực (Vùng - Tỉnh) và AM Quản lý cho từng bưu cục
    df_7d['Khu_Vuc'] = df_7d.apply(
        lambda r: f"{r.get('Vùng', '')} - {r.get('Tỉnh', '')}".strip(" -"), axis=1
    )
    df_7d['AM_Name'] = df_7d.apply(
        lambda r: r.get('AM', '') or r.get('AM name', ''), axis=1
    )
    
    # Gom nhóm theo Bưu cục ('Chi tiết')
    grouped = df_7d.groupby('Chi tiết').agg(
        Khu_Vuc=('Khu_Vuc', 'first'),
        AM_Quản_Lý=('AM_Name', 'first'),
        Total_Volume=('Vol_Clean', 'sum'),
        Total_GTC_Volume=('GTC_Vol_Clean', 'sum')
    ).reset_index()
    
    # Tính tỷ lệ GTC 7 ngày
    grouped['GTC_Rate_7d'] = grouped['Total_GTC_Volume'] / grouped['Total_Volume']
    
    # Lọc các bưu cục có GTC < 50% và có sản lượng
    low_gtc = grouped[(grouped['GTC_Rate_7d'] < 0.50) & (grouped['Total_Volume'] > 0)].copy()
    low_gtc = low_gtc.sort_values(by='GTC_Rate_7d')
    
    print(f"📍 Phát hiện {len(low_gtc)} bưu cục có tỷ lệ GTC 7 ngày < 50%.")
    
    # 2. Chuẩn bị bảng dữ liệu để đẩy lên Google Sheet
    headers = ['STT', 'Bưu cục', 'Khu vực', 'AM Quản lý', 'Tỷ lệ GTC (7 ngày)', 'Sản lượng GTC', 'Tổng sản lượng Volume']
    
    upload_rows = [headers]
    for idx, (_, row) in enumerate(low_gtc.iterrows(), 1):
        gtc_rate_str = f"{row['GTC_Rate_7d'] * 100:.2f}%"
        upload_rows.append([
            str(idx),
            str(row['Chi tiết']),        # Cột 'Chi tiết' chính là Tên Bưu cục
            str(row['Khu_Vuc']),         # Khu vực (Vùng - Tỉnh)
            str(row['AM_Quản_Lý']),      # AM Quản lý
            gtc_rate_str,                # Tỷ lệ GTC %
            int(round(row['Total_GTC_Volume'])), # Sản lượng GTC
            int(round(row['Total_Volume']))      # Tổng Volume
        ])
        
    # 3. Đẩy lên tab 'GTC 7D dưới 50%' trên Google Sheet
    print(f"\n🚀 --- 2. ĐANG ĐẨY KẾT QUẢ LÊN TAB '{TARGET_SHEET_NAME}' ---")
    try:
        try:
            ws_target = sh.worksheet(TARGET_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws_target = sh.add_worksheet(title=TARGET_SHEET_NAME, rows="100", cols="10")
            
        # Xóa dữ liệu cũ và ghi dữ liệu mới
        ws_target.clear()
        ws_target.update(upload_rows)
        print(f"✅ Đã cập nhật thành công {len(upload_rows) - 1} dòng kết quả lên tab '{TARGET_SHEET_NAME}'!")
        
    except Exception as e:
        print(f"❌ Lỗi khi ghi lên tab '{TARGET_SHEET_NAME}': {e}")

if __name__ == "__main__":
    analyze_and_push_gtc_7d()
