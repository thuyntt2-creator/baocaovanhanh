"""
NTB FD Analysis – Google Sheets Auto Writer
============================================
Cách dùng:
  1. Điền CREDENTIALS_PATH và SPREADSHEET_ID bên dưới
  2. Chạy: python ntb_fd_gsheet.py
  3. Script tự đọc 3 raw sheets → tính FD → ghi vào 6 analysis sheets

Yêu cầu:
  pip install gspread google-auth pandas matplotlib requests pillow
"""

import sys, os
sys.stdout.reconfigure(encoding='utf-8')

import time
from PIL import Image
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import warnings
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import requests
import io
warnings.filterwarnings('ignore')

# ============================================================
#  CONFIG – CHỈ CẦN SỬA CÁC DÒNG NÀY
# ============================================================
CREDENTIALS_PATH = r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json'
SPREADSHEET_ID   = '15Z-aMM6OFfiWUXd2Zwz6BFNq_Y0KWwHiVDqxkioHufM'  # <-- ID file GSheet

# ── Telegram ──
TELEGRAM_BOT_TOKEN = '8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M'
TELEGRAM_CHAT_ID   = '-5058464865'

# ── GTalk ──
GTALK_OA_TOKEN     = '2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv'
GTALK_CHANNEL_ID   = '2073028116810764288'
# ============================================================

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_gspread_client(spreadsheet_id=SPREADSHEET_ID):
    candidates = [
        CREDENTIALS_PATH,
        os.path.join(BASE_DIR, 'credentials.json'),
        'credentials.json'
    ]
    for cred_path in candidates:
        if os.path.exists(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...")

    auth_user_candidates = [
        os.path.join(BASE_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for auth_user_file in auth_user_candidates:
        if os.path.exists(auth_user_file):
            try:
                from google.oauth2.credentials import Credentials as UserCredentials
                creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


RAW_SHEETS = {
    'TONG': 'Raw FD_Tổng',
    'COD':  'RAW FD_COD',
    'TTS':  'RAW FD_TTS',
}

TINH_LIST = ['Khánh Hòa', 'Lâm Đồng', 'Bình Thuận', 'Ninh Thuận', 'Đắk Nông']

# ── Màu sắc (RGB tuple) ──────────────────────────────────────
CLR = {
    'darkBlue' : {'red': 0.122, 'green': 0.306, 'blue': 0.475},
    'midBlue'  : {'red': 0.180, 'green': 0.459, 'blue': 0.714},
    'lightBlue': {'red': 0.839, 'green': 0.894, 'blue': 0.941},
    'altRow'   : {'red': 0.961, 'green': 0.976, 'blue': 1.000},
    'yellow'   : {'red': 1.000, 'green': 0.949, 'blue': 0.800},
    'orange'   : {'red': 0.773, 'green': 0.353, 'blue': 0.067},
    'redBg'    : {'red': 1.000, 'green': 0.878, 'blue': 0.878},
    'redFont'  : {'red': 0.753, 'green': 0.000, 'blue': 0.000},
    'greenBg'  : {'red': 0.886, 'green': 0.937, 'blue': 0.851},
    'greenFont': {'red': 0.216, 'green': 0.341, 'blue': 0.137},
    'white'    : {'red': 1.000, 'green': 1.000, 'blue': 1.000},
    'black'    : {'red': 0.000, 'green': 0.000, 'blue': 0.000},
    'gray'     : {'red': 0.941, 'green': 0.941, 'blue': 0.941},
}

# ============================================================
#  DATA LOADING
# ============================================================
def parse_pct(s):
    if pd.isna(s) or s == '': return None
    try:
        return float(str(s).replace('%','').replace(',','.'))
    except:
        return None

def get_tinh_from_bc(bc, bc_to_tinh=None):
    bc_clean = bc.strip()
    if bc_to_tinh and bc_clean in bc_to_tinh and bc_to_tinh[bc_clean]:
        return bc_to_tinh[bc_clean]
    
    bc_upper = bc_clean.upper()
    if '(KHO)' in bc_upper or 'KHÁNH HÒA' in bc_upper or 'KHANH HOA' in bc_upper:
        return 'Khánh Hòa'
    if '(LDO)' in bc_upper or 'LÂM ĐỒNG' in bc_upper or 'LAM DONG' in bc_upper:
        return 'Lâm Đồng'
    if '(BTH)' in bc_upper or 'BÌNH THUẬN' in bc_upper or 'BINH THUAN' in bc_upper:
        return 'Bình Thuận'
    if '(NTH)' in bc_upper or 'NINH THUẬN' in bc_upper or 'NINH THUAN' in bc_upper:
        return 'Ninh Thuận'
    if '(DNO)' in bc_upper or 'ĐẮK NÔNG' in bc_upper or 'DAK NONG' in bc_upper:
        return 'Đắk Nông'
    
    for t in TINH_LIST:
        if t.upper() in bc_upper:
            return t
    return 'Khác'

def load_raw(spreadsheet, sheet_name, bc_to_tinh=None):
    print(f'  Đọc sheet: {sheet_name}')
    sh = spreadsheet
    ws = sh.worksheet(sheet_name)
    data = ws.get_all_values()
    if len(data) < 2:
        return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = [str(c).strip() for c in df.columns]
    n_cols = len(df.columns)
    
    # Detect if the first column is Tinh or AM_raw
    first_col_vals = df.iloc[:, 0].dropna().astype(str).unique()[:5]
    is_first_col_tinh = False
    for val in first_col_vals:
        val_upper = val.upper()
        if 'NTB -' in val_upper or any(t.upper() in val_upper for t in TINH_LIST):
            is_first_col_tinh = True
            break
            
    if is_first_col_tinh:
        if n_cols >= 6:
            df.columns = ['Tinh', 'BC', 'Time', 'vol_pct', 'ret_pct', 'AM_raw'] + list(df.columns[6:])
            df = df[['Tinh','BC','Time','vol_pct','ret_pct','AM_raw']]
        else:
            df.columns = ['Tinh', 'BC', 'Time', 'vol_pct', 'ret_pct'] + list(df.columns[5:])
            df = df[['Tinh','BC','Time','vol_pct','ret_pct']]
            df['AM_raw'] = ''
    else:
        # First column is AM_raw
        if n_cols >= 6:
            df.columns = ['AM_raw', 'BC', 'Time', 'vol_pct', 'ret_pct', 'extra_col'] + list(df.columns[6:])
            df = df[['AM_raw', 'BC', 'Time', 'vol_pct', 'ret_pct']]
        else:
            df.columns = ['AM_raw', 'BC', 'Time', 'vol_pct', 'ret_pct'] + list(df.columns[5:])
            df = df[['AM_raw', 'BC', 'Time', 'vol_pct', 'ret_pct']]
        df['Tinh'] = df['BC'].apply(lambda x: get_tinh_from_bc(x, bc_to_tinh))
        
    df['date'] = pd.to_datetime(df['Time'].str[:10], errors='coerce').dt.normalize()
    df['vol']  = df['vol_pct'].apply(parse_pct)
    df['ret']  = df['ret_pct'].apply(parse_pct)
    df['BC'] = df['BC'].fillna('').str.strip()
    df['AM_raw'] = df['AM_raw'].fillna('').str.strip()
    df['Tinh'] = df['Tinh'].fillna('').str.strip()
    df = df.dropna(subset=['vol','ret','date'])
    return df

def load_cocau(spreadsheet):
    sh = spreadsheet
    try:
        ws = sh.worksheet('CoCauVung')
        data = ws.get_all_values()
        if len(data) < 2:
            return {}, {}
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # Deduplicate column names
        cols = []
        for i, col in enumerate(df.columns):
            c = str(col).strip()
            if c in cols:
                cols.append(f"{c}_{i}")
            else:
                cols.append(c)
        df.columns = cols
        
        bc_col = 'Bưu cục' if 'Bưu cục' in df.columns else df.columns[1]
        tinh_col = 'Tỉnh' if 'Tỉnh' in df.columns else df.columns[2]
        
        am_cols = [c for c in df.columns if c == 'AM']
        am_col = am_cols[0] if am_cols else df.columns[3]
            
        df_clean = df[[bc_col, tinh_col, am_col]].copy()
        df_clean.columns = ['BC', 'Tinh', 'AM']
        df_clean['BC'] = df_clean['BC'].str.strip()
        df_clean['Tinh'] = df_clean['Tinh'].str.strip()
        df_clean['AM'] = df_clean['AM'].str.strip()
        
        df_clean = df_clean[df_clean['BC'] != '']
        
        bc_to_am = dict(zip(df_clean['BC'], df_clean['AM']))
        bc_to_tinh = dict(zip(df_clean['BC'], df_clean['Tinh']))
        
        print(f'   Đọc CoCauVung: {len(bc_to_am)} BC mapped to AM, {len(bc_to_tinh)} BC mapped to Tinh')
        return bc_to_am, bc_to_tinh
    except Exception as e:
        print(f'   ⚠️ Lỗi đọc CoCauVung: {e}')
        return {}, {}

# ============================================================
#  DATA SHEET – Vol giao Ca1+Ca2 để tính Vol trả tuyệt đối
# ============================================================
def load_data_sheet(spreadsheet):
    """Đọc sheet 'data', lấy Vol Ca1+Ca2 theo BC+ngày, kèm AM (cột U)"""
    try:
        sh = spreadsheet
        ws = sh.worksheet('data')
        data = ws.get_all_values()
        if len(data) < 2:
            return pd.DataFrame(), {}
        df = pd.DataFrame(data[1:], columns=data[0])
        df = df.rename(columns={'Chi tiết': 'BC', 'Loại Hàng': 'LoaiHang'})
        df['BC'] = df['BC'].fillna('').str.strip()
        df['date']   = pd.to_datetime(df['Time'].str[:10], errors='coerce').dt.normalize()
        
        # Loại bỏ dấu phẩy ngăn cách hàng nghìn (ví dụ "1,020" -> "1020") trước khi chuyển sang số
        df['Volume'] = pd.to_numeric(df['Volume'].str.replace(',', '', regex=False), errors='coerce').fillna(0)

        # BC → AM map (từ cột AM trong data sheet)
        bc_am_data = {}
        if 'AM' in df.columns:
            bc_am_data = df.drop_duplicates('BC').set_index('BC')['AM'].to_dict()

        # Chỉ lấy Ca1 + Ca2
        ca12 = df[df['LoaiHang'].isin(['Hàng Mới Ca 1', 'Hàng Mới Ca 2'])]
        vol  = ca12.groupby(['BC','date'])['Volume'].sum().reset_index()
        vol.columns = ['BC', 'date', 'Vol_giao']
        print(f'   Đọc data sheet: {len(vol)} BC-ngày, {len(bc_am_data)} BC có AM')
        return vol, bc_am_data
    except Exception as e:
        print(f'   ⚠️ Không đọc được data sheet: {e}')
        return pd.DataFrame(), {}

def calc_vol_tra(df_raw, df_vol, date):
    """
    Vol trả BC = Vol_giao (data sheet) × %Return (Raw FD)
    Tỷ trọng  = Vol_trả_BC / Tổng_Vol_trả_NTB
    """
    date = pd.Timestamp(date).normalize()
    raw_day = df_raw[df_raw['date'] == date][['BC','ret']].copy()
    if df_vol.empty:
        return {}, {}, {}
    vol_day = df_vol[df_vol['date'] == date].copy()
    merged  = raw_day.merge(vol_day, on='BC', how='inner')
    merged['Vol_tra'] = merged['Vol_giao'] * merged['ret'] / 100
    total = merged['Vol_tra'].sum()
    vol_tra_dict  = dict(zip(merged['BC'], merged['Vol_tra'].round(1)))
    ty_trong_dict = dict(zip(merged['BC'], (merged['Vol_tra']/total*100).round(2))) if total > 0 else {}
    vol_giao_dict = dict(zip(merged['BC'], merged['Vol_giao']))
    return vol_tra_dict, ty_trong_dict, vol_giao_dict

def calc_vol_tra_am(df_raw, df_vol, date, bc_am_data, bc_to_am):
    """Vol trả + Tỷ trọng theo AM (AM map ưu tiên từ data sheet cột U)"""
    date = pd.Timestamp(date).normalize()
    raw_day = df_raw[df_raw['date'] == date][['BC','ret','AM_raw']].copy()
    def get_am(row):
        bc = row['BC']
        # Ưu tiên lấy từ CoCauVung trước (cơ cấu mới do user cập nhật)
        if bc in bc_to_am and bc_to_am[bc]:
            return bc_to_am[bc]
        if bc in bc_am_data and bc_am_data[bc]:
            return bc_am_data[bc]
        if row['AM_raw']:
            return row['AM_raw']
        return 'Chưa phân'
    raw_day['AM'] = raw_day.apply(get_am, axis=1)
    if df_vol.empty:
        return {}, {}
    vol_day = df_vol[df_vol['date'] == date].copy()
    merged  = raw_day.merge(vol_day, on='BC', how='inner')
    merged['Vol_tra'] = merged['Vol_giao'] * merged['ret'] / 100
    am_vol  = merged.groupby('AM')['Vol_tra'].sum()
    total   = am_vol.sum()
    ty_trong_dict = (am_vol / total * 100).round(2).to_dict() if total > 0 else {}
    vol_tra_dict  = am_vol.round(1).to_dict()
    return vol_tra_dict, ty_trong_dict


# ============================================================
#  FD CALCULATIONS
# ============================================================
def weighted_fd(sub):
    if sub.empty or sub['vol'].sum() == 0:
        return None
    return (sub['vol'] * sub['ret']).sum() / sub['vol'].sum()

def calc_fd_bc(df, date):
    date = pd.Timestamp(date).normalize()
    sub = df[df['date'] == date]
    if sub.empty: return {}
    return sub.groupby('BC').apply(weighted_fd).to_dict()

def calc_fd_am(df, date, bc_to_am, bc_am_data=None):
    date = pd.Timestamp(date).normalize()
    sub = df[df['date'] == date].copy()
    if sub.empty: return {}
    bc_am_data = bc_am_data or {}
    def get_am(row):
        bc = row['BC']
        # Ưu tiên lấy từ CoCauVung trước (cơ cấu mới do user cập nhật)
        if bc in bc_to_am and bc_to_am[bc]:
            return bc_to_am[bc]
        if bc in bc_am_data and bc_am_data[bc]:
            return bc_am_data[bc]
        if row['AM_raw']:
            return row['AM_raw']
        return 'Chưa phân'
    sub['AM'] = sub.apply(get_am, axis=1)
    return sub.groupby('AM').apply(weighted_fd).to_dict()

def calc_fd_tinh(df, date):
    date = pd.Timestamp(date).normalize()
    sub = df[df['date'] == date].copy()
    if sub.empty: return {}
    sub['Tinh_clean'] = sub['Tinh'].str.replace('NTB - ', '', regex=False)
    return sub.groupby('Tinh_clean').apply(weighted_fd).to_dict()

def calc_tytrong(df, date):
    date = pd.Timestamp(date).normalize()
    sub = df[df['date'] == date].copy()
    if sub.empty: return {}
    sub['w'] = sub['vol'] * sub['ret']
    total = sub['w'].sum()
    if total == 0: return {}
    return (sub.groupby('BC')['w'].sum() / total * 100).to_dict()

# ============================================================
#  BUILD ANALYSIS DATAFRAMES
# ============================================================
def build_bc_table(df, bc_to_am, latest, n1, n7, df_vol=None, bc_am_data=None):
    bc_am_data = bc_am_data or {}
    fd_n  = calc_fd_bc(df, latest)
    fd_n1 = calc_fd_bc(df, n1)
    fd_n7 = calc_fd_bc(df, n7)
    # Vol trả tuyệt đối nếu có data sheet, fallback %vol×%ret
    if df_vol is not None and not df_vol.empty:
        vol_tra_n, tt_n, vol_giao_n = calc_vol_tra(df, df_vol, latest)
        use_abs = True
    else:
        tt_n = calc_tytrong(df, latest)
        vol_tra_n = {}; vol_giao_n = {}; use_abs = False

    rows = []
    bc_am_from_raw = df[df['AM_raw'] != ''].drop_duplicates('BC').set_index('BC')['AM_raw'].to_dict()
    def get_am(bc):
        # Ưu tiên lấy từ CoCauVung trước (cơ cấu mới do user cập nhật)
        return bc_to_am.get(bc) or bc_am_data.get(bc) or bc_am_from_raw.get(bc) or ''

    for bc, fdn in sorted(fd_n.items(), key=lambda x: -x[1]):
        fdn1 = fd_n1.get(bc)
        fdn7 = fd_n7.get(bc)
        tt   = tt_n.get(bc)
        rows.append({
            'BC'         : bc,
            'AM'         : get_am(bc),
            'FD_N'       : fdn,
            'FD_N1'      : fdn1,
            'vs_N1'      : (fdn - fdn1) if fdn1 is not None else None,
            'FD_N7'      : fdn7,
            'vs_N7'      : (fdn - fdn7) if fdn7 is not None else None,
            'Vol_giao'   : vol_giao_n.get(bc),
            'Vol_tra'    : vol_tra_n.get(bc),
            'TyTrong'    : tt,
            'VuotTarget' : (fdn - 4.5) if fdn is not None else None,
        })
    df_out = pd.DataFrame(rows)
    # ── Sort theo %FD (cột D) giảm dần, thay vì Tỷ trọng ──
    df_out = df_out.sort_values('FD_N', ascending=False, na_position='last')
    return df_out

def calc_tytrong_am(df, date, bc_to_am, bc_am_data=None):
    bc_am_data = bc_am_data or {}
    date = pd.Timestamp(date).normalize()
    sub = df[df["date"] == date].copy()
    if sub.empty: return {}
    def get_am(row):
        bc = row["BC"]
        # Ưu tiên lấy từ CoCauVung trước (cơ cấu mới do user cập nhật)
        if bc in bc_to_am and bc_to_am[bc]:
            return bc_to_am[bc]
        if bc in bc_am_data and bc_am_data[bc]:
            return bc_am_data[bc]
        if row["AM_raw"]:
            return row["AM_raw"]
        return "Chưa phân"
    sub["AM"] = sub.apply(get_am, axis=1)
    sub["w"] = sub["vol"] * sub["ret"]
    total = sub["w"].sum()
    if total == 0: return {}
    return (sub.groupby("AM")["w"].sum() / total * 100).to_dict()

def build_am_table(df, bc_to_am, latest, n1, n7, df_vol=None, bc_am_data=None):
    bc_am_data = bc_am_data or {}
    fd_n  = calc_fd_am(df, latest, bc_to_am, bc_am_data)
    fd_n1 = calc_fd_am(df, n1,     bc_to_am, bc_am_data)
    fd_n7 = calc_fd_am(df, n7,     bc_to_am, bc_am_data)
    if df_vol is not None and not df_vol.empty:
        vol_tra_n, tt_n = calc_vol_tra_am(df, df_vol, latest, bc_am_data, bc_to_am)
    else:
        tt_n = calc_tytrong_am(df, latest, bc_to_am, bc_am_data)
        vol_tra_n = {}
    rows = []
    for am, fdn in sorted(fd_n.items(), key=lambda x: -x[1]):
        fdn1 = fd_n1.get(am)
        fdn7 = fd_n7.get(am)
        rows.append({
            "AM"        : am,
            "FD_N"      : fdn,
            "FD_N1"     : fdn1,
            "vs_N1"     : (fdn - fdn1) if fdn1 is not None else None,
            "FD_N7"     : fdn7,
            "vs_N7"     : (fdn - fdn7) if fdn7 is not None else None,
            "Vol_tra"   : vol_tra_n.get(am),
            "TyTrong"   : tt_n.get(am),
            "VuotTarget": (fdn - 4.5) if fdn is not None else None,
        })
    df_out = pd.DataFrame(rows)
    df_out = df_out.sort_values("FD_N", ascending=False, na_position="last")
    return df_out

def build_tinh_table(df, latest, n1, n7):
    fd_n  = calc_fd_tinh(df, latest)
    fd_n1 = calc_fd_tinh(df, n1)
    fd_n7 = calc_fd_tinh(df, n7)
    rows = []
    for tinh in TINH_LIST:
        fdn  = fd_n.get(tinh)
        fdn1 = fd_n1.get(tinh)
        fdn7 = fd_n7.get(tinh)
        rows.append({
            'Tinh'      : tinh,
            'FD_N'      : fdn,
            'FD_N1'     : fdn1,
            'vs_N1'     : (fdn - fdn1) if fdn is not None and fdn1 is not None else None,
            'FD_N7'     : fdn7,
            'vs_N7'     : (fdn - fdn7) if fdn is not None and fdn7 is not None else None,
            'VuotTarget': (fdn - 4.5) if fdn is not None else None,
        })

    # ── Tổng NTB (weighted avg toàn vùng theo %vol) ──────────
    def total_fd(date):
        d = pd.Timestamp(date).normalize()
        sub = df[df['date'] == d]
        if sub.empty or sub['vol'].sum() == 0:
            return None
        return (sub['vol'] * sub['ret']).sum() / sub['vol'].sum()

    fdn  = total_fd(latest)
    fdn1 = total_fd(n1)
    fdn7 = total_fd(n7)
    rows.append({
        'Tinh'      : 'Tổng NTB',
        'FD_N'      : fdn,
        'FD_N1'     : fdn1,
        'vs_N1'     : (fdn - fdn1) if fdn is not None and fdn1 is not None else None,
        'FD_N7'     : fdn7,
        'vs_N7'     : (fdn - fdn7) if fdn is not None and fdn7 is not None else None,
        'VuotTarget': (fdn - 4.5) if fdn is not None else None,
    })
    return pd.DataFrame(rows)

def build_trend_bc(df, bc_to_am, dates, bc_am_data=None):
    bc_am_data = bc_am_data or {}
    bc_am_from_raw = df[df['AM_raw'] != ''].drop_duplicates('BC').set_index('BC')['AM_raw'].to_dict()
    def get_am(bc):
        return bc_am_data.get(bc) or bc_am_from_raw.get(bc) or bc_to_am.get(bc, '')

    fd_cache = {}
    tt_cache = {}
    for d in dates:
        for dd in [d, d-timedelta(1), d-timedelta(7)]:
            if dd not in fd_cache:
                fd_cache[dd] = calc_fd_bc(df, dd)
        tt_cache[d] = calc_tytrong(df, d)

    all_bcs = []
    for d in dates:
        for bc in fd_cache[d]:
            if bc not in all_bcs: all_bcs.append(bc)

    # sort by latest date FD desc
    d0 = dates[0]
    all_bcs.sort(key=lambda bc: -fd_cache[d0].get(bc, 0))

    rows = []
    for bc in all_bcs:
        row = {'BC': bc, 'AM': get_am(bc)}
        for d in dates:
            dk  = d.strftime('%d/%m')
            fdn = fd_cache[d].get(bc)
            fn1 = fd_cache[d-timedelta(1)].get(bc)
            fn7 = fd_cache[d-timedelta(7)].get(bc)
            tt  = tt_cache[d].get(bc)
            row[f'FD_{dk}']  = fdn
            row[f'vN1_{dk}'] = (fdn - fn1) if fdn is not None and fn1 is not None else None
            row[f'vCK_{dk}'] = (fdn - fn7) if fdn is not None and fn7 is not None else None
            row[f'TT_{dk}']  = tt
        rows.append(row)
    return pd.DataFrame(rows)

def build_trend_am(df, bc_to_am, dates, bc_am_data=None):
    bc_am_data = bc_am_data or {}
    fd_cache = {}
    for d in dates:
        for dd in [d, d-timedelta(1), d-timedelta(7)]:
            if dd not in fd_cache:
                fd_cache[dd] = calc_fd_am(df, dd, bc_to_am, bc_am_data)

    all_ams = list(set(am for d in dates for am in fd_cache[d]))
    d0 = dates[0]
    all_ams.sort(key=lambda am: -fd_cache[d0].get(am, 0))

    rows = []
    for am in all_ams:
        row = {'AM': am}
        for d in dates:
            dk  = d.strftime('%d/%m')
            fdn = fd_cache[d].get(am)
            fn1 = fd_cache[d-timedelta(1)].get(am)
            fn7 = fd_cache[d-timedelta(7)].get(am)
            row[f'FD_{dk}']  = fdn
            row[f'vN1_{dk}'] = (fdn - fn1) if fdn is not None and fn1 is not None else None
            row[f'vCK_{dk}'] = (fdn - fn7) if fdn is not None and fn7 is not None else None
        rows.append(row)
    return pd.DataFrame(rows)

def build_trend_tinh(df, dates):
    fd_cache = {}
    for d in dates:
        for dd in [d, d-timedelta(1), d-timedelta(7)]:
            if dd not in fd_cache:
                fd_cache[dd] = calc_fd_tinh(df, dd)

    rows = []
    for tinh in TINH_LIST:
        row = {'Tinh': tinh}
        for d in dates:
            dk  = d.strftime('%d/%m')
            fdn = fd_cache[d].get(tinh)
            fn1 = fd_cache[d-timedelta(1)].get(tinh)
            fn7 = fd_cache[d-timedelta(7)].get(tinh)
            row[f'FD_{dk}']  = fdn
            row[f'vN1_{dk}'] = (fdn - fn1) if fdn is not None and fn1 is not None else None
            row[f'vCK_{dk}'] = (fdn - fn7) if fdn is not None and fn7 is not None else None
        rows.append(row)
    return pd.DataFrame(rows)

# ============================================================
#  WRITE TO GSHEET – HELPERS
# ============================================================
def rgb(clr_key):
    return CLR[clr_key]

def fmt_pct(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    return round(val / 100, 4)  # GSheet lưu số thực, format sau

def fmt_delta(val):
    return fmt_pct(val)

def cell_data(value, bg=None, bold=False, fg=None, fmt=None, halign='LEFT'):
    d = {'userEnteredValue': {}}
    if isinstance(value, str):
        d['userEnteredValue']['stringValue'] = value
    elif isinstance(value, (int, float)):
        d['userEnteredValue']['numberValue'] = value
    else:
        d['userEnteredValue']['stringValue'] = str(value) if value is not None else ''

    fmt_obj = {
        'textFormat': {
            'bold': bold,
            'foregroundColor': rgb(fg) if fg else rgb('black'),
            'fontFamily': 'Arial',
            'fontSize': 9,
        },
        'horizontalAlignment': halign,
        'verticalAlignment': 'MIDDLE',
    }
    if bg:
        fmt_obj['backgroundColor'] = rgb(bg)
    if fmt:
        fmt_obj['numberFormat'] = {'type': 'NUMBER', 'pattern': fmt}

    d['userEnteredFormat'] = fmt_obj
    return d

# ============================================================
#  WRITE SNAPSHOT SHEET
# ============================================================
def write_snapshot(sh, df, bc_to_am, latest, n1, n7, label, df_vol=None, bc_am_data=None,
                    tbl_bc=None, tbl_am=None, tbl_tinh=None):
    print(f'  Ghi sheet: {sh.title}')
    sh.clear()

    if tbl_bc is None:
        tbl_bc   = build_bc_table(df, bc_to_am, latest, n1, n7, df_vol, bc_am_data)
    if tbl_am is None:
        tbl_am   = build_am_table(df, bc_to_am, latest, n1, n7, df_vol, bc_am_data)
    if tbl_tinh is None:
        tbl_tinh = build_tinh_table(df, latest, n1, n7)

    rows_data = []

    # ── Row 1: info ──
    rows_data.append([
        cell_data(f'SNAPSHOT %FD – {label}', bg='darkBlue', bold=True, fg='white', halign='LEFT'),
        cell_data(f'N = {latest.strftime("%d/%m/%Y")}', bg='darkBlue', fg='white', halign='CENTER'),
        cell_data(f'N-1 = {n1.strftime("%d/%m")}',     bg='darkBlue', fg='white', halign='CENTER'),
        cell_data(f'N-7 = {n7.strftime("%d/%m")} (cùng thứ)', bg='darkBlue', fg='white', halign='CENTER'),
        cell_data('', bg='darkBlue'), cell_data('', bg='darkBlue'),
        cell_data('', bg='darkBlue'), cell_data('', bg='darkBlue'),
    ])

    rows_data.append([cell_data('')] * 8)

    # ── Bảng BC ──
    rows_data.append([
        cell_data('🏪 TẤT CẢ BƯU CỤC  (sort %FD ↓)  |  🔴 = FD>4.5% & Tỷ trọng>3%', bg='lightBlue', bold=True, halign='LEFT'),
        *[cell_data('', bg='lightBlue')] * 9
    ])
    rows_data.append([
        cell_data('Bưu Cục',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('AM',           bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N)',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-1)',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-1',       bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-7)',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-7',       bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('Vol giao',     bg='orange',  bold=True, fg='white', halign='CENTER'),
        cell_data('Vol trả',      bg='orange',  bold=True, fg='white', halign='CENTER'),
        cell_data('Tỷ trọng trả', bg='orange',  bold=True, fg='white', halign='CENTER'),
    ])

    for i, (_, r) in enumerate(tbl_bc.iterrows()):
        alt = i % 2 == 1
        bg  = 'altRow' if alt else None
        fdn  = r['FD_N']
        tt   = r['TyTrong']
        is_priority = (fdn is not None and fdn > 4.5 and tt is not None and tt > 3)
        row_bg = 'redBg' if is_priority else bg
        row_fg = 'redFont' if is_priority else None
        vol_giao = r.get('Vol_giao')
        vol_tra  = r.get('Vol_tra')

        rows_data.append([
            cell_data(r['BC'],  bg=row_bg, bold=is_priority, fg=row_fg, halign='LEFT'),
            cell_data(r['AM'],  bg=row_bg, halign='LEFT'),
            fd_cell(r['FD_N'],  alt),
            fd_cell(r['FD_N1'], alt),
            delta_cell(r['vs_N1']),
            fd_cell(r['FD_N7'], alt),
            delta_cell(r['vs_N7']),
            cell_data(round(vol_giao) if vol_giao is not None and not pd.isna(vol_giao) else '', bg=bg, fmt='#,##0', halign='CENTER'),
            cell_data(round(vol_tra,1) if vol_tra is not None and not pd.isna(vol_tra) else '', bg=bg, fmt='#,##0.0', halign='CENTER'),
            cell_data(fmt_pct(tt), bg='yellow', fmt='0.0%', halign='CENTER'),
        ])

    rows_data.append([cell_data('')] * 10)

    # ── Bảng AM ──
    rows_data.append([
        cell_data('👤 THEO AM – ngày N  (🔴 = FD>4.5% và Tỷ trọng>3%)', bg='lightBlue', bold=True, halign='LEFT'),
        *[cell_data('', bg='lightBlue')] * 6
    ])
    rows_data.append([
        cell_data('AM',           bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N)',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-1)',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-1',       bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-7)',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-7',       bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('Vol trả',      bg='orange',  bold=True, fg='white', halign='CENTER'),
        cell_data('Tỷ trọng trả', bg='orange',  bold=True, fg='white', halign='CENTER'),
    ])
    for i, (_, r) in enumerate(tbl_am.iterrows()):
        alt = i % 2 == 1
        bg  = 'altRow' if alt else None
        fdn  = r['FD_N']
        tt   = r.get('TyTrong')
        vol_tra = r.get('Vol_tra')
        is_priority = (fdn is not None and fdn > 4.5 and tt is not None and tt > 3)
        row_bg = 'redBg' if is_priority else bg
        row_fg = 'redFont' if is_priority else None
        rows_data.append([
            cell_data(r['AM'],             bg=row_bg, bold=is_priority, fg=row_fg, halign='LEFT'),
            fd_cell(r['FD_N'],  alt),
            fd_cell(r['FD_N1'], alt),
            delta_cell(r['vs_N1']),
            fd_cell(r['FD_N7'], alt),
            delta_cell(r['vs_N7']),
            cell_data(round(vol_tra,1) if vol_tra is not None and not pd.isna(vol_tra) else '', bg=bg, fmt='#,##0.0', halign='CENTER'),
            cell_data(fmt_pct(tt),         bg='yellow', fmt='0.0%', halign='CENTER'),
        ])

    rows_data.append([cell_data('')] * 8)

    # ── Bảng Tỉnh ──
    rows_data.append([
        cell_data('🗺️ THEO TỈNH – ngày N', bg='lightBlue', bold=True, halign='LEFT'),
        *[cell_data('', bg='lightBlue')] * 5
    ])
    rows_data.append([
        cell_data('Tỉnh',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N)',   bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-1)', bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-1',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('%FD (N-7)', bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('vs N-7',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
    ])
    for i, (_, r) in enumerate(tbl_tinh.iterrows()):
        is_total = (r['Tinh'] == 'Tổng NTB')
        alt = i % 2 == 1
        bg  = 'lightBlue' if is_total else ('altRow' if alt else None)
        rows_data.append([
            cell_data(r['Tinh'],           bg=bg, bold=is_total, halign='LEFT'),
            fd_cell(r['FD_N'],  alt) if not is_total else cell_data(fmt_pct(r['FD_N']), bg=bg, bold=True, fmt='0.0%', halign='CENTER'),
            fd_cell(r['FD_N1'], alt) if not is_total else cell_data(fmt_pct(r['FD_N1']), bg=bg, bold=True, fmt='0.0%', halign='CENTER'),
            delta_cell(r['vs_N1']),
            cell_data(fmt_pct(r['FD_N7']), bg=bg, bold=is_total, fmt='0.0%', halign='CENTER'),
            delta_cell(r['vs_N7']),
        ])

    batch_update_sheet(sh, rows_data)
    set_col_widths(sh, [300, 150, 90, 90, 80, 90, 80, 100])

# ============================================================
#  WRITE TREND SHEET
# ============================================================
def write_trend(sh, df, bc_to_am, dates, label, bc_am_data=None):
    print(f'  Ghi sheet: {sh.title}')
    sh.clear()

    tbl_bc   = build_trend_bc(df, bc_to_am, dates, bc_am_data)
    tbl_am   = build_trend_am(df, bc_to_am, dates, bc_am_data)
    tbl_tinh = build_trend_tinh(df, dates)

    rows_data = []

    # banner
    date_range = f'{dates[-1].strftime("%d/%m")} → {dates[0].strftime("%d/%m/%Y")}'
    rows_data.append([
        cell_data(f'TREND 8 NGÀY %FD – {label}  |  {date_range}  |  🔴 FD>4.5% & Tỷ trọng>3%',
                  bg='darkBlue', bold=True, fg='white', halign='LEFT'),
    ])
    rows_data.append([cell_data('')])

    # ── Bảng BC trend ──
    rows_data.append([
        cell_data('🏪 TREND BƯU CỤC (kèm AM + Tỷ trọng trả)', bg='lightBlue', bold=True, halign='LEFT'),
    ])

    # header ngày
    hdr_dates = [cell_data('Bưu Cục', bg='darkBlue', bold=True, fg='white', halign='CENTER'),
                 cell_data('AM',      bg='darkBlue', bold=True, fg='white', halign='CENTER')]
    WKDAY = ['T2','T3','T4','T5','T6','T7','CN']
    for d in dates:
        dk = d.strftime('%d/%m')
        wk = WKDAY[d.weekday()]
        hdr_dates += [
            cell_data(f'{dk}\n({wk})', bg='darkBlue', bold=True, fg='white', halign='CENTER'),
            cell_data('', bg='darkBlue'),
            cell_data('', bg='darkBlue'),
            cell_data('', bg='darkBlue'),
        ]
    rows_data.append(hdr_dates)

    # sub-header
    sub_hdr = [
        cell_data('Bưu Cục', bg='midBlue', bold=True, fg='white', halign='CENTER'),
        cell_data('AM',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
    ]
    for d in dates:
        sub_hdr += [
            cell_data('%FD',      bg='midBlue', bold=True, fg='white', halign='CENTER'),
            cell_data('vs N-1',   bg='midBlue', bold=True, fg='white', halign='CENTER'),
            cell_data('vs CK',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
            cell_data('Tỷ trọng', bg='orange',  bold=True, fg='white', halign='CENTER'),
        ]
    rows_data.append(sub_hdr)

    for i, (_, r) in enumerate(tbl_bc.iterrows()):
        alt = i % 2 == 1
        bg  = 'altRow' if alt else None
        d0  = dates[0].strftime('%d/%m')
        fdn0 = r.get(f'FD_{d0}')
        tt0  = r.get(f'TT_{d0}')
        is_p = (fdn0 is not None and fdn0 > 4.5 and tt0 is not None and tt0 > 3)
        lbg  = 'redBg' if is_p else bg
        lfg  = 'redFont' if is_p else None

        row = [
            cell_data(r['BC'], bg=lbg, bold=is_p, fg=lfg, halign='LEFT'),
            cell_data(r['AM'], bg=lbg, halign='LEFT'),
        ]
        for d in dates:
            dk = d.strftime('%d/%m')
            fdn = r.get(f'FD_{dk}')
            vn1 = r.get(f'vN1_{dk}')
            vck = r.get(f'vCK_{dk}')
            tt  = r.get(f'TT_{dk}')
            row += [
                fd_cell(fdn, alt),
                delta_cell(vn1),
                delta_cell(vck),
                cell_data(fmt_pct(tt),  bg='yellow', fmt='0.0%', halign='CENTER'),
            ]
        rows_data.append(row)

    rows_data.append([cell_data('')])

    # ── Bảng AM trend ──
    rows_data.append([cell_data('👤 TREND THEO AM', bg='lightBlue', bold=True, halign='LEFT')])

    hdr2 = [cell_data('AM', bg='darkBlue', bold=True, fg='white', halign='CENTER')]
    for d in dates:
        dk = d.strftime('%d/%m'); wk = WKDAY[d.weekday()]
        hdr2 += [cell_data(f'{dk}\n({wk})', bg='darkBlue', bold=True, fg='white', halign='CENTER'),
                 cell_data('', bg='darkBlue'), cell_data('', bg='darkBlue')]
    rows_data.append(hdr2)

    sub2 = [cell_data('AM', bg='midBlue', bold=True, fg='white', halign='CENTER')]
    for d in dates:
        sub2 += [cell_data('%FD',    bg='midBlue', bold=True, fg='white', halign='CENTER'),
                 cell_data('vs N-1', bg='midBlue', bold=True, fg='white', halign='CENTER'),
                 cell_data('vs CK',  bg='midBlue', bold=True, fg='white', halign='CENTER')]
    rows_data.append(sub2)

    for i, (_, r) in enumerate(tbl_am.iterrows()):
        alt = i % 2 == 1; bg = 'altRow' if alt else None
        row = [cell_data(r['AM'], bg=bg, halign='LEFT')]
        for d in dates:
            dk = d.strftime('%d/%m')
            row += [fd_cell(r.get(f'FD_{dk}'), alt),
                    delta_cell(r.get(f'vN1_{dk}')), delta_cell(r.get(f'vCK_{dk}'))]
        rows_data.append(row)

    rows_data.append([cell_data('')])

    # ── Bảng Tỉnh trend ──
    rows_data.append([cell_data('🗺️ TREND THEO TỈNH', bg='lightBlue', bold=True, halign='LEFT')])
    rows_data.append(hdr2[:1] + hdr2[1:])  # reuse AM headers
    rows_data.append(sub2[:1] + sub2[1:])

    for i, (_, r) in enumerate(tbl_tinh.iterrows()):
        alt = i % 2 == 1; bg = 'altRow' if alt else None
        row = [cell_data(r['Tinh'], bg=bg, halign='LEFT')]
        for d in dates:
            dk = d.strftime('%d/%m')
            row += [fd_cell(r.get(f'FD_{dk}'), alt),
                    delta_cell(r.get(f'vN1_{dk}')), delta_cell(r.get(f'vCK_{dk}'))]
        rows_data.append(row)

    batch_update_sheet(sh, rows_data)

    # col widths: BC=280, AM=130, sau đó mỗi ngày 4 col x 80
    widths = [280, 130] + [80] * (len(dates) * 4 + 10)
    set_col_widths(sh, widths)

# ============================================================
#  GSHEET BATCH HELPERS
# ============================================================
# ── 3 mức FD ────────────────────────────────────────────────
# ≤4.5% = Tốt (xanh), 4.5-6% = Trung bình (vàng), ≥6% = Cần cải thiện (đỏ)
CLR['good']    = {'red': 0.851, 'green': 0.918, 'blue': 0.827}   # xanh nhạt
CLR['goodFont']= {'red': 0.153, 'green': 0.392, 'blue': 0.098}
CLR['warn']    = {'red': 1.000, 'green': 0.949, 'blue': 0.800}   # vàng
CLR['warnFont']= {'red': 0.600, 'green': 0.400, 'blue': 0.000}
CLR['bad']     = {'red': 1.000, 'green': 0.878, 'blue': 0.878}   # đỏ nhạt
CLR['badFont'] = {'red': 0.753, 'green': 0.000, 'blue': 0.000}

def fd_cell(val, alt=False):
    """Cell cho %FD với 3 mức màu"""
    bg_alt = 'altRow' if alt else None
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return cell_data('', bg=bg_alt, halign='CENTER')
    v = round(val / 100, 4)
    if val <= 4.5:
        return cell_data(v, bg='good',  fg='goodFont', fmt='0.0%', halign='CENTER')
    elif val < 6.0:
        return cell_data(v, bg='warn',  fg='warnFont', fmt='0.0%', halign='CENTER')
    else:
        return cell_data(v, bg='bad',   fg='badFont',  fmt='0.0%', halign='CENTER')

def delta_cell(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return cell_data('', halign='CENTER')
    v = round(val / 100, 4)
    if val > 0.05:
        return cell_data(v, bg='bad',   fg='badFont',  fmt='▲0.0%;▼0.0%;-', halign='CENTER')
    elif val < -0.05:
        return cell_data(v, bg='good',  fg='goodFont', fmt='▲0.0%;▼0.0%;-', halign='CENTER')
    else:
        return cell_data(v, fmt='▲0.0%;▼0.0%;-', halign='CENTER')

def batch_update_sheet(sh, rows_data):
    """Ghi toàn bộ rows_data vào sheet bằng batchUpdate"""
    max_cols = max(len(r) for r in rows_data)
    # pad rows
    for r in rows_data:
        while len(r) < max_cols:
            r.append(cell_data(''))

    body = {
        'requests': [{
            'updateCells': {
                'rows': [{'values': row} for row in rows_data],
                'fields': 'userEnteredValue,userEnteredFormat',
                'start': {'sheetId': sh.id, 'rowIndex': 0, 'columnIndex': 0}
            }
        }]
    }
    sh.spreadsheet.batch_update(body)

def set_col_widths(sh, widths):
    # Giới hạn số cột để tránh vượt quá giới hạn sheet, và bọc try/except
    # để lỗi set width không làm crash toàn bộ script
    max_cols = 50
    widths = widths[:max_cols]
    requests = []
    for i, w in enumerate(widths):
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sh.id,
                    'dimension': 'COLUMNS',
                    'startIndex': i,
                    'endIndex': i+1,
                },
                'properties': {'pixelSize': w},
                'fields': 'pixelSize'
            }
        })
    try:
        sh.spreadsheet.batch_update({'requests': requests})
    except Exception as e:
        print(f'  ⚠️ Không set được column widths: {e}')

def ensure_sheet(spreadsheet, name):
    try:
        return spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=200, cols=50)

# ============================================================
#  TELEGRAM REPORT – tạo ảnh ghép 3 bảng + gửi
# ============================================================
def fd_color(val):
    """Trả về màu nền cho %FD: xanh <=4.5, vàng 4.5-6, đỏ >=6"""
    if val is None or pd.isna(val):
        return '#FFFFFF'
    if val <= 4.5:
        return '#D9EAD3'
    elif val < 6.0:
        return '#FFF2CC'
    else:
        return '#FFE0E0'

def fd_textcolor(val):
    if val is None or pd.isna(val):
        return '#000000'
    if val <= 4.5:
        return '#274E13'
    elif val < 6.0:
        return '#7F6000'
    else:
        return '#C00000'

def delta_str(val):
    if val is None or pd.isna(val):
        return '-'
    arrow = '▲' if val > 0 else ('▼' if val < 0 else '-')
    return f'{arrow}{abs(val):.1f}%'

def delta_color(val):
    if val is None or pd.isna(val):
        return ('#FFFFFF', '#000000')
    if val > 0.05:
        return ('#FFE0E0', '#C00000')
    elif val < -0.05:
        return ('#E2EFDA', '#375623')
    return ('#FFFFFF', '#000000')


def draw_table(ax, df, columns, col_labels, title, col_widths=None,
                fd_cols=None, delta_cols=None, highlight_cols=None,
                highlight_color='#FFE69C', bold_rows=None):
    """Vẽ 1 bảng vào ax (matplotlib axes)"""
    fd_cols        = fd_cols or []
    delta_cols     = delta_cols or []
    highlight_cols = highlight_cols or []
    bold_rows      = bold_rows or []

    ax.axis('off')
    ax.set_title(title, fontsize=22, fontweight='bold', loc='left', pad=18,
                  color='#1F4E79')

    n_rows = len(df) + 1  # +1 header
    n_cols = len(columns)

    cell_text  = []
    cell_color = []
    text_color = []

    # Header row
    cell_text.append(col_labels)
    cell_color.append(['#2E75B6']*n_cols)
    text_color.append(['white']*n_cols)

    for ridx, (_, row) in enumerate(df.iterrows()):
        is_bold_row = ridx in bold_rows
        texts = []; colors = []; tcolors = []
        for c in columns:
            val = row.get(c)
            if c in fd_cols:
                texts.append(f'{val:.1f}%' if val is not None and not pd.isna(val) else '-')
                colors.append(fd_color(val))
                tcolors.append(fd_textcolor(val))
            elif c in delta_cols:
                texts.append(delta_str(val))
                bg, fg = delta_color(val)
                colors.append(bg); tcolors.append(fg)
            elif c in highlight_cols:
                if isinstance(val, float) and not pd.isna(val):
                    if 'TyTrong' in c or 'Trong' in c:
                        texts.append(f'{val:.1f}%')
                    else:
                        texts.append(f'{val:,.1f}')
                else:
                    texts.append(str(val) if val is not None else '-')
                colors.append(highlight_color); tcolors.append('#000000')
            else:
                texts.append(str(val) if val is not None else '-')
                colors.append('#F5F9FF' if ridx%2==1 else '#FFFFFF')
                tcolors.append('#000000')
        if is_bold_row:
            colors = ['#D6E4F0']*n_cols
        cell_text.append(texts)
        cell_color.append(colors)
        text_color.append(tcolors)

    tbl = ax.table(cellText=cell_text, cellColours=cell_color,
                    cellLoc='center', loc='upper left',
                    colWidths=col_widths,
                    bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(17)

    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor('#BFBFBF')
        cell.set_text_props(color=text_color[r][c])
        if r == 0:
            cell.set_text_props(weight='bold', color='white')
        if r-1 in bold_rows and r > 0:
            cell.set_text_props(weight='bold')
        if c == 0:
            cell.set_text_props(ha='left')
            cell._loc = 'left'


def render_fd_report(tbl_bc, tbl_am, tbl_tinh, label, date_str):
    """Vẽ ảnh ghép 3 bảng: Top10 BC, AM, Tỉnh"""
    n_bc   = min(10, len(tbl_bc)) + 1   # +1 header
    n_am   = len(tbl_am) + 1
    n_tinh = len(tbl_tinh) + 1
    total_rows = n_bc + n_am + n_tinh
    fig_h = max(19, total_rows * 0.90)

    fig = plt.figure(figsize=(16, fig_h), dpi=150)
    gs  = fig.add_gridspec(3, 1, height_ratios=[n_bc, n_am, n_tinh], hspace=0.32,
                            top=0.93, bottom=0.02, left=0.03, right=0.97)

    fig.suptitle(f'📊 BÁO CÁO %FD – {label}  |  {date_str}', fontsize=26,
                  fontweight='bold', color='#1F4E79', y=0.985)

    # ── Top 10 BC theo %FD (đổi từ tỷ trọng) ──
    top10 = tbl_bc.sort_values('FD_N', ascending=False, na_position='last').head(10).copy()
    ax1 = fig.add_subplot(gs[0])
    cols   = ['BC','AM','FD_N','FD_N1','vs_N1','FD_N7','vs_N7','TyTrong']
    labels = ['Bưu Cục','AM','%FD (N)','%FD (N-1)','vs N-1','%FD (N-7)','vs N-7','Tỷ trọng']
    draw_table(ax1, top10, cols, labels,
               '🏪 TOP 10 BƯU CỤC THEO %FD',
               col_widths=[0.30,0.14,0.09,0.09,0.08,0.09,0.08,0.09],
               fd_cols=['FD_N','FD_N1','FD_N7'],
               delta_cols=['vs_N1','vs_N7'],
               highlight_cols=['TyTrong'])

    # ── AM ──
    am_sorted = tbl_am.sort_values('TyTrong', ascending=False, na_position='last').copy()
    ax2 = fig.add_subplot(gs[1])
    cols2   = ['AM','FD_N','FD_N1','vs_N1','FD_N7','vs_N7','Vol_tra','TyTrong']
    labels2 = ['AM','%FD (N)','%FD (N-1)','vs N-1','%FD (N-7)','vs N-7','Vol trả','Tỷ trọng']
    draw_table(ax2, am_sorted, cols2, labels2,
               '👤 TẤT CẢ AM',
               col_widths=[0.22,0.10,0.10,0.09,0.10,0.09,0.10,0.10],
               fd_cols=['FD_N','FD_N1','FD_N7'],
               delta_cols=['vs_N1','vs_N7'],
               highlight_cols=['Vol_tra','TyTrong'])

    # ── Tỉnh ──
    ax3 = fig.add_subplot(gs[2])
    cols3   = ['Tinh','FD_N','FD_N1','vs_N1','FD_N7','vs_N7']
    labels3 = ['Tỉnh','%FD (N)','%FD (N-1)','vs N-1','%FD (N-7)','vs N-7']
    bold_rows = [len(tbl_tinh)-1]  # dòng cuối = Tổng NTB
    draw_table(ax3, tbl_tinh, cols3, labels3,
               '🗺️ THEO TỈNH',
               col_widths=[0.22,0.13,0.13,0.13,0.13,0.13],
               fd_cols=['FD_N','FD_N1','FD_N7'],
               delta_cols=['vs_N1','vs_N7'],
               bold_rows=bold_rows)

    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf


def send_telegram_photo(image_buf, caption):
    url = f'https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto'
    files = {'photo': ('report.png', image_buf, 'image/png')}
    data  = {'chat_id': TELEGRAM_CHAT_ID, 'caption': caption}
    resp  = requests.post(url, files=files, data=data, timeout=60)
    if resp.status_code == 200:
        print(f'  ✅ Đã gửi Telegram: {caption[:50]}...')
    else:
        print(f'  ⚠️ Lỗi gửi Telegram ({resp.status_code}): {resp.text[:200]}')


# ============================================================
#  GTALK SENDER
# ============================================================
def send_photo_gtalk(image_buf, caption=""):
    if not GTALK_OA_TOKEN or not GTALK_CHANNEL_ID:
        print("⚠️ Không tìm thấy GTALK_OA_TOKEN hoặc GTALK_CHANNEL_ID. Bỏ qua gửi GTalk.")
        return False

    print("📡 Đang gửi ảnh báo cáo sang GTalk...")
    try:
        image_buf.seek(0)
        img = Image.open(image_buf)
        width, height = img.size
        file_size = len(image_buf.getvalue())
        image_buf.seek(0)
    except Exception as e:
        print(f"❌ Lỗi đọc ảnh: {e}")
        return False

    # Step 1: Initiate Upload
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": "report.png",
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": f'{{"width": {width}, "height": {height}}}',
        "oaToken": GTALK_OA_TOKEN
    }
    headers = {"Content-Type": "application/json"}
    try:
        res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
        if res_init.status_code != 200:
            print(f"❌ Lỗi initiate upload HTTP {res_init.status_code}: {res_init.text}")
            return False
        res_data = res_init.json()
        if res_data.get("errorCode") != "success":
            print(f"❌ Lỗi initiate upload API: {res_data.get('error')}")
            return False
        
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi initiate upload GTalk: {e}")
        return False

    # Step 2: Upload to S3
    try:
        image_buf.seek(0)
        headers_put = {"Content-Type": "image/png"}
        res_put = requests.put(presigned_url, data=image_buf, headers=headers_put, timeout=60, verify=False)
        if res_put.status_code != 200:
            print(f"❌ Lỗi PUT lên S3 HTTP {res_put.status_code}: {res_put.text}")
            return False
    except Exception as e:
        print(f"❌ Lỗi upload file lên S3 GTalk: {e}")
        return False

    # Step 3: Complete Upload
    complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
    payload_complete = {
        "oaToken": GTALK_OA_TOKEN,
        "UploadId": upload_id
    }
    try:
        res_comp = requests.post(complete_url, json=payload_complete, headers=headers, timeout=20, verify=False)
        if res_comp.status_code != 200:
            print(f"❌ Lỗi complete upload HTTP {res_comp.status_code}: {res_comp.text}")
            return False
        res_data_comp = res_comp.json()
        if res_data_comp.get("errorCode") != "success":
            print(f"❌ Lỗi complete upload API: {res_data_comp.get('error')}")
            return False
        file_id = res_data_comp["data"]["Id"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi complete upload GTalk: {e}")
        return False

    # Step 4: Send Message
    send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    payload_send = {
        "channelId": GTALK_CHANNEL_ID,
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": caption,
                "items": [
                    {
                        "image": {
                            "fileId": file_id,
                            "width": width,
                            "height": height
                        }
                    }
                ]
            }
        },
        "oaToken": GTALK_OA_TOKEN
    }
    try:
        res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
        if res_send.status_code == 200:
            res_data_send = res_send.json()
            if res_data_send.get("errorCode") == "success":
                print("  ✅ Đã gửi ảnh sang GTalk thành công!")
                return True
            else:
                print(f"  ❌ Lỗi gửi tin nhắn GTalk API: {res_data_send.get('error')}")
        else:
            print(f"  ❌ Lỗi HTTP {res_send.status_code}: {res_send.text}")
    except Exception as e:
        print(f"  ❌ Lỗi kết nối khi gửi tin nhắn GTalk: {e}")
    return False


def main():
    print('🔐 Kết nối Google Sheets...')
    gc = get_gspread_client(SPREADSHEET_ID)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    print(f'✅ Đã kết nối: {spreadsheet.title}')

    print('👥 Đọc CoCauVung...')
    bc_to_am, bc_to_tinh = load_cocau(spreadsheet)
    print(f'   {len(bc_to_am)} BC có AM mapping')

    # Debug: kiểm tra BC nào chưa có AM sau khi load raw đầu tiên
    df_check = load_raw(spreadsheet, RAW_SHEETS['TONG'], bc_to_tinh)
    if not df_check.empty:
        missing = [bc for bc in df_check['BC'].unique() if bc not in bc_to_am]
        if missing:
            print(f'  ⚠️ {len(missing)} BC chưa có AM trong CoCauVung:')
            for bc in sorted(missing):
                print(f'     - [{bc}]')
        else:
            print('  ✅ Tất cả BC đều có AM mapping!')

    # Load data sheet (Vol Ca1+Ca2)
    print('\n📦 Đọc data sheet (Vol giao Ca1+Ca2)...')
    df_vol, bc_am_data = load_data_sheet(spreadsheet)

    CONFIGS = [
        ('TONG', 'Snapshot – FD Tổng', 'Trend – FD Tổng'),
        ('COD',  'Snapshot – FD COD',  'Trend – FD COD'),
        ('TTS',  'Snapshot – FD TTS',  'Trend – FD TTS'),
    ]

    for key, snap_name, trend_name in CONFIGS:
        print(f'\n📊 Xử lý {key}...')
        df = load_raw(spreadsheet, RAW_SHEETS[key], bc_to_tinh)
        if df.empty:
            print(f'  ⚠️ Sheet {RAW_SHEETS[key]} trống, bỏ qua')
            continue

        latest = df['date'].max()
        n1     = latest - timedelta(days=1)
        n7     = latest - timedelta(days=7)
        dates_8 = list(reversed(sorted(df['date'].unique())[-8:]))
        dates_8 = [pd.Timestamp(d).normalize() for d in dates_8]

        label = key

        print(f'  📅 Latest: {latest.strftime("%d/%m/%Y")} | N-1: {n1.strftime("%d/%m")} | N-7: {n7.strftime("%d/%m")}')

        snap_sh  = ensure_sheet(spreadsheet, snap_name)
        trend_sh = ensure_sheet(spreadsheet, trend_name)

        # Build các bảng 1 lần, dùng chung cho GSheet + ảnh Telegram
        tbl_bc   = build_bc_table(df, bc_to_am, latest, n1, n7, df_vol, bc_am_data)
        tbl_am   = build_am_table(df, bc_to_am, latest, n1, n7, df_vol, bc_am_data)
        tbl_tinh = build_tinh_table(df, latest, n1, n7)

        write_snapshot(snap_sh,  df, bc_to_am, latest, n1, n7, label, df_vol, bc_am_data,
                        tbl_bc=tbl_bc, tbl_am=tbl_am, tbl_tinh=tbl_tinh)
        write_trend   (trend_sh, df, bc_to_am, dates_8, label, bc_am_data)

        # ── Gửi báo cáo Telegram & GTalk ──────────────────────
        try:
            print(f'  🖼️ Tạo ảnh báo cáo {label}...')
            date_str = latest.strftime('%d/%m/%Y')
            label_map = {'TONG':'FD TỔNG (COD+NON COD)', 'COD':'FD COD', 'TTS':'FD TTS'}
            buf = render_fd_report(tbl_bc, tbl_am, tbl_tinh, label_map.get(label,label), date_str)
            caption = f'📊 BÁO CÁO %FD – {label_map.get(label,label)}\n📅 Ngày: {date_str}\n\n' \
                      f'%Return = Số đơn trả / Tổng đơn kết thúc giao\n' \
                      f'Trong đó:\n\n' \
                      f'Mẫu số: Đơn GTC + Đơn chuyển trả (chờ/đang/đã trả)\n' \
                      f'Tử số: Đơn chuyển trả (chờ/đang/đã trả)\n' \
                      f'Thời điểm tính: Theo ngày kết thúc giao'
            
            # Gửi Telegram
            buf.seek(0)
            send_telegram_photo(buf, caption)
            
            # Gửi GTalk
            buf.seek(0)
            send_photo_gtalk(buf, caption)
        except Exception as e:
            print(f'  ⚠️ Lỗi tạo/gửi ảnh {label}: {e}')

    print('\n✅ Hoàn tất! Mở Google Sheet để xem kết quả.')

if __name__ == '__main__':
    main()
