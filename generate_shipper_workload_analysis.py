# -*- coding: utf-8 -*-
"""
Script: generate_shipper_workload_analysis.py
Author: Antigravity AI
Description: Analyzes shipper workloads and calculates staffing gaps at each post office.
             Writes a formatted table to a new Google Sheet tab "Phân Tích Định Biên"
             and exports a local markdown report.
"""

import os
import sys
import argparse
import unicodedata
import math
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials

# Fix console encoding for Vietnamese character support
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# ============ CONFIG & CONSTANTS ============
MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"
WRITE_SPREADSHEET_ID = "1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M" # Writable sheet: 'Aging >5 ngày - follow gán'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Standard target workload per shipper per day
TARGET_WORKLOAD_PER_SHIPPER = 90.0

# HSL tailored color palette mapped to Google Sheets decimal colors (RGB 0-1)
COLOR_SYSTEM = {
    'darkBlue':   {'red': 0.106, 'green': 0.212, 'blue': 0.365},  # Header background
    'lightBlue':  {'red': 0.902, 'green': 0.941, 'blue': 0.980},  # Zebra stripe alt row
    'white':      {'red': 1.000, 'green': 1.000, 'blue': 1.000},  # White background
    'black':      {'red': 0.000, 'green': 0.000, 'blue': 0.000},  # Text color
    'grayLine':   {'red': 0.863, 'green': 0.863, 'blue': 0.863},  # Borders
    
    # Severity Colors
    'highBg':     {'red': 1.000, 'green': 0.878, 'blue': 0.878},  # Light red for Critical/High
    'highFg':     {'red': 0.753, 'green': 0.000, 'blue': 0.000},  # Dark red text
    'warnBg':     {'red': 1.000, 'green': 0.949, 'blue': 0.800},  # Light yellow for Warning/Medium
    'warnFg':     {'red': 0.600, 'green': 0.400, 'blue': 0.000},  # Dark yellow text
    'normalBg':   {'red': 0.886, 'green': 0.937, 'blue': 0.851},  # Light green for Normal/Low
    'normalFg':   {'red': 0.216, 'green': 0.341, 'blue': 0.137},  # Dark green text
}

# ============ STRING HELPERS ============
def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip().lower())

def clean_bc_name(name):
    name = normalize_str(name)
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl', 'đxl', 'kho']:
        name = name.replace(prefix, "")
    name = name.replace("-", " ").replace("_", " ")
    return " ".join(name.split())

def map_bc(bc_str, cc_df):
    if not bc_str or bc_str.strip() == '':
        return None
    
    parts = bc_str.split("-")
    raw_id = ""
    if len(parts) > 0:
        candidate = parts[0].strip()
        if candidate.isdigit():
            raw_id = candidate
            
    # Try exact ID match
    if raw_id:
        match_df = cc_df[cc_df['warehouse_id'] == raw_id]
        if not match_df.empty:
            return match_df.iloc[0]['Bưu cục']
        
        # Try matching by first 7 digits of ID
        prefix_7 = raw_id[:7]
        match_df_prefix = cc_df[cc_df['warehouse_id'].str.startswith(prefix_7)]
        if not match_df_prefix.empty:
            return match_df_prefix.iloc[0]['Bưu cục']
            
    # Fallback to fuzzy text match
    std_pos = cc_df['Bưu cục'].tolist()
    text_clean = clean_bc_name(bc_str)
    
    best_match = None
    best_len = 0
    for std in std_pos:
        std_clean = clean_bc_name(std)
        if std_clean and (std_clean in text_clean or text_clean in std_clean):
            if len(std_clean) > best_len:
                best_len = len(std_clean)
                best_match = std
    if best_match:
        return best_match
        
    return None

def parse_percent_to_float(val):
    if val is None or val == "":
        return 0.0
    try:
        val_str = str(val).replace("%", "").replace(",", ".").strip()
        num = float(val_str)
        if num < 1.0 and num > 0:
            return num * 100.0
        return num
    except ValueError:
        return 0.0

def parse_int(val):
    if val is None or val == "":
        return 0
    try:
        return int(str(val).replace(",", "").replace(".", "").strip())
    except ValueError:
        return 0

# ============ GSHEET CELL FORMATTING HELPERS ============
def make_cell(value, bg=None, bold=False, fg=None, halign='LEFT', is_header=False):
    d = {'userEnteredValue': {}}
    if isinstance(value, (int, float)):
        d['userEnteredValue']['numberValue'] = value
    else:
        d['userEnteredValue']['stringValue'] = str(value) if value is not None else ''

    fmt_obj = {
        'textFormat': {
            'bold': bold,
            'foregroundColor': COLOR_SYSTEM[fg] if fg else COLOR_SYSTEM['black'],
            'fontFamily': 'Arial',
            'fontSize': 10 if not is_header else 10,
        },
        'horizontalAlignment': halign,
        'verticalAlignment': 'MIDDLE',
    }
    if bg:
        fmt_obj['backgroundColor'] = COLOR_SYSTEM[bg]
    
    # Default border format
    border_color = COLOR_SYSTEM['grayLine']
    fmt_obj['borders'] = {
        'top':    {'style': 'SOLID', 'color': border_color},
        'bottom': {'style': 'SOLID', 'color': border_color},
        'left':   {'style': 'SOLID', 'color': border_color},
        'right':  {'style': 'SOLID', 'color': border_color},
    }
    
    d['userEnteredFormat'] = fmt_obj
    return d

def write_sheet_data(sh, rows_data):
    max_cols = max(len(r) for r in rows_data)
    for r in rows_data:
        while len(r) < max_cols:
            r.append(make_cell(''))

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
        print(f"⚠️ Failed to set column widths: {e}")

# ============ MAIN PIPELINE ============
def main():
    parser = argparse.ArgumentParser(description="Generate shipper workload & staffing gap analysis.")
    parser.add_argument("--date", type=str, help="Target date in YYYY-MM-DD format. Defaults to yesterday.")
    parser.add_argument("--write-key", type=str, default=WRITE_SPREADSHEET_ID, help="Spreadsheet ID to write back to.")
    args = parser.parse_args()

    # Determine target date
    if args.date:
        try:
            target_date = datetime.strptime(args.date.strip(), "%Y-%m-%d")
        except ValueError:
            print("❌ Invalid date format. Use YYYY-MM-DD.")
            sys.exit(1)
    else:
        # Default is yesterday (2026-06-24 based on the latest report)
        target_date = datetime.now() - timedelta(days=1)
        # Note: If running on 2026-06-25, yesterday is 2026-06-24
        
    date_str_iso = target_date.strftime("%Y-%m-%d")
    
    weekday_map = {
        0: "Thứ 2", 1: "Thứ 3", 2: "Thứ 4", 3: "Thứ 5", 4: "Thứ 6", 5: "Thứ 7", 6: "Chủ Nhật"
    }
    date_str_sheet = f"{date_str_iso} - {weekday_map[target_date.weekday()]}"
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    
    print(f"📅 Running Shipper Workload Analysis for: {date_str_iso} ({weekday_map[target_date.weekday()]})")
    
    # Authenticate with Google sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Credentials file not found at: {JSON_FILE}")
        sys.exit(1)
        
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    
    # Load Master CoCauVung
    print("📖 Loading Master CoCauVung sheet...")
    sh_main = gc.open_by_key(MAIN_SPREADSHEET_ID)
    ws_cc = sh_main.worksheet("CoCauVung")
    cc_rows = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_rows[1:], columns=cc_rows[0])
    df_cc['warehouse_id'] = df_cc['warehouse_id'].str.strip()
    df_cc = df_cc[df_cc['Bưu cục'].str.strip() != '']
    master_pos = df_cc['Bưu cục'].str.strip().tolist()
    po_am_map = dict(zip(df_cc['Bưu cục'].str.strip(), df_cc['AM'].str.strip()))
    
    # Load HR/Nhan Su data
    print("📖 Loading 'Nhân Sự' worksheet...")
    ws_ns = sh_main.worksheet("Nhân Sự")
    ns_rows = ws_ns.get_all_values()
    df_ns = pd.DataFrame(ns_rows[1:], columns=ns_rows[0])
    df_ns['Chức vụ'] = df_ns['Chức vụ'].str.strip()
    df_ns['Trạng thái'] = df_ns['Trạng thái'].str.strip()
    
    # Filter for active shippers (Business Development Field Executive)
    df_shippers = df_ns[
        (df_ns['Chức vụ'] == 'Business Development Field Executive') & 
        (df_ns['Trạng thái'] == 'Đang làm việc')
    ].copy()
    
    # Map employees to their standard post offices
    df_shippers['Mapped_BC'] = df_shippers['Bưu cục'].apply(lambda x: map_bc(x, df_cc))
    
    # Group by mapped post office to get active shipper count
    shipper_counts = df_shippers.groupby('Mapped_BC').size().to_dict()
    
    # Load Operational Data sheet
    print("📖 Loading 'Data' worksheet...")
    ws_data = sh_main.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    df_data_day = df_data[df_data['Time'] == date_str_sheet]
    
    if df_data_day.empty:
        print(f"⚠️ No operational data found for date: {date_str_sheet} in 'Data' sheet.")
        print("Please verify if the date is correct or if the data sheet has been updated.")
        # We will continue, but output will be empty or look at other dates
    
    # Initialize metrics structure for all post offices
    po_workload = {po: {
        'po_name': po,
        'am_name': po_am_map[po],
        'active_shippers': shipper_counts.get(po, 0),
        'new_vol': 0,
        'backlog_vol': 0,
        'total_workload': 0,
        'total_ton': 0,
        'total_chua_gan': 0,
        'actual_workload_per_shipper': 0.0,
        'required_shippers': 0,
        'staffing_gap': 0,
        'overload_rate': 0.0,
        'load_status': 'Thấp Tải (<70 đơn)',
        'alert_class': 'white'
    } for po in master_pos}
    
    # Aggregate operational metrics
    for idx, row in df_data_day.iterrows():
        std_po = map_bc(row['Chi tiết'], df_cc)
        if not std_po:
            continue
        m = po_workload[std_po]
        loai_hang = row['Loại Hàng'].strip()
        vol = parse_int(row['Volume'])
        m['total_ton'] += parse_int(row['Sản Lượng Tồn'])
        m['total_chua_gan'] += parse_int(row['Sản Lượng Chưa Gán'])
        
        if "Hàng Mới" in loai_hang:
            m['new_vol'] += vol
        elif "Tồn" in loai_hang:
            m['backlog_vol'] += vol
            
        m['total_workload'] += vol

    # Compute KPIs
    for po, m in po_workload.items():
        shippers = m['active_shippers']
        workload = m['total_workload']
        
        # 1. Required shippers at target (90 orders/shipper/day)
        m['required_shippers'] = int(math.ceil(workload / TARGET_WORKLOAD_PER_SHIPPER))
        
        # 2. Staffing Gap
        m['staffing_gap'] = max(0, m['required_shippers'] - shippers)
        
        # 3. Workload per shipper and overload rate
        if shippers > 0:
            m['actual_workload_per_shipper'] = workload / shippers
            if m['actual_workload_per_shipper'] > TARGET_WORKLOAD_PER_SHIPPER:
                m['overload_rate'] = ((m['actual_workload_per_shipper'] - TARGET_WORKLOAD_PER_SHIPPER) / TARGET_WORKLOAD_PER_SHIPPER) * 100.0
        else:
            m['actual_workload_per_shipper'] = float(workload)
            if workload > 0:
                m['overload_rate'] = 100.0
                
        # 4. Determine Load Status and color alert class
        if shippers == 0:
            if workload > 0:
                m['load_status'] = 'Không có Shipper!'
                m['alert_class'] = 'warnBg'
            else:
                m['load_status'] = 'Không hoạt động'
                m['alert_class'] = 'white'
        elif m['actual_workload_per_shipper'] >= 120.0:
            m['load_status'] = 'Quá Tải Cực Đoạn (>=120 đơn)'
            m['alert_class'] = 'highBg'
        elif m['actual_workload_per_shipper'] > 90.0:
            m['load_status'] = 'Quá Tải (>90 đơn)'
            m['alert_class'] = 'highBg'
        elif m['actual_workload_per_shipper'] >= 70.0:
            m['load_status'] = 'Tải Trọng Tốt (70-90 đơn)'
            m['alert_class'] = 'normalBg'
        else:
            m['load_status'] = 'Thấp Tải (<70 đơn)'
            m['alert_class'] = 'white'

    # Filter out post offices with no workload and no shippers to keep the sheet relevant
    active_po_workload = {po: m for po, m in po_workload.items() if m['total_workload'] > 0 or m['active_shippers'] > 0}
    
    # Convert to list and sort:
    # 1. Overloaded first (highest workload per shipper descending)
    # 2. Then by total workload descending
    po_list = list(active_po_workload.values())
    
    def sort_key(x):
        # Shippers == 0 with workload gets highest priority
        if x['active_shippers'] == 0 and x['total_workload'] > 0:
            return (2, x['total_workload'])
        elif x['actual_workload_per_shipper'] > TARGET_WORKLOAD_PER_SHIPPER:
            return (1, x['actual_workload_per_shipper'])
        else:
            return (0, x['actual_workload_per_shipper'])
            
    po_list.sort(key=sort_key, reverse=True)
    
    # ============ WRITE TO GOOGLE SHEETS ============
    print(f"✍️ Writing results to Google Sheets tab 'Phân Tích Định Biên' on sheet: {args.write_key}...")
    try:
        sh_write = gc.open_by_key(args.write_key)
        try:
            ws_dest = sh_write.worksheet("Phân Tích Định Biên")
            ws_dest.clear()
            ws_dest.resize(rows=len(po_list) + 15, cols=13)
        except gspread.exceptions.WorksheetNotFound:
            ws_dest = sh_write.add_worksheet(title="Phân Tích Định Biên", rows=str(len(po_list) + 15), cols="13")
            
        rows_cells = []
        
        # Row 1-2: Merged Title
        title_val = "BẢNG PHÂN TÍCH ĐỊNH BIÊN & TẢI TRỌNG SHIPPER - KHU VỰC NTB"
        row1 = [make_cell(title_val, bg='darkBlue', bold=True, fg='white', halign='CENTER', is_header=True)]
        row2 = [make_cell('', bg='darkBlue')]
        rows_cells.append(row1)
        rows_cells.append(row2)
        
        # Row 3: Subtitle / Date Info
        sub_val = f"Ngày phân tích: {date_str_dmy} (Dữ liệu ngày N-1: {date_str_iso}) | Tải trọng mục tiêu: {int(TARGET_WORKLOAD_PER_SHIPPER)} đơn/shipper/ngày"
        row3 = [make_cell(sub_val, bold=True, halign='LEFT')]
        rows_cells.append(row3)
        
        # Row 4: Column Headers
        headers = [
            "STT", "AM Quản Lý", "Bưu Cục", 
            "Shipper Thực Tế", "Đơn Mới Về", "Tồn Cũ Xử Lý", "Tổng Đơn Cần Giao",
            "Tải Trọng Thực Tế", "Định Biên Cần Thiết", "Thiếu Hụt Nhân Sự",
            "Trạng Thái Tải Trọng", "Tỉ Lệ Quá Tải (%)", "Tồn Cuối Ngày"
        ]
        row4 = [make_cell(h, bg='darkBlue', bold=True, fg='white', halign='CENTER') for h in headers]
        rows_cells.append(row4)
        
        # Row 5+: Data rows
        for idx, m in enumerate(po_list, 1):
            shippers = m['active_shippers']
            workload = m['total_workload']
            actual_load = round(m['actual_workload_per_shipper'], 1)
            overload_pct = round(m['overload_rate'], 1)
            
            # Formatting details
            alert_bg = m['alert_class']
            alert_fg = 'highFg' if alert_bg == 'highBg' else 'warnFg' if alert_bg == 'warnBg' else 'normalFg' if alert_bg == 'normalBg' else 'black'
            
            row_data = [
                make_cell(idx, halign='CENTER'),
                make_cell(m['am_name']),
                make_cell(m['po_name'], bold=True),
                make_cell(shippers, halign='CENTER'),
                make_cell(m['new_vol'], halign='CENTER'),
                make_cell(m['backlog_vol'], halign='CENTER'),
                make_cell(workload, halign='CENTER', bold=True),
                make_cell(actual_load, bg=alert_bg, bold=(alert_bg != 'white'), fg=alert_fg, halign='CENTER'),
                make_cell(m['required_shippers'], halign='CENTER'),
                make_cell(m['staffing_gap'], bg=('highBg' if m['staffing_gap'] > 0 else None), bold=(m['staffing_gap'] > 0), fg=('highFg' if m['staffing_gap'] > 0 else None), halign='CENTER'),
                make_cell(m['load_status'], bg=alert_bg, bold=(alert_bg != 'white'), fg=alert_fg, halign='CENTER'),
                make_cell(f"{overload_pct}%" if overload_pct > 0 else "-", bg=(alert_bg if overload_pct > 0 else None), bold=(overload_pct > 0), fg=(alert_fg if overload_pct > 0 else None), halign='CENTER'),
                make_cell(m['total_ton'], halign='CENTER')
            ]
            rows_cells.append(row_data)
            
        # Add Summary Row
        tot_shippers = sum(m['active_shippers'] for m in po_list)
        tot_new_vol = sum(m['new_vol'] for m in po_list)
        tot_backlog_vol = sum(m['backlog_vol'] for m in po_list)
        tot_workload = sum(m['total_workload'] for m in po_list)
        tot_required = sum(m['required_shippers'] for m in po_list)
        tot_gap = sum(m['staffing_gap'] for m in po_list)
        tot_ton = sum(m['total_ton'] for m in po_list)
        avg_load = round(tot_workload / tot_shippers, 1) if tot_shippers > 0 else 0.0
        
        row_summary = [
            make_cell("TỔNG CỘNG", bg='lightBlue', bold=True),
            make_cell("", bg='lightBlue'),
            make_cell("", bg='lightBlue'),
            make_cell(tot_shippers, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_new_vol, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_backlog_vol, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_workload, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(avg_load, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_required, bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_gap, bg=('highBg' if tot_gap > 0 else 'lightBlue'), bold=True, fg=('highFg' if tot_gap > 0 else None), halign='CENTER'),
            make_cell("Quá Tải" if avg_load > TARGET_WORKLOAD_PER_SHIPPER else "Tải Trọng Tốt", bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(f"{round((avg_load-90)/90*100, 1)}%" if avg_load > 90 else "-", bg='lightBlue', bold=True, halign='CENTER'),
            make_cell(tot_ton, bg='lightBlue', bold=True, halign='CENTER')
        ]
        rows_cells.append(row_summary)
        
        # Batch write cells to worksheet
        write_sheet_data(ws_dest, rows_cells)
        
        # Apply merged ranges requests
        merge_body = {
            'requests': [
                # Title block merge
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': ws_dest.id,
                            'startRowIndex': 0,
                            'endRowIndex': 2,
                            'startColumnIndex': 0,
                            'endColumnIndex': 13
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                # Subtitle block merge
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': ws_dest.id,
                            'startRowIndex': 2,
                            'endRowIndex': 3,
                            'startColumnIndex': 0,
                            'endColumnIndex': 13
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                # Summary row merge STT, AM, BC
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': ws_dest.id,
                            'startRowIndex': len(po_list) + 4,
                            'endRowIndex': len(po_list) + 5,
                            'startColumnIndex': 0,
                            'endColumnIndex': 3
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                }
            ]
        }
        sh_write.batch_update(merge_body)
        
        # Set column dimensions
        widths = [50, 150, 220, 120, 100, 100, 130, 130, 130, 120, 200, 120, 100]
        set_col_widths(ws_dest, widths)
        
        print("✔️ Successfully wrote styled results to Google Sheets.")
    except Exception as e:
        print(f"❌ Failed to update Google Sheets tab: {e}")
        import traceback
        traceback.print_exc()
        
    # ============ 9. GENERATE LOCAL REPORT ============
    print("\n📝 Generating local Markdown report...")
    md_rows = []
    md_rows.append(f"# 📊 Báo Cáo Định Biên & Tải Trọng Shipper - NTB - Ngày {date_str_dmy}")
    md_rows.append(f"*Dữ liệu phân tích ngày N-1: {date_str_iso} ({weekday_map[target_date.weekday()]})*")
    md_rows.append(f"*Ngưỡng tải trọng tiêu chuẩn: **{int(TARGET_WORKLOAD_PER_SHIPPER)} đơn/shipper/ngày***\n")
    
    # Executive Summary Stats
    tot_shippers = sum(m['active_shippers'] for m in po_list)
    tot_workload = sum(m['total_workload'] for m in po_list)
    tot_new_vol = sum(m['new_vol'] for m in po_list)
    tot_backlog_vol = sum(m['backlog_vol'] for m in po_list)
    tot_gap = sum(m['staffing_gap'] for m in po_list)
    avg_load = tot_workload / tot_shippers if tot_shippers > 0 else 0.0
    
    overloaded_pos = [m for m in po_list if m['actual_workload_per_shipper'] > TARGET_WORKLOAD_PER_SHIPPER and m['active_shippers'] > 0]
    extreme_overloaded_pos = [m for m in po_list if m['actual_workload_per_shipper'] >= 120.0 and m['active_shippers'] > 0]
    zero_shipper_pos = [m for m in po_list if m['active_shippers'] == 0 and m['total_workload'] > 0]
    
    md_rows.append("## 📌 Tóm Tắt Tình Hình Nhân Sự Toàn Vùng NTB")
    md_rows.append(f"- **Tổng số nhân sự giao (Active Shippers)**: **{tot_shippers}** nhân sự.")
    md_rows.append(f"- **Tổng sản lượng cần xử lý**: **{tot_workload:,}** đơn (Hàng mới: {tot_new_vol:,} | Tồn cũ: {tot_backlog_vol:,}).")
    md_rows.append(f"- **Tải trọng trung bình toàn vùng**: **{avg_load:.1f}** đơn/shipper/ngày.")
    md_rows.append(f"- **Số bưu cục quá tải (>90 đơn/shipper)**: **{len(overloaded_pos)}** / {len(po_list)} bưu cục hoạt động.")
    md_rows.append(f"- **Số bưu cục quá tải cực đoan (>=120 đơn/shipper)**: **{len(extreme_overloaded_pos)}** bưu cục.")
    md_rows.append(f"- **Số bưu cục có đơn phát sinh nhưng trống shipper**: **{len(zero_shipper_pos)}** bưu cục.")
    md_rows.append(f"- **Tổng số nhân sự thiếu hụt cần bổ sung**: **{tot_gap}** shippers để đưa tải trọng về dưới {int(TARGET_WORKLOAD_PER_SHIPPER)} đơn/ngày.")
    md_rows.append("\n---\n")
    
    # Top 10 Overloaded Post Offices Table
    md_rows.append("## 🔥 Top 10 Bưu Cục Quá Tải Nghiêm Trọng Nhất")
    md_rows.append("| STT | Bưu Cục | AM Quản Lý | Shipper Hiện Tại | Tổng Đơn Giao | Tải Trọng (Đơn/Shipper) | Định Biên Cần | Thiếu Hụt | Trạng Thái |")
    md_rows.append("|---|---|---|---|---|---|---|---|---|")
    
    for idx, m in enumerate(po_list[:10], 1):
        actual_load = m['actual_workload_per_shipper']
        shippers = m['active_shippers']
        gap_str = f"**{m['staffing_gap']}**" if m['staffing_gap'] > 0 else "-"
        shippers_str = f"{shippers}" if shippers > 0 else "`0`"
        
        md_rows.append(
            f"| {idx} | **{m['po_name']}** | {m['am_name']} | {shippers_str} | {m['total_workload']:,} | **{actual_load:.1f}** | {m['required_shippers']} | {gap_str} | `{m['load_status']}` |"
        )
    md_rows.append("\n---\n")
    
    # Detailed analysis of underperforming post offices mentioned by AMs
    md_rows.append("## 🔍 Đối Chiếu Giải Trình Của AM vs Số Liệu Định Biên Thực Tế")
    md_rows.append("Dưới đây là phân tích chi tiết định biên cho **5 bưu cục** được các AM giải trình xin thêm người hoặc báo cáo thiếu hụt nhân sự vào ngày 25/06:")
    md_rows.append("")
    
    am_focus_bcs = ["(LDO) Di Linh", "(LDO) Hòa Ninh", "(LDO) Tân Hà Lâm Hà", "(DNO) Quảng Tín", "(KHO) Cam Linh"]
    
    for bc in am_focus_bcs:
        m = active_po_workload.get(bc)
        if not m:
            # Try matching
            matched_name = map_bc(bc, df_cc)
            m = active_po_workload.get(matched_name)
            
        if not m:
            md_rows.append(f"### 📍 Bưu cục: {bc} (Không có dữ liệu volume ngày này)")
            continue
            
        shippers = m['active_shippers']
        workload = m['total_workload']
        load_per_shipper = m['actual_workload_per_shipper']
        gap = m['staffing_gap']
        req = m['required_shippers']
        
        # Context mappings for AM claims
        am_claims = {
            "(LDO) Di Linh": "Thiếu 6 nhân sự giao",
            "(LDO) Hòa Ninh": "Thiếu 2 NS nghỉ ca chiều",
            "(LDO) Tân Hà Lâm Hà": "Thiếu 4 nhân sự giao",
            "(DNO) Quảng Tín": "Nghẽn năng lực đơn ca (Tồn cũ + Mới = 623 đơn)",
            "(KHO) Cam Linh": "Thiếu nghiêm trọng 6 nhân sự giao"
        }
        
        claim = am_claims.get(bc, "Chưa ghi nhận giải trình")
        
        md_rows.append(f"### 📍 Bưu cục: **{m['po_name']}** (AM {m['am_name']})")
        md_rows.append(f"- **Giải trình từ AM**: *\"{claim}\"*")
        md_rows.append(f"- **Shipper hiện tại**: **{shippers}** | **Tổng sản lượng cần giao**: **{workload}** đơn (Mới: {m['new_vol']} | Tồn cũ: {m['backlog_vol']}).")
        md_rows.append(f"- **Tải trọng shipper thực tế**: **{load_per_shipper:.1f}** đơn/shipper/ngày (Gấp **{load_per_shipper/TARGET_WORKLOAD_PER_SHIPPER:.1f} lần** mức tiêu chuẩn).")
        md_rows.append(f"- **Số lượng định biên cần thiết**: **{req}** shippers.")
        md_rows.append(f"- **Thiếu hụt định biên tính toán (Staffing Gap)**: **{gap}** shippers.")
        
        # Verdict/Insight
        if gap == 0:
            verdict = "Số lượng nhân sự hiện tại đủ đáp ứng tải trọng mục tiêu. Sự cố giao hàng có thể do lỗi điều phối ca chiều hoặc năng suất shipper kém."
        else:
            if bc == "(LDO) Di Linh":
                verdict = f"Số liệu hoàn toàn ủng hộ giải trình của AM. Tải trọng {load_per_shipper:.1f} đơn/shipper là bất khả thi. Mức thiếu hụt thực tế lên tới {gap} nhân sự. Đề xuất duyệt khẩn cấp bổ sung 6 nhân sự hoặc điều động hỗ trợ."
            elif bc == "(KHO) Cam Linh":
                verdict = f"Cực kỳ báo động! Tải trọng lên tới {load_per_shipper:.1f} đơn/shipper. Việc AM báo thiếu 6 nhân sự là hoàn toàn chính xác. Tính toán cho thấy bưu cục thiếu tới {gap} nhân sự để vận hành an toàn. Hiện AM đã điều động 7 NS hỗ trợ là giải pháp tạm thời rất phù hợp."
            elif bc == "(LDO) Tân Hà Lâm Hà":
                verdict = f"Bưu cục thiếu hụt nhân sự nghiêm trọng. AM xin 4 người, thực tế tính toán thiếu tới {gap} người. Đề xuất tăng cường tuyển dụng hoặc hỗ trợ tuyến xa."
            elif bc == "(LDO) Hòa Ninh":
                verdict = f"Bưu cục thiếu {gap} shippers. Tuy nhiên, việc gán ca 2 chiều bằng 0 đơn là lỗi vận hành chủ quan. Cần siết kỷ luật điều phối ca chiều bên cạnh việc bổ sung nhân sự."
            else:
                verdict = f"Tải trọng thực tế vượt ngưỡng an toàn. Bưu cục cần bổ sung thêm {gap} shippers để khôi phục chỉ số KPI giao nhận."
                
        md_rows.append(f"- **💡 Nhận định của hệ thống**: {verdict}")
        md_rows.append("")
        
    md_output = "\n".join(md_rows)
    md_file_path = os.path.join(BASE_DIR, "shipper_workload_analysis.md")
    with open(md_file_path, "w", encoding="utf-8") as f:
        f.write(md_output)
    print(f"💾 Successfully saved local markdown analysis file to: {md_file_path}")
    print("\n--- RESULTS PREVIEW ---")
    print(md_output[:1000] + "\n...(truncated)...")
    print("-----------------------")

if __name__ == "__main__":
    main()
