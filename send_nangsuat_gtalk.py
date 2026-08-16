# -*- coding: utf-8 -*-
"""
Gửi ảnh NĂNG SUẤT NVPTT ĐẦY ĐỦ (không lọc mức thấp) theo từng AM, dữ liệu ngày N-1,
đọc từ tab 'BaoCao' (tab này luôn khóa số liệu ngày hôm qua).
Mỗi AM 1 ảnh / mỗi bưu cục 1 ảnh, gửi vào 1 group GTalk chung (theo token/ID mới).
"""
import os
import re
import sys
import json
import time
import unicodedata
from datetime import datetime, timedelta

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding='utf-8')

def get_http_session():
    session = requests.Session()
    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[500, 502, 503, 504, 429],
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

# ===== CẤU HÌNH GOOGLE SHEET =====
BAOCAO_SHEET_KEY = "1IUWdxN-VEC64OcciE09I_-3DHaTu39XhITblqMKt6Ww"
BAOCAO_TAB_NAME = "BaoCao"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def get_credentials_path():
    candidates = [
        os.path.join(BASE_DIR, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json"
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def parse_pct(val):
    """'71.62%' hoặc '0.7162' -> 71.62 ; rỗng/lỗi -> 0.0"""
    if val is None:
        return 0.0
    s = str(val).strip().replace("%", "").replace(",", ".")
    if s == "":
        return 0.0
    try:
        f = float(s)
    except ValueError:
        return 0.0
    # Nếu là dạng số thập phân kiểu 0.xx (tỷ lệ) thì nhân 100
    return f * 100 if 0 <= f <= 1.5 else f


def parse_num(val):
    if val is None or str(val).strip() == "":
        return 0
    s = str(val).strip().replace(".", "").replace(",", "")
    try:
        return int(float(s))
    except ValueError:
        return 0


SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    'credentials.json'
]


def get_gspread_client(spreadsheet_id=BAOCAO_SHEET_KEY):
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...", flush=True)

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
                creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


# ===== ĐỌC TAB BAOCAO (đầy đủ, không lọc mức thấp) =====

def read_baocao_full(sheet_key=BAOCAO_SHEET_KEY, tab_name=BAOCAO_TAB_NAME):
    gc = get_gspread_client(sheet_key)
    sh = gc.open_by_key(sheet_key)
    ws = sh.worksheet(tab_name)

    all_values = ws.get_all_values()
    print(f"🔎 [DEBUG] Tab '{tab_name}' có {len(all_values)} dòng x tối đa {max((len(r) for r in all_values), default=0)} cột.", flush=True)

    def norm(s):
        return unicodedata.normalize('NFC', str(s)).strip().lower()

    # Tìm dòng ngày báo cáo (VD: "Ngày" | 25/07/2026 ...)
    report_date_str = None
    for row in all_values[:5]:
        if row and norm(row[0]).startswith("ngày"):
            for cell in row[1:]:
                if cell.strip():
                    report_date_str = cell.strip()
                    break
            if report_date_str:
                break

    # Tìm dòng header (so khớp sau khi chuẩn hoá Unicode NFC để tránh lệch dấu tiếng Việt)
    header_idx = None
    for idx, row in enumerate(all_values):
        row_join = " ".join(norm(c) for c in row)
        if "nhan vien" in row_join.replace("â", "a").replace("nhân viên", "nhan vien") or "nhân viên" in row_join:
            if "bưu" in row_join or "buu" in row_join:
                header_idx = idx
                break

    if header_idx is None:
        # In ra 10 dòng đầu để dễ debug nếu vẫn không tìm thấy
        print("❌ [DEBUG] Không khớp được dòng header. 10 dòng đầu tiên của sheet:", flush=True)
        for idx, row in enumerate(all_values[:10]):
            print(f"   dòng {idx}: {row}", flush=True)
        raise ValueError("Không tìm thấy dòng header trong tab BaoCao (cần cột 'Bưu Cục' và 'Nhân Viên').")

    headers = [h.strip() for h in all_values[header_idx]]
    data_rows = all_values[header_idx + 1:]
    print(f"🔎 [DEBUG] Header tìm thấy ở dòng {header_idx}: {headers}", flush=True)
    print(f"🔎 [DEBUG] Số dòng dữ liệu phía sau header: {len(data_rows)}", flush=True)

    def col_idx(*candidates):
        for cand in candidates:
            for i, h in enumerate(headers):
                if norm(h).replace(" ", "") == norm(cand).replace(" ", ""):
                    return i
        return None

    i_manv = col_idx("Mã NV", "MaNV")
    i_bc = col_idx("Bưu Cục", "BuuCuc")
    i_am = col_idx("AM")
    i_name = col_idx("Nhân Viên", "NhanVien")
    i_gan = col_idx("Tổng đơn gán giao", "Gán Giao", "GanGiao")
    i_tc = col_idx("Số đơn GTC", "Giao TC", "GiaoTC")
    i_pct = col_idx("%GTC")
    i_xephang = col_idx("Xếp hạng", "XepHang")
    i_note = col_idx("Note")
    i_danhgia = col_idx("Đánh Giá", "DanhGia")

    print(f"🔎 [DEBUG] Cột map -> Bưu Cục:{i_bc}  AM:{i_am}  Nhân Viên:{i_name}  Gán:{i_gan}  "
          f"TC:{i_tc}  %GTC:{i_pct}  Xếp hạng:{i_xephang}  Note:{i_note}  Đánh giá:{i_danhgia}", flush=True)

    missing = [name for name, i in [
        ("Bưu Cục", i_bc), ("AM", i_am), ("Nhân Viên", i_name), ("%GTC", i_pct)
    ] if i is None]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc trong tab BaoCao: {missing}. Header đọc được: {headers}")

    grouped = {}
    skipped_examples = []
    total_scanned = 0
    for row in data_rows:
        if not any(c.strip() for c in row):
            continue  # dòng trống hoàn toàn

        total_scanned += 1
        bc = row[i_bc].strip() if i_bc < len(row) else ""
        am = row[i_am].strip() if i_am < len(row) else ""
        name = row[i_name].strip() if i_name < len(row) else ""
        if not bc or not am or not name:
            if len(skipped_examples) < 5:
                skipped_examples.append(row)
            continue

        gan = parse_num(row[i_gan]) if (i_gan is not None and i_gan < len(row)) else 0
        tc = parse_num(row[i_tc]) if (i_tc is not None and i_tc < len(row)) else 0
        pct = parse_pct(row[i_pct]) if i_pct < len(row) else 0.0
        xep_hang = row[i_xephang].strip() if (i_xephang is not None and i_xephang < len(row)) else ""
        note = row[i_note].strip() if (i_note is not None and i_note < len(row)) else ""
        danh_gia = row[i_danhgia].strip() if (i_danhgia is not None and i_danhgia < len(row)) else ""

        grouped.setdefault(am, {}).setdefault(bc, []).append(
            (name, gan, tc, pct, xep_hang, note, danh_gia)
        )

    result = []
    for am_name, bc_map in grouped.items():
        bcs = [(bc_name, staff_list) for bc_name, staff_list in bc_map.items()]
        result.append({"am": am_name, "bcs": bcs})

    total_staff = sum(len(s) for am in result for _, s in am["bcs"])
    print(f"🔎 [DEBUG] Tổng số dòng có dữ liệu đã quét: {total_scanned}, bị bỏ qua (thiếu Bưu Cục/AM/Nhân Viên): "
          f"{total_scanned - total_staff}", flush=True)
    if total_staff == 0 and skipped_examples:
        print("❌ [DEBUG] Ví dụ vài dòng bị bỏ qua (raw row) để xem giá trị thực tế:", flush=True)
        for r in skipped_examples:
            print(f"   {r}", flush=True)
    print(f"📊 Đã đọc {total_staff} NVPTT, thuộc {len(result)} AM.", flush=True)
    return result, report_date_str


# ===== CẤU HÌNH GTALK CHÍNH THỨC =====
GTALK_OA_TOKEN = os.environ.get("NVPTT_GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"

AM_CHANNEL_MAP = {
    "Nguyễn Ngọc Khánh": "2077277510775197696",
    "Nguyễn Duy Long": "2077277718988836864",
    "Lê Thanh Nhựt": "2077277754418147328",
    "Trần Văn Phước": "2077277797832024064",
    "Trần Thị Nhung": "2077277827057745920",
    "Huỳnh Thị Kim Chi": "2077277857186131968",
    "Phan Đình Duy": "2077278383281876992",
    "Phạm Bá Thành Công": "2077277907735883776",
    "Thái Thị Thanh Thư": "2077277934947827712",
    "Nguyễn Thanh Long": "2077277974325506048",
    "Nguyễn Hoàng Phi": "2077278021170323456",
    "Trầm Hữu Tiến": "2077278046487142400",
    "Nguyễn Lê Nguyên Vũ": "2077278095459835904",
    "Lê Văn Trường": "2077278127814696960",
    "Hồng Bích Nga": "2077278157729837056",
    "Lê Minh Đại": "2077278182818799616",
    "Phan Thị Ngọc Diễm": "2079827073949868032",
    "Lê Hồng Minh Tâm": "2079827054540226560",
    "Cao Thị Thanh Thủy":	"2083241927281995776",
}


def get_channel_for_am(am_name):
    norm_target = unicodedata.normalize('NFC', am_name.strip()).lower()
    for name, channel_id in AM_CHANNEL_MAP.items():
        if unicodedata.normalize('NFC', name.strip()).lower() == norm_target:
            return channel_id
    default_ch = "2077277510775197696"
    print(f"⚠️ Không tìm thấy AM '{am_name}' trong AM_CHANNEL_MAP — dùng channel mặc định.", flush=True)
    return default_ch


# ===== TẠO HTML & CHỤP ẢNH BẰNG PLAYWRIGHT =====

def build_report_html(am_name, bc_name, staff, report_date_str):
    total_gan = sum(s[1] for s in staff)
    total_tc = sum(s[2] for s in staff)
    total_pct = (total_tc / total_gan * 100) if total_gan else 0.0
    total_rows = len(staff)

    if total_pct >= 80:
        overall_badge_cls = "bg-emerald-600 text-white"
    elif total_pct >= 60:
        overall_badge_cls = "bg-amber-500 text-white"
    else:
        overall_badge_cls = "bg-rose-600 text-white"

    tot_tot = sum(1 for s in staff if "tốt" in str(s[4]).lower())
    tot_kha = sum(1 for s in staff if "khá" in str(s[4]).lower())
    tot_caithein = sum(1 for s in staff if "cải thiện" in str(s[4]).lower() or "kém" in str(s[4]).lower() or "yếu" in str(s[4]).lower())

    rows_html = ""
    for idx, (name, gan, tc, pct, xep_hang, note, danh_gia) in enumerate(staff, start=1):
        if "_" in name:
            code, real_name = name.split("_", 1)
            name_html = f'<span class="emp-code">{code}</span> <span class="emp-name">{real_name}</span>'
        else:
            name_html = f'<span class="emp-name">{name}</span>'

        if pct >= 80:
            pct_badge = f'<span class="badge badge-success">{pct:.2f}%</span>'
        elif pct >= 60:
            pct_badge = f'<span class="badge badge-warning">{pct:.2f}%</span>'
        else:
            pct_badge = f'<span class="badge badge-danger">{pct:.2f}%</span>'

        xh_str = str(xep_hang)
        xh_lower = xh_str.lower()
        if "tốt" in xh_lower:
            xh_badge = f'<span class="badge badge-subtle-success">✨ {xh_str}</span>'
        elif "khá" in xh_lower:
            xh_badge = f'<span class="badge badge-subtle-blue">🔹 {xh_str}</span>'
        elif "cải thiện" in xh_lower or "yếu" in xh_lower or "kém" in xh_lower:
            xh_badge = f'<span class="badge badge-subtle-danger">⚠️ {xh_str}</span>'
        elif xh_str:
            xh_badge = f'<span class="badge badge-gray">{xh_str}</span>'
        else:
            xh_badge = '-'

        note_str = str(note)
        note_lower = note_str.lower()
        if "hỗ trợ" in note_lower:
            note_badge = f'<span class="badge badge-purple">{note_str}</span>'
        elif note_str:
            note_badge = f'<span class="badge badge-gray">{note_str}</span>'
        else:
            note_badge = '-'

        dg_str = str(danh_gia)
        dg_lower = dg_str.lower()
        if dg_lower == "đạt":
            dg_badge = f'<span class="badge badge-success-solid">✓ Đạt</span>'
        elif "không" in dg_lower:
            dg_badge = f'<span class="badge badge-danger-solid">✕ Chưa đạt</span>'
        elif dg_str:
            dg_badge = f'<span class="badge badge-gray">{dg_str}</span>'
        else:
            dg_badge = '<span style="color:#cbd5e1;">-</span>'

        even_cls = "row-even" if idx % 2 == 0 else "row-odd"

        rows_html += f"""
        <tr class="{even_cls}">
            <td class="col-stt">{idx}</td>
            <td class="col-name">{name_html}</td>
            <td class="col-num col-gan">{gan:,}</td>
            <td class="col-num col-tc">{tc:,}</td>
            <td class="col-pct">{pct_badge}</td>
            <td class="col-xh">{xh_badge}</td>
            <td class="col-note">{note_badge}</td>
            <td class="col-dg">{dg_badge}</td>
        </tr>
        """

    eval_summary = f"{tot_tot} Tốt"
    if tot_kha > 0:
        eval_summary += f" / {tot_kha} Khá"
    if tot_caithein > 0:
        eval_summary += f" / {tot_caithein} Cần Cải Thiện"

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:ital,wght@0,400;0,500;0,600;0,700;0,800;1,400&family=Plus+Jakarta+Sans:wght@500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, sans-serif;
            -webkit-font-smoothing: antialiased;
        }}

        body {{
            background: #eef2f6;
            padding: 24px;
            display: inline-block;
            min-width: 1200px;
        }}

        .container {{
            background: #ffffff;
            border-radius: 20px;
            box-shadow: 0 16px 36px -10px rgba(15, 23, 42, 0.1), 0 0 0 1px rgba(15, 23, 42, 0.05);
            overflow: hidden;
            width: 1240px;
        }}

        .header {{
            background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #0284c7 100%);
            color: #ffffff;
            padding: 28px 36px 24px 36px;
            position: relative;
        }}

        .header-top {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 20px;
        }}

        .header-title-box {{
            flex: 1;
        }}

        .title-tag {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(255, 255, 255, 0.2);
            border: 1px solid rgba(255, 255, 255, 0.35);
            color: #ffffff;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 1px;
            text-transform: uppercase;
            padding: 5px 14px;
            border-radius: 99px;
            margin-bottom: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        }}

        .main-title {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 28px;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: #ffffff;
            margin-bottom: 8px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}

        .sub-info {{
            display: flex;
            align-items: center;
            gap: 12px;
            color: #e0f2fe;
            font-size: 14px;
            font-weight: 500;
        }}

        .info-pill {{
            background: rgba(255, 255, 255, 0.15);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.25);
            padding: 6px 16px;
            border-radius: 10px;
            color: #ffffff;
            font-size: 14px;
            font-weight: 600;
        }}
        .info-pill strong {{
            color: #fef08a;
            font-weight: 800;
        }}

        .date-badge {{
            background: rgba(255, 255, 255, 0.18);
            border: 1px solid rgba(255, 255, 255, 0.3);
            padding: 8px 18px;
            border-radius: 14px;
            color: #e0f2fe;
            font-size: 12px;
            font-weight: 700;
            text-align: right;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
        }}
        .date-badge strong {{
            color: #ffffff;
            display: block;
            font-size: 16px;
            font-weight: 800;
            margin-top: 2px;
        }}

        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 14px;
            margin-top: 6px;
        }}

        .kpi-card {{
            background: rgba(255, 255, 255, 0.18);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.3);
            border-radius: 14px;
            padding: 14px 18px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }}

        .kpi-card.highlight {{
            background: rgba(255, 255, 255, 0.28);
            border-color: rgba(255, 255, 255, 0.5);
        }}

        .kpi-label {{
            font-size: 12px;
            color: rgba(255, 255, 255, 0.88);
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }}

        .kpi-val {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 25px;
            font-weight: 800;
            color: #ffffff;
        }}

        .kpi-val.accent {{
            color: #fef08a;
            text-shadow: 0 1px 3px rgba(0,0,0,0.15);
        }}
        
        .kpi-sub {{
            font-size: 11px;
            color: rgba(255, 255, 255, 0.75);
            margin-top: 2px;
            font-weight: 500;
        }}

        .table-container {{
            padding: 20px 24px 24px 24px;
            background: #ffffff;
        }}

        table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }}

        th {{
            background: #f1f5f9;
            color: #1e293b;
            font-size: 12px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            padding: 13px 14px;
            border-bottom: 2px solid #cbd5e1;
            text-align: left;
        }}

        th.text-right {{ text-align: right; }}
        th.text-center {{ text-align: center; }}

        td {{
            padding: 13px 14px;
            font-size: 14px;
            border-bottom: 1px solid #e2e8f0;
            color: #0f172a;
            vertical-align: middle;
        }}

        tr.row-even {{ background-color: #ffffff; }}
        tr.row-odd {{ background-color: #f8fafc; }}

        .col-stt {{
            width: 44px;
            text-align: center;
            color: #64748b;
            font-size: 13px;
            font-weight: 700;
        }}

        .col-name {{
            font-weight: 600;
            color: #0f172a;
        }}

        .emp-code {{
            display: inline-block;
            background: #f1f5f9;
            color: #334155;
            font-family: monospace;
            font-size: 12px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
            border: 1px solid #cbd5e1;
            margin-right: 6px;
        }}

        .emp-name {{
            font-weight: 700;
            color: #0f172a;
        }}

        .col-num {{
            text-align: right;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 15px;
            font-weight: 700;
        }}

        .col-gan {{ color: #1e293b; }}
        .col-tc {{ color: #2563eb; }}

        .col-pct {{
            text-align: right;
            width: 110px;
        }}

        .col-xh {{
            text-align: center;
            width: 170px;
        }}

        .col-note {{
            text-align: center;
            width: 170px;
        }}

        .col-dg {{
            text-align: center;
            width: 120px;
        }}

        .badge {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 5px 12px;
            border-radius: 99px;
            font-size: 12px;
            font-weight: 700;
            white-space: nowrap;
        }}

        .badge-success {{
            background: #dcfce7;
            color: #15803d;
            border: 1px solid #86efac;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 13px;
        }}

        .badge-warning {{
            background: #fef3c7;
            color: #b45309;
            border: 1px solid #fde047;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 13px;
        }}

        .badge-danger {{
            background: #ffe4e6;
            color: #be123c;
            border: 1px solid #fda4af;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 13px;
        }}

        .badge-subtle-success {{
            background: #f0fdf4;
            color: #166534;
            border: 1px solid #bbf7d0;
        }}

        .badge-subtle-blue {{
            background: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #bfdbfe;
        }}

        .badge-subtle-danger {{
            background: #fff1f2;
            color: #9f1239;
            border: 1px solid #fecdd3;
        }}

        .badge-purple {{
            background: #f3e8ff;
            color: #6b21a8;
            border: 1px solid #d8b4fe;
        }}

        .badge-gray {{
            background: #f1f5f9;
            color: #475569;
            border: 1px solid #cbd5e1;
        }}

        .badge-success-solid {{
            background: #16a34a;
            color: #ffffff;
            box-shadow: 0 2px 4px rgba(22, 163, 74, 0.2);
        }}

        .badge-danger-solid {{
            background: #dc2626;
            color: #ffffff;
            box-shadow: 0 2px 4px rgba(220, 38, 38, 0.2);
        }}

        .summary-row {{
            background: #e2e8f0 !important;
            color: #0f172a;
        }}

        .summary-row td {{
            color: #0f172a;
            font-weight: 800;
            border-top: 2px solid #cbd5e1;
            padding: 14px;
        }}

        .summary-label {{
            font-size: 14px;
            font-weight: 800;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            color: #1e3a8a;
        }}

        .summary-val {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 16px;
            font-weight: 800;
            text-align: right;
            color: #1e40af;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-top">
                <div class="header-title-box">
                    <div class="title-tag">🚀 BÁO CÁO NĂNG SUẤT NVPTT</div>
                    <h1 class="main-title">BƯU CỤC: {bc_name}</h1>
                    <div class="sub-info">
                        <div class="info-pill">AM Quản lý: <strong>{am_name}</strong></div>
                        <div class="info-pill">Quy mô: <strong>{total_rows} NVPTT</strong></div>
                    </div>
                </div>
                <div class="date-badge">
                    <span>NGÀY BÁO CÁO</span>
                    <strong>{report_date_str}</strong>
                </div>
            </div>

            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">TỔNG ĐƠN GÁN</div>
                    <div class="kpi-val">{total_gan:,}</div>
                    <div class="kpi-sub">Đơn hàng đã gán giao</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">GIAO THÀNH CÔNG</div>
                    <div class="kpi-val accent">{total_tc:,}</div>
                    <div class="kpi-sub">Số đơn đã hoàn tất</div>
                </div>
                <div class="kpi-card highlight">
                    <div class="kpi-label">TỶ LỆ GTC NVPTT</div>
                    <div class="kpi-val accent">{total_pct:.2f}%</div>
                    <div class="kpi-sub">Tính trên tổng đơn gán</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">ĐÁNH GIÁ NĂNG SUẤT</div>
                    <div class="kpi-val" style="font-size: 18px; line-height: 1.4; margin-top: 2px;">
                        {eval_summary}
                    </div>
                    <div class="kpi-sub">Phân loại nhân sự</div>
                </div>
            </div>
        </div>

        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th class="text-center" style="width: 44px;">#</th>
                        <th>Nhân viên PTT</th>
                        <th class="text-right">Đơn Gán</th>
                        <th class="text-right">Đơn GTC</th>
                        <th class="text-right">% GTC</th>
                        <th class="text-center">Xếp Hạng</th>
                        <th class="text-center">Loại NV (Note)</th>
                        <th class="text-center">Đánh Giá</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr class="summary-row">
                        <td class="text-center" style="color: #64748b;">-</td>
                        <td class="summary-label">TỔNG CỘNG NVPTT</td>
                        <td class="summary-val">{total_gan:,}</td>
                        <td class="summary-val" style="color: #2563eb;">{total_tc:,}</td>
                        <td class="summary-val" style="text-align: right;">
                            <span class="badge {overall_badge_cls}" style="font-size: 14px; padding: 4px 12px; font-weight: 800;">{total_pct:.2f}%</span>
                        </td>
                        <td colspan="3"></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>"""
    return html


def generate_report_image(am_name, bc_name, staff, report_date_str, out_path, bc_idx=0):
    """Tạo 1 ảnh cho ĐÚNG 1 bưu cục của 1 AM bằng Playwright (với fallback Pillow)."""
    try:
        from playwright.sync_api import sync_playwright
        html_content = build_report_html(am_name, bc_name, staff, report_date_str)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(device_scale_factor=2)
            page.set_content(html_content)
            page.evaluate("document.fonts.ready")
            container = page.query_selector(".container")
            if container:
                container.screenshot(path=out_path)
            else:
                page.screenshot(path=out_path, full_page=True)
            browser.close()
        print(f"✅ Đã tạo ảnh thành công bằng Playwright: {out_path}", flush=True)
        return out_path
    except Exception as err:
        print(f"⚠️ Playwright render lỗi ({err}), chuyển sang dùng Pillow fallback...", flush=True)
        return _generate_report_image_pillow(am_name, bc_name, staff, report_date_str, out_path)


def _generate_report_image_pillow(am_name, bc_name, staff, report_date_str, out_path):
    """Fallback dùng PIL nếu không có Playwright."""
    font_bold_candidates = [r"C:\Windows\Fonts\arialbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
    font_reg_candidates = [r"C:\Windows\Fonts\arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]

    def _first(paths):
        for p in paths:
            if os.path.exists(p):
                return p
        return paths[0]

    f_b = _first(font_bold_candidates)
    f_r = _first(font_reg_candidates)

    def font(sz, b=False):
        return ImageFont.truetype(f_b if b else f_r, sz)

    f_h1 = font(28, True)
    f_sub = font(16)
    f_name = font(16, True)
    f_val = font(16, True)

    total_gan = sum(s[1] for s in staff)
    total_tc = sum(s[2] for s in staff)
    total_pct = (total_tc / total_gan * 100) if total_gan else 0.0
    total_rows = len(staff)

    row_h = 44
    H = 160 + 50 + total_rows * row_h + 40
    W = 1200

    img = Image.new("RGB", (W, H), (238, 242, 246))
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 140], fill=(37, 99, 235))
    d.text((30, 20), f"BÁO CÁO NĂNG SUẤT NVPTT - {bc_name}", font=f_h1, fill=(255, 255, 255))
    d.text((30, 65), f"AM: {am_name}  ·  Ngày: {report_date_str}  ·  {total_rows} NV", font=f_sub, fill=(254, 240, 138))
    d.text((30, 95), f"Tổng Gán: {total_gan:,}  ·  Tổng GTC: {total_tc:,}  ·  Tỷ lệ: {total_pct:.2f}%", font=f_sub, fill=(255, 255, 255))

    y = 150
    for idx, (name, gan, tc, pct, xep_hang, note, danh_gia) in enumerate(staff, 1):
        bg_col = (255, 255, 255) if idx % 2 == 1 else (248, 250, 252)
        d.rectangle([20, y, W - 20, y + row_h - 2], fill=bg_col)
        d.text((35, y + 10), f"{idx}. {name}", font=f_name, fill=(15, 23, 42))
        d.text((450, y + 10), f"Gán: {gan:,}", font=f_val, fill=(30, 41, 59))
        d.text((600, y + 10), f"GTC: {tc:,}", font=f_val, fill=(37, 99, 235))
        d.text((750, y + 10), f"{pct:.2f}%", font=f_val, fill=(22, 163, 74) if pct >= 80 else (225, 29, 72))
        d.text((900, y + 10), f"{xep_hang}", font=f_sub, fill=(71, 85, 105))
        y += row_h

    img.save(out_path)
    return out_path


# ===== GỬI ẢNH LÊN GTALK =====

def upload_image_to_gtalk(image_path, channel_id, oa_token, session=None):
    if session is None:
        session = get_http_session()

    file_name = os.path.basename(image_path)
    file_size = os.path.getsize(image_path)
    with open(image_path, "rb") as f:
        file_bytes = f.read()

    with Image.open(image_path) as im:
        width, height = im.size

    init_payload = {
        "ChannelId": channel_id,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": width, "height": height}),
        "oaToken": oa_token
    }

    presigned_url = None
    upload_id = None

    for attempt in range(1, 4):
        try:
            resp_init = session.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=30)
            if resp_init.status_code == 200:
                init_data = resp_init.json()
                if init_data.get("errorCode") == "success":
                    presigned_url = init_data["data"]["PresignedURL"]
                    upload_id = init_data["data"]["UploadId"]
                    break
                else:
                    print(f"⚠️ GTalk initiate-upload logic error: {init_data}", flush=True)
                    return None, None
            else:
                print(f"⚠️ GTalk initiate-upload HTTP error {resp_init.status_code}: {resp_init.text}", flush=True)
        except Exception as e:
            print(f"⚠️ Thử lần {attempt}/3 initiate-upload bị lỗi mạng: {e}", flush=True)
            time.sleep(attempt * 2)
    else:
        return None, None

    for attempt in range(1, 4):
        try:
            resp_put = session.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=60)
            if resp_put.status_code == 200:
                break
            else:
                print(f"⚠️ GTalk put-file HTTP error {resp_put.status_code}: {resp_put.text}", flush=True)
        except Exception as e:
            print(f"⚠️ Thử lần {attempt}/3 put-file bị lỗi mạng: {e}", flush=True)
            time.sleep(attempt * 2)
    else:
        return None, None

    for attempt in range(1, 4):
        try:
            resp_comp = session.post(
                "https://mbff.ghn.vn/api/gtalk/complete-upload",
                json={"oaToken": oa_token, "UploadId": upload_id},
                timeout=30
            )
            if resp_comp.status_code == 200:
                comp_data = resp_comp.json()
                if comp_data.get("errorCode") == "success":
                    return comp_data["data"]["Id"], (width, height)
                else:
                    print(f"⚠️ GTalk complete-upload logic error: {comp_data}", flush=True)
                    return None, None
            else:
                print(f"⚠️ GTalk complete-upload HTTP error {resp_comp.status_code}: {resp_comp.text}", flush=True)
        except Exception as e:
            print(f"⚠️ Thử lần {attempt}/3 complete-upload bị lỗi mạng: {e}", flush=True)
            time.sleep(attempt * 2)
    return None, None


def send_report_to_gtalk(image_path, caption, channel_id, oa_token=None, session=None):
    if session is None:
        session = get_http_session()

    oa_token = oa_token or GTALK_OA_TOKEN
    print("📡 Đang upload ảnh lên GTalk...", flush=True)
    file_id, size = upload_image_to_gtalk(image_path, channel_id, oa_token, session=session)
    if not file_id:
        print("❌ Upload ảnh lên GTalk thất bại.", flush=True)
        return False

    width, height = size
    send_payload = {
        "channelId": channel_id,
        "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": caption,
                "items": [
                    {"image": {"fileId": file_id, "width": width, "height": height}}
                ]
            }
        },
        "oaToken": oa_token
    }

    for attempt in range(1, 4):
        try:
            r_send = session.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload, timeout=30)
            if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                print("✅ Đã gửi ảnh vào GTalk thành công!", flush=True)
                return True
            else:
                print(f"❌ Gửi tin nhắn GTalk thất bại (Lần {attempt}): {r_send.text}", flush=True)
        except Exception as e:
            print(f"⚠️ Thử lần {attempt}/3 send-message bị lỗi mạng: {e}", flush=True)
            time.sleep(attempt * 2)

    return False


def main():
    yesterday = datetime.now() - timedelta(days=1)
    default_date_str = yesterday.strftime("%d/%m/%Y")

    try:
        data, sheet_date_str = read_baocao_full()
    except Exception as e:
        print(f"❌ Lỗi đọc dữ liệu từ Google Sheet: {e}", flush=True)
        return

    report_date_str = sheet_date_str or default_date_str

    if not data:
        print("⚠️ Không có dữ liệu NVPTT nào trong tab BaoCao.", flush=True)
        return

    # Gộp thành danh sách (am_name, bc_name, staff) — mỗi bưu cục 1 ảnh riêng
    jobs = []
    for am_block in data:
        am_name = am_block["am"]
        for bc_idx, (bc_name, staff) in enumerate(am_block["bcs"]):
            jobs.append((am_name, bc_name, staff, bc_idx))

    print(f"📦 Sẽ tạo và gửi {len(jobs)} ảnh (mỗi bưu cục 1 ảnh), ngày báo cáo: {report_date_str}", flush=True)

    session = get_http_session()

    for i, (am_name, bc_name, staff, bc_idx) in enumerate(jobs, start=1):
        total_nv = len(staff)

        safe_name = "".join(c if c.isalnum() else "_" for c in f"{am_name}_{bc_name}")
        out_path = os.path.join(BASE_DIR, f"nvptt_{safe_name}.png")

        print(f"\n[{i}/{len(jobs)}] 🖼️ Đang tạo ảnh cho AM: {am_name} · Bưu cục: {bc_name} ({total_nv} NVPTT)...", flush=True)
        generate_report_image(am_name, bc_name, staff, report_date_str, out_path, bc_idx=bc_idx)

        caption = (
            f"<b>BÁO CÁO NĂNG SUẤT NVPTT</b>\n"
            f"AM: <b>{am_name}</b> · Bưu cục: <b>{bc_name}</b> · {total_nv} nhân viên\n"
            f"Ngày báo cáo {report_date_str}"
        )

        channel_id = get_channel_for_am(am_name)
        try:
            ok = send_report_to_gtalk(out_path, caption, channel_id=channel_id, session=session)
            if not ok:
                print(f"⚠️ Gửi ảnh AM '{am_name}' · Bưu cục '{bc_name}' thất bại, tiếp tục...", flush=True)
        except Exception as err:
            print(f"⚠️ Lỗi ngoại lệ khi gửi ảnh AM '{am_name}' · Bưu cục '{bc_name}': {err}. Tiếp tục công việc tiếp theo...", flush=True)

        try:
            os.remove(out_path)
        except Exception:
            pass

        if i < len(jobs):
            time.sleep(2)

    print("\n🎉 Đã xử lý xong toàn bộ bưu cục.", flush=True)


if __name__ == "__main__":
    main()
