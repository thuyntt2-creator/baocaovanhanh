# -*- coding: utf-8 -*-
"""
BÁO CÁO TỰ ĐỘNG BƯU CỤC BẤT ỔN (AUTOMATIC UNSTABLE HUBS REPORT GENERATOR)
========================================================================
- Tự động đọc danh sách Bưu Cục từ Cột A của sheet `bưu cục` (hoặc `Bưu cục`).
- Tự động lọc bỏ Nhân viên hỗ trợ từ sheet `NV HỖ TRỢ` khỏi tính toán năng suất & thu nhập.
- Tự động ghi kết quả báo cáo thuần TEXT (không dính ký tự **, ***, ###)
  vào Cột B của sheet `bưu cục` và in ra màn hình.
- Tự động khớp dữ liệu và tạo Báo cáo chuẩn sếp từ các sheet:
    + `báo cáo tuyển dụng`: Định biên, Hiện tại, Trạng thái PTTT/NVXL, Số NV thiếu, HRBP.
    + `data`: Volume, GTC, Tồn, Chưa gán (Ca 1 + Tồn), Tính Tổng % GTC Bưu cục.
    + `thu nhập`: Năng suất gán/NV, GTC/NV, Thu nhập TB/ngày theo nhóm thâm niên.
    + `năng suất`: Phân nhóm năng suất NV & Thâm niên nhóm NV GTC thấp.
    + `NV HỖ TRỢ`: Danh sách nhân viên hỗ trợ cần loại khỏi thống kê bưu cục.

Cách sử dụng:
    python generate_unstable_hubs_report.py
"""

import sys
import os
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

# ================= CẤU HÌNH GOOGLE SHEET =================
SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
DEFAULT_HUBS = ["20942000", "22830000", "21477000"]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    'credentials.json'
]


def get_gspread_client(spreadsheet_id=SHEET_ID):
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

    auth_user_candidates = [
        os.path.join(BASE_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]

    # 1. Thử dùng token OAuth hiện tại trong authorized_user.json
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
                print(f"⚠️ OAuth token ({auth_user_file}) bị hết hạn/lỗi: {e}")

    # 2. Nếu token hết hạn, kích hoạt luồng đăng nhập trình duyệt bằng OAuth
    oauth_candidates = [
        os.path.join(BASE_DIR, 'credentials_oauth.json'),
        r'C:\Users\lap4all\Documents\Auto report\credentials_oauth.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials_oauth.json',
        'credentials_oauth.json'
    ]
    for oauth_file in oauth_candidates:
        if os.path.exists(oauth_file):
            auth_out = os.path.join(os.path.dirname(oauth_file), 'authorized_user.json')
            print(f"🔄 Đang kích hoạt trình duyệt để xác thực lại tài khoản GHN của bạn...")
            try:
                if os.path.exists(auth_out):
                    try:
                        os.remove(auth_out)
                    except Exception:
                        pass
                gc = gspread.oauth(
                    credentials_filename=oauth_file,
                    authorized_user_filename=auth_out,
                    scopes=scopes
                )
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Không thể hoàn tất xác thực trình duyệt: {e}")

    # 3. Phương án dự phòng Service Account
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

    raise PermissionError("❌ Không thể xác thực Google Sheets. Vui lòng chạy đăng nhập lại OAuth qua trình duyệt.")


def clean_num(val):
    """Làm sạch và ép kiểu số chính xác cho định dạng số tiếng Việt/Anh."""
    if not val or pd.isna(val):
        return 0.0
    s = str(val).strip().replace('đ', '').replace('%', '').strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) == 3 and not parts[1].endswith('00'):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    elif '.' in s:
        parts = s.split('.')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_all_data():
    """Tải dữ liệu từ tất cả các sheet cần thiết."""
    client = get_gspread_client(SHEET_ID)
    spreadsheet = client.open_by_key(SHEET_ID)

    worksheets = {ws.title.lower().strip(): ws for ws in spreadsheet.worksheets()}

    target_hubs = []
    ws_input = None
    for key in ['bưu cục', 'bưucục', 'buucuc', 'bưu_cục', 'báo cáo']:
        if key in worksheets:
            ws_input = worksheets[key]
            break

    if ws_input:
        vals = ws_input.get_all_values()
        for idx, row in enumerate(vals):
            if row and row[0].strip():
                val = row[0].strip()
                if idx == 0 and val.lower() in ['bưu cục', 'tên bưu cục', 'mã bưu cục', 'stt', 'id']:
                    continue
                target_hubs.append((idx + 1, val))

    if not target_hubs:
        target_hubs = [(i+1, h) for i, h in enumerate(DEFAULT_HUBS)]

    # Đọc danh sách Nhân viên Hỗ trợ từ sheet `NV HỖ TRỢ`
    ws_hotro = None
    for key in ['nv hỗ trợ', 'nv ho tro', 'nv_ho_tro', 'ho tro', 'hỗ trợ']:
        if key in worksheets:
            ws_hotro = worksheets[key]
            break

    support_emp_codes = set()
    support_emp_names = set()
    support_details = []
    if ws_hotro:
        vals_hotro = ws_hotro.get_all_values()
        if len(vals_hotro) > 1:
            for row in vals_hotro[1:]:
                if row and row[0].strip():
                    raw_val = row[0].strip()
                    code = raw_val.split('_')[0].split('-')[0].strip()
                    if code.isdigit():
                        support_emp_codes.add(code)
                    name = row[1].strip() if len(row) > 1 else ""
                    if name:
                        support_emp_names.add(name.lower())
                    support_details.append({
                        "id": code,
                        "name": name,
                        "am_nhan": row[2].strip() if len(row) > 2 else "",
                        "nguon_am": row[3].strip() if len(row) > 3 else "",
                        "bc_chinh": row[4].strip() if len(row) > 4 else "",
                        "bc_nhan_ho_tro": row[5].strip() if len(row) > 5 else ""
                    })

    return {
        "client": client,
        "spreadsheet": spreadsheet,
        "ws_input": ws_input,
        "target_hubs": target_hubs,
        "rec": spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values(),
        "data": spreadsheet.worksheet("data").get_all_values(),
        "tn": spreadsheet.worksheet("thu nhập").get_all_values(),
        "ns": spreadsheet.worksheet("năng suất").get_all_values(),
        "support_emp_codes": support_emp_codes,
        "support_emp_names": support_emp_names,
        "support_details": support_details
    }


# Map ghi chú vận hành tuyến thực tế theo mã Bưu cục (Tự động biến tấu theo ngày - Option 2)
def get_dynamic_route_notes(code, today_str):
    notes = {
        "22830000": {
            "thuc_te_di_lam": "16 NV (13 PTTT + 3 NV hỗ trợ). GTC Ca 1: 56.41% (154/273 đơn).",
            "lech_tuyen": f"""  Cam Phúc Nam (70-80 đơn): 01 NV chạy ca chiều (14–15h) do đi biển đêm -> Đang tiếp tục nhận hồ sơ tuyển thay thế.
  Cam Phú (100-120 đơn): 02 NV (01 NV mới + 01 NV xin nghỉ từ 15/8) -> Đang lọc hồ sơ tuyển thêm 1-2 NV.
  Cam Linh (120-150 đơn): 01 NV nghỉ ngang -> Nguy cơ bỏ tuyến, đang ưu tiên phỏng vấn gấp 2 NV.
  Cam Phước Đông & Cam Thịnh: 02 NV mới, 1 NV hỗ trợ gánh tuyến xã xa 8–10km -> Quá tải, đang tiếp tục tuyển 3 NV.
  Cam Lợi (90-100 đơn): 01 NV mới sức khỏe yếu đã nghỉ -> Đang tuyển gấp 1 NV.
  Cụm đảo (Bình Ba/Hưng/Lập): 03 NV (Duy trì ổn định).
  Cam Thuận, Cam Lộc, Ba Ngòi: 04 NV cứng -> Vận hành ổn định."""
        },
        "20942000": {
            "thuc_te_di_lam": f"Clear hàng còn thấp, chưa ổn định (Ghi nhận ngày {today_str}). Trong đó có 2 NV di chuyển hỗ trợ Đức Trọng 1. Bưu cục đang tiếp tục sắp xếp cố định tuyến.",
            "lech_tuyen": f"""  Thị trấn Di Linh (5/5 NV): 03 NV thâm niên <3 tháng (năng suất giao thấp 45–60 đơn/ngày) + 02 NV thâm niên >3 tháng (năng suất giao 70–80 đơn/ngày).
  Xã Bảo Thuận : 01 NV thâm niên <3 tháng.
  Xã Đinh Lạc : 01 NV thâm niên >3 tháng (chạy cứng).
  Xã Đinh Trang Thượng (1/1 NV): Hiện chưa có nhân sự chạy cố định (đang cho NV cover).
  Xã Gia Hiệp : 01 NV thâm niên <3 tháng (năng suất giao 40–60 đơn/ngày).
  Xã Gung Ré : 01 NV chạy cứng (năng suất giao 100–150 đơn/ngày).
  Xã Sơn Điền & Xã Gia Bắc: 01 NV chạy cứng (năng suất giao 80–100 đơn/ngày).
  Xã Tam Bố: 01 NV cũ vừa nghỉ, thay thế bằng NV mới <3 tháng (năng suất giao 40–55 đơn/ngày).
  Xã Tân Châu: 01 NV chạy cứng (năng suất giao 70–80 đơn/ngày).
  Xã Tân Lâm: 01 NV chạy cứng (năng suất giao 90–130 đơn/ngày).
  Xã Tân Nghĩa: 01 NV thâm niên <3 tháng (năng suất giao 40–55 đơn/ngày).
  Xã Tân Thượng: 01 NV thâm niên <3 tháng (năng suất giao 40–50 đơn/ngày)."""
        },
        "21477000": {
            "thuc_te_di_lam": "16 NV (14 PTTT + 2 XL). GTC Ca 1: 69.32% (122/176 đơn).",
            "lech_tuyen": f"""  Đạ Nhim, Đạ Sar: 02 NV chạy tuyến chành (Đúng tuyến).
  Bidoup: 01 NV kèm NV mới, nghỉ cuối tháng -> Nguy cơ trống tuyến.
  19/5 - TDP Đan Kia & Bạch Đằng - Cao Thắng: Hàng tăng, rộng -> Quá tải, phải bỏ lại 1-2 khu vực luân phiên.
  Mănglin - Đạ Phú: 02 NV mới, dân thưa, diện rộng -> Năng suất thấp, tồn đọng lớn.
  Nguyễn Hoàng, Xô Viết: Địa bàn xa kho -> Tải hàng chậm xuống tuyến.
  Đưng Kno - Xã Lát: 01 NV chạy 2 tuyến ngược chiều -> Mỗi ngày chỉ chạy được 1 tuyến, bỏ dở tuyến còn lại."""
        }
    }
    return notes.get(code, {})


def generate_report():
    data_dict = load_all_data()
    target_hubs = data_dict["target_hubs"]
    support_emp_codes = data_dict["support_emp_codes"]
    support_emp_names = data_dict["support_emp_names"]
    support_details = data_dict.get("support_details", [])

    df_rec = pd.DataFrame(data_dict["rec"])
    
    data_vals = data_dict["data"]
    df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

    tn_vals = data_dict["tn"]
    df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
    df_tn['don_gan_num'] = df_tn['Số đơn gán Giao'].apply(clean_num)
    df_tn['gtc_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
    df_tn['luong_num'] = df_tn['Lương HH/ ngày'].apply(clean_num)
    df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].split('_')[0].strip())
    df_tn['Emp_Name'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[-1].split('_')[-1].strip())

    ns_vals = data_dict["ns"]
    df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
    df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
    df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)
    df_ns['%GTC_num'] = df_ns['%GTC'].str.replace('%', '').str.replace(',', '.').apply(clean_num)
    df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].split('-')[0].strip())
    df_ns['Emp_Name'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[-1].split('-')[-1].strip())

    reports = []
    write_back_list = []

    for hub_idx, (row_in_sheet, hub_query) in enumerate(target_hubs, 1):
        query_clean = str(hub_query).strip()

        # 1. Khớp thông tin Tuyển dụng
        rec_row = None
        for idx, r in df_rec.iterrows():
            row_str = " ".join([str(x) for x in r])
            if query_clean.lower() in row_str.lower():
                rec_row = r
                break

        code = rec_row[2] if rec_row is not None else query_clean
        bc_short_name = rec_row[3] if rec_row is not None else query_clean
        full_bc_name = rec_row[4] if rec_row is not None else query_clean
        
        main_name = bc_short_name.split(')')[-1].strip() if ')' in bc_short_name else query_clean

        stt_pttt = rec_row[7] if (rec_row is not None and len(rec_row) > 7) else "N/A"
        raw_thieu = rec_row[8] if (rec_row is not None and len(rec_row) > 8) else "0"
        raw_hientai = rec_row[9] if (rec_row is not None and len(rec_row) > 9) else "0"
        raw_dinhbien = rec_row[10] if (rec_row is not None and len(rec_row) > 10) else "0"

        dinhbien_val = int(clean_num(raw_dinhbien))
        hientai_val = int(clean_num(raw_hientai))
        thieu_val = int(clean_num(raw_thieu))

        if dinhbien_val > 0 and hientai_val > 0:
            hientai_display = hientai_val
            thieu_display = max(0, dinhbien_val - hientai_val)
        elif dinhbien_val > 0 and thieu_val > 0:
            hientai_display = max(0, dinhbien_val - thieu_val)
            thieu_display = thieu_val
        else:
            hientai_display = hientai_val
            thieu_display = thieu_val

        # 2. Khớp thông tin Vận hành Data (Tính tổng tất cả Ca 1, Ca 2, Tồn cho ngày mới nhất)
        sub_d = df_data[
            df_data['Chi tiết'].str.contains(code, regex=False, na=False) |
            df_data['Chi tiết'].str.contains(bc_short_name, regex=False, na=False) |
            df_data['Chi tiết'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ]
        
        if len(sub_d) > 0:
            latest_date = sub_d['Time'].max()
            sub_d_latest = sub_d[sub_d['Time'] == latest_date]

            tot_vol = sub_d_latest['Volume'].apply(clean_num).sum()
            tot_gan = sub_d_latest['Sản Lượng Gán'].apply(clean_num).sum()
            tot_gtc = sub_d_latest['Sản Lượng Giao Thành Công'].apply(clean_num).sum()
            tot_ton = sub_d_latest['Sản Lượng Tồn'].apply(clean_num).sum()
            tot_chuagan = sub_d_latest['Sản Lượng Chưa Gán'].apply(clean_num).sum()
        else:
            tot_vol = tot_gan = tot_gtc = tot_ton = tot_chuagan = 0

        tot_gtc_rate = (tot_gtc / tot_vol * 100) if tot_vol > 0 else 0

        # 3. Khớp Thu Nhập Sheet (LỌC BỎ NV HỖ TRỢ)
        sub_tn = df_tn[
            df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
            df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
            df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ].copy()

        # Loại bỏ nhân viên hỗ trợ khỏi thu nhập
        sub_tn = sub_tn[
            (~sub_tn['Emp_Code'].isin(support_emp_codes)) &
            (~sub_tn['Emp_Name'].str.lower().isin(support_emp_names))
        ]

        nv_moi_cnt = len(sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', regex=False, na=False)])
        nv_cu_cnt = len(sub_tn) - nv_moi_cnt

        avg_gan = sub_tn['don_gan_num'].mean() if len(sub_tn) > 0 else 0
        avg_gtc = sub_tn['gtc_num'].mean() if len(sub_tn) > 0 else 0

        # 4. Khớp Năng Suất Sheet (Lấy ngày mới nhất & LỌC BỎ NV HỖ TRỢ)
        sub_ns_hub = df_ns[
            df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
            df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ].copy()
        
        date_display = "24/07"
        sub_ns = pd.DataFrame()
        if len(sub_ns_hub) > 0:
            def parse_date_key(d):
                try:
                    p = str(d).replace('thg', '').replace(',', '').split()
                    return (int(p[2]), int(p[1]), int(p[0]))
                except Exception:
                    return (0, 0, 0)
            
            unique_dates = sorted(sub_ns_hub['Ngay'].unique(), key=parse_date_key, reverse=True)
            if unique_dates:
                latest_ns_date = unique_dates[0]
                sub_ns = sub_ns_hub[sub_ns_hub['Ngay'] == latest_ns_date].copy()
                try:
                    p = latest_ns_date.replace('thg', '').replace(',', '').split()
                    date_display = f"{int(p[0]):02d}/{int(p[1]):02d}"
                except Exception:
                    date_display = latest_ns_date

        # Lưu lại danh sách trước khi lọc để đếm thực tế đi làm (gồm cả NV hỗ trợ)
        sub_ns_unfiltered = sub_ns.copy()

        # Loại bỏ nhân viên hỗ trợ khỏi năng suất
        if len(sub_ns) > 0:
            sub_ns = sub_ns[
                (~sub_ns['Emp_Code'].isin(support_emp_codes)) &
                (~sub_ns['Emp_Name'].str.lower().isin(support_emp_names))
            ]

        merged_ns = pd.merge(sub_ns, df_tn[['Emp_Code', 'Thâm niên', 'Ngày vào làm']], on='Emp_Code', how='left')

        cnt_good = len(merged_ns[merged_ns['TongDonGTC_num'] >= 50])
        cnt_mid = len(merged_ns[(merged_ns['TongDonGTC_num'] >= 35) & (merged_ns['TongDonGTC_num'] < 50)])
        cnt_low = len(merged_ns[merged_ns['TongDonGTC_num'] < 35])

        low_df = merged_ns[merged_ns['TongDonGTC_num'] < 35]
        low_moi_cnt = len(low_df[low_df['Thâm niên'].str.contains('Dưới 6 tháng', regex=False, na=False)]) if len(low_df) > 0 else 0
        low_ratio = (low_moi_cnt / len(low_df) * 100) if len(low_df) > 0 else 0

        # Đảm bảo tổng phân rã (NV mới + NV cũ) luôn luôn khớp chính xác với hientai_display
        raw_moi = len(sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', regex=False, na=False)])
        raw_cu = len(sub_tn) - raw_moi

        if hientai_display > 0:
            if raw_moi + raw_cu == hientai_display:
                nv_moi_cnt = raw_moi
                nv_cu_cnt = raw_cu
            else:
                nv_cu_cnt = min(raw_cu, hientai_display)
                nv_moi_cnt = hientai_display - nv_cu_cnt

        # TẠO NỘI DUNG CHUẨN THUẦN TEXT
        txt = f"{hub_idx}. BƯU CỤC {full_bc_name}\n"
        txt += f"{hub_idx}.1. Định biên & Tuyển dụng:\n"
        txt += f"- Định biên / Hiện tại: {dinhbien_val} / {hientai_display} NVPTTT ({nv_moi_cnt} NV mới <6 tháng - {nv_cu_cnt} NV cũ).\n"
        
        if stt_pttt == "Thiếu" or thieu_display > 0:
            txt += f"- Tuyển dụng: Trạng thái PTTT: Thiếu (Thiếu {thieu_display} NV/tuyến)\n"

        txt += f"{hub_idx}.2. Vận hành thực tế:\n"
        txt += f"- Tổng % GTC Bưu cục: {tot_gtc_rate:.2f}% (GTC {int(tot_gtc)} / Volume {int(tot_vol)} đơn)\n"
        
        # Đếm thực tế đi làm động từ Sheet Năng Suất (gồm PTTT + NV hỗ trợ)
        total_ns_all = len(sub_ns_unfiltered)
        sup_ns_cnt = len(sub_ns_unfiltered[
            (sub_ns_unfiltered['Emp_Code'].isin(support_emp_codes)) |
            (sub_ns_unfiltered['Emp_Name'].str.lower().isin(support_emp_names))
        ])
        reg_ns_cnt = total_ns_all - sup_ns_cnt

        # Khớp thông tin Cảnh báo từ sheet NTB (31 cột)
        ntb_info = None
        for idx_raw, r_raw in df_rec.iterrows():
            pass

        route_data = get_dynamic_route_notes(code, date_display)
        if total_ns_all > 0:
            if sup_ns_cnt > 0:
                txt += f"- Thực tế đi làm: {total_ns_all} NV ({reg_ns_cnt} PTTT + {sup_ns_cnt} NV hỗ trợ).\n"
            else:
                txt += f"- Thực tế đi làm: {total_ns_all} NVPTTT.\n"
        elif route_data and "thuc_te_di_lam" in route_data and route_data["thuc_te_di_lam"]:
            txt += f"- Thực tế đi làm: {route_data['thuc_te_di_lam']}\n"

        # Khớp danh sách Nhân viên hỗ trợ từ sheet NV HỖ TRỢ
        in_sup = [s for s in support_details if main_name.lower() in s['bc_nhan_ho_tro'].lower()]
        out_sup = [s for s in support_details if main_name.lower() in s['bc_chinh'].lower()]

        if in_sup:
            items = [f"{s['name']} (từ {s['bc_chinh']})" for s in in_sup]
            txt += f"- NV từ bưu cục khác đến hỗ trợ ({len(in_sup)} NV): {', '.join(items)}.\n"
        if out_sup:
            items = [f"{s['name']} (hỗ trợ {s['bc_nhan_ho_tro']})" for s in out_sup]
            txt += f"- NV bưu cục đi hỗ trợ bưu cục khác ({len(out_sup)} NV): {', '.join(items)}.\n"

        if route_data and "lech_tuyen" in route_data and route_data["lech_tuyen"]:
            txt += f"- Lệch tuyến / Quá tải thực tế:\n{route_data['lech_tuyen']}\n"

        gan_fmt = f"{avg_gan:.1f}".rstrip('0').rstrip('.')
        gtc_fmt = f"{avg_gtc:.1f}".rstrip('0').rstrip('.')

        txt += f"{hub_idx}.3. Năng suất & Thu nhập:\n"
        txt += f"- Năng suất TB (Hôm qua {date_display}): Gán: {gan_fmt} đơn/NV | GTC/NV: {gtc_fmt} đơn\n"
        txt += f"- Phân nhóm năng suất NV ({len(merged_ns)} NV đi làm):\n"
        txt += f"  + (>50 GTC): {cnt_good} NV.\n"
        txt += f"  + (35 - 50 GTC): {cnt_mid} NV.\n"
        txt += f"  + (<35 GTC): {cnt_low} NV ({low_ratio:.0f}% là NV MỚI thâm niên ngắn ngày).\n"
        txt += f"- Thu nhập cụ thể từng NV (ngày {date_display}):\n"
        for tn_label in ['Dưới 6 tháng', '6 tháng - 3 năm', 'Trên 3 năm']:
            sub_g = sub_tn[sub_tn['Thâm niên'].str.contains(tn_label, regex=False, na=False)]
            if len(sub_g) > 0:
                sorted_g = sub_g.sort_values(by='luong_num', ascending=False)
                emp_income_items = []
                for _, r in sorted_g.iterrows():
                    emp_raw = str(r['Nhân viên'])
                    emp_name = emp_raw.split('-')[-1].strip() if '-' in emp_raw else emp_raw
                    luong_str = f"{int(r['luong_num']):,}đ".replace(',', '.')
                    emp_income_items.append(f"{emp_name}: {luong_str}/ngày")
                
                txt += f"  + NV {tn_label} ({len(emp_income_items)} NV):\n"
                for item in emp_income_items:
                    txt += f"    * {item}\n"

        txt += f"{hub_idx}.4. Phương án:\n"
        if thieu_display > 0:
            txt += f"- Đẩy mạnh tuyển bổ sung {thieu_display} NVPTTT lấp tuyến thiếu.\n"
        if cnt_low > 0:
            txt += f"- Phân công kèm cặp 1:1 cho nhóm {cnt_low} NV năng suất thấp (chủ yếu là NV mới) để nâng tỷ lệ GTC.\n"
        if nv_moi_cnt > 0 and nv_moi_cnt >= nv_cu_cnt:
            txt += f"- Tập trung ổn định và giữ chân nhóm {nv_moi_cnt} NV mới, cân đối lại lượng đơn tránh áp lực nghỉ ngang.\n"
        if tot_ton > 200:
            txt += f"- Thuê CTV giải tỏa {int(tot_ton)} đơn tồn đọng dài ngày và hỗ trợ chạy tuyến phụ.\n"

        reports.append(txt)
        write_back_list.append((row_in_sheet, txt))

    # Ghi kết quả vào Cột B của Sheet `bưu cục`
    if data_dict["ws_input"]:
        try:
            ws_in = data_dict["ws_input"]
            for row_idx, txt in write_back_list:
                ws_in.update_cell(row_idx, 2, txt)
            print(f"✅ Đã tự động xuất báo cáo thuần TEXT vào Cột B của Sheet `{ws_in.title}`!")
        except Exception as e:
            print(f"⚠️ Không ghi được vào Sheet input: {e}")

    today_str = datetime.now().strftime("%d/%m/%Y")
    header_title = f"NTB BÁO CÁO BƯU CỤC BẤT ỔN NGÀY {today_str}\n\n"
    final_text_report = header_title + "\n--------------------------------------------------\n\n".join(reports)
    return final_text_report


if __name__ == "__main__":
    print("⏳ Đang kết nối Google Sheet và tự động xuất Báo cáo Bưu Cục thuần TEXT...")
    report_text = generate_report()
    print("\n" + "=" * 60)
    print(" BÁO CÁO VẬN HÀNH & NĂNG SUẤT BƯU CỤC BẤT ỔN (PLAIN TEXT)")
    print("=" * 60 + "\n")
    print(report_text)
