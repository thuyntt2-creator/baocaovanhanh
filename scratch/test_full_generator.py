import sys
import os
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

# ================= CẤU HÌNH GOOGLE SHEET =================
SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"
DEFAULT_HUBS = ["22830000", "20942000"]


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
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f"❌ Không tìm thấy file {CREDENTIALS_FILE}")

    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
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

    # Read NV HỖ TRỢ Sheet
    ws_hotro = None
    for key in ['nv hỗ trợ', 'nv ho tro', 'nv_ho_tro', 'ho tro', 'hỗ trợ']:
        if key in worksheets:
            ws_hotro = worksheets[key]
            break

    support_emp_codes = set()
    support_emp_names = set()
    if ws_hotro:
        vals_hotro = ws_hotro.get_all_values()
        for row in vals_hotro:
            if row and row[0].strip():
                raw_val = row[0].strip()
                if raw_val.lower() in ['nhân viên', 'nhan vien', 'stt', 'mã nv', 'mã nhân viên']:
                    continue
                code = raw_val.split('_')[0].split('-')[0].strip()
                if code.isdigit():
                    support_emp_codes.add(code)
                name = raw_val.split('_')[-1].split('-')[-1].strip()
                if name and not name.isdigit():
                    support_emp_names.add(name.lower())

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
        "support_emp_names": support_emp_names
    }


# Map ghi chú vận hành tuyến thực tế theo mã Bưu cục
ROUTE_NOTES = {
    "22830000": {
        "thuc_te_di_lam": "17 NV (15 PTTT + 2 XL). GTC Ca 1: 56.41% (154/273 đơn).",
        "lech_tuyen": """  Cam Phúc Nam (70-80 đơn): 01 NV chạy ca chiều (14–15h) do đi biển đêm -> Đang tuyển thay thế
  Cam Phú (100-120 đơn): 02 NV (01 NV mới + 01 NV xin nghỉ từ 15/8) -> Tuyển thêm 1-2 NV
  Cam Linh (120-150 đơn): 01 NV nghỉ từ 10/8 nghỉ ngang -> Nguy cơ bỏ tuyến, tuyển gấp 2 NV
  Cam Phước Đông & Cam Thịnh: 02 NV mới , 1 nv hỗ trợ gánh tuyến xã xa 8–10km -> Quá tải, tuyển gấp 3 NV
  Cam Lợi (90-100 đơn): 01 NV mới sức khỏe yếu đã nghỉ -> Tuyển gấp 1 NV
  Cụm đảo (Bình Ba/Hưng/Lập): 03 NV 
 Cam Thuận, Cam Lộc, Ba Ngòi: 04 NV cứng -> Vận hành ổn định."""
    },
    "20942000": {
        "thuc_te_di_lam": "Clear hàng còn thấp, chưa ổn định. Trong đó có 2 NV di chuyển hỗ trợ Đức Trọng 1. Bưu cục chưa sắp tuyến cố định.",
        "lech_tuyen": """  Thị trấn Di Linh (5/5 NV): 03 NV thâm niên <3 tháng (năng suất giao thấp 45–60 đơn/ngày) + 02 NV thâm niên >3 tháng (năng suất giao 70–80 đơn/ngày).
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
  Xã Tân Thượng: 01 NV thâm niên <3 tháng (năng suất giao 40–50 đơn/ngày"""
    },
    "21477000": {
        "thuc_te_di_lam": "16 NV (14 PTTT + 2 XL). GTC Ca 1: 69.32% (122/176 đơn).",
        "lech_tuyen": """  Đạ Nhim, Đạ Sar: 02 NV chạy tuyến chành (Đúng tuyến).
  Bidoup: 01 NV kèm NV mới, nghỉ cuối tháng -> Nguy cơ trống tuyến.
  19/5 - TDP Đan Kia & Bạch Đằng - Cao Thắng: Hàng tăng, rộng -> Quá tải, phải bỏ lại 1-2 khu vực luân phiên.
  Mănglin - Đạ Phú: 02 NV mới, dân thưa, diện rộng -> Năng suất thấp, tồn đọng lớn.
  Nguyễn Hoàng, Xô Viết: Địa bàn xa kho -> Tải hàng chậm xuống tuyến.
  Đưng Kno - Xã Lát: 01 NV chạy 2 tuyến ngược chiều -> Mỗi ngày chỉ chạy được 1 tuyến, bỏ dở tuyến còn lại."""
    }
}


def generate_report():
    data_dict = load_all_data()
    target_hubs = data_dict["target_hubs"]
    support_emp_codes = data_dict["support_emp_codes"]
    support_emp_names = data_dict["support_emp_names"]

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

        # 1. Match in Tuyển dụng sheet
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

        hrbp = rec_row[6] if rec_row is not None else "N/A"
        stt_pttt = rec_row[7] if rec_row is not None else "N/A"
        thieu_pttt = rec_row[8] if rec_row is not None else "0"
        hientai_pttt = rec_row[9] if rec_row is not None else "0"
        dinhbien_pttt = rec_row[10] if rec_row is not None else "0"

        # 2. Match in Data Sheet
        sub_d = df_data[
            df_data['Chi tiết'].str.contains(code, regex=False, na=False) |
            df_data['Chi tiết'].str.contains(bc_short_name, regex=False, na=False) |
            df_data['Chi tiết'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ]
        
        if len(sub_d) > 0:
            latest_date = sub_d['Time'].max()
            sub_d_latest = sub_d[sub_d['Time'] == latest_date]

            row_ca1 = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Ca 1', regex=False, na=False)]
            row_ton = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Tồn', regex=False, na=False)]

            vol_ca1 = clean_num(row_ca1['Volume'].values[0]) if len(row_ca1) > 0 else 0
            gan_ca1 = clean_num(row_ca1['Sản Lượng Gán'].values[0]) if len(row_ca1) > 0 else 0
            gtc_ca1 = clean_num(row_ca1['Sản Lượng Giao Thành Công'].values[0]) if len(row_ca1) > 0 else 0
            ton_ca1 = clean_num(row_ca1['Sản Lượng Tồn'].values[0]) if len(row_ca1) > 0 else 0
            chuagan_ca1 = clean_num(row_ca1['Sản Lượng Chưa Gán'].values[0]) if len(row_ca1) > 0 else 0

            vol_ton = clean_num(row_ton['Volume'].values[0]) if len(row_ton) > 0 else 0
            gan_ton = clean_num(row_ton['Sản Lượng Gán'].values[0]) if len(row_ton) > 0 else 0
            gtc_ton = clean_num(row_ton['Sản Lượng Giao Thành Công'].values[0]) if len(row_ton) > 0 else 0
            ton_ton = clean_num(row_ton['Sản Lượng Tồn'].values[0]) if len(row_ton) > 0 else 0
            chuagan_ton = clean_num(row_ton['Sản Lượng Chưa Gán'].values[0]) if len(row_ton) > 0 else 0
        else:
            vol_ca1 = gan_ca1 = gtc_ca1 = ton_ca1 = chuagan_ca1 = 0
            vol_ton = gan_ton = gtc_ton = ton_ton = chuagan_ton = 0

        tot_vol = vol_ca1 + vol_ton
        tot_gtc = gtc_ca1 + gtc_ton
        tot_gtc_rate = (tot_gtc / tot_vol * 100) if tot_vol > 0 else 0

        # 3. Match in Thu Nhập Sheet (LOẠI BỎ NV HỖ TRỢ)
        sub_tn = df_tn[
            df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
            df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
            df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ].copy()

        # Lọc bỏ NV HỖ TRỢ
        sub_tn = sub_tn[
            (~sub_tn['Emp_Code'].isin(support_emp_codes)) &
            (~sub_tn['Emp_Name'].str.lower().isin(support_emp_names))
        ]

        nv_moi_cnt = len(sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', regex=False, na=False)])
        nv_cu_cnt = len(sub_tn) - nv_moi_cnt

        avg_gan = sub_tn['don_gan_num'].mean() if len(sub_tn) > 0 else 0
        avg_gtc = sub_tn['gtc_num'].mean() if len(sub_tn) > 0 else 0

        # 4. Match in Năng Suất Sheet (Lấy ngày mới nhất & LOẠI BỎ NV HỖ TRỢ)
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

        # Lọc bỏ NV HỖ TRỢ khỏi Năng Suất
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

        # TẠO NỘI DUNG CHUẨN THUẦN TEXT
        txt = f"{hub_idx}. BƯU CỤC {full_bc_name}\n"
        txt += f"{hub_idx}.1. Định biên & Tuyển dụng:\n"
        txt += f"- Định biên / Hiện tại: {int(clean_num(dinhbien_pttt))} / {int(clean_num(hientai_pttt))} NVPTTT ({nv_moi_cnt} NV mới <6 tháng - {nv_cu_cnt} NV cũ).\n"
        
        if stt_pttt == "Thiếu" or clean_num(thieu_pttt) > 0:
            txt += f"- Tuyển dụng: Trạng thái PTTT: Thiếu (Thiếu {int(clean_num(thieu_pttt))} NV/tuyến)\n"

        txt += f"{hub_idx}.2. Vận hành thực tế:\n"
        txt += f"- Tổng % GTC Bưu cục (Mới + Tồn): {tot_gtc_rate:.2f}% (GTC {int(tot_gtc)} / Volume {int(tot_vol)} đơn)\n"
        
        # Ghi chú vận hành tuyến thực tế
        route_data = ROUTE_NOTES.get(code, {})
        if route_data:
            if "thuc_te_di_lam" in route_data and route_data["thuc_te_di_lam"]:
                txt += f"- Thực tế đi làm: {route_data['thuc_te_di_lam']}\n"
            if "lech_tuyen" in route_data and route_data["lech_tuyen"]:
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
        if clean_num(thieu_pttt) > 0:
            txt += f"- Đẩy mạnh tuyển bổ sung {int(clean_num(thieu_pttt))} NVPTTT lấp tuyến thiếu.\n"
        if cnt_low > 0:
            txt += f"- Phân công kèm cặp 1:1 cho nhóm {cnt_low} NV năng suất thấp (chủ yếu là NV mới) để nâng tỷ lệ GTC.\n"
        if nv_moi_cnt > 0 and nv_moi_cnt >= nv_cu_cnt:
            txt += f"- Tập trung ổn định và giữ chân nhóm {nv_moi_cnt} NV mới, cân đối lại lượng đơn tránh áp lực nghỉ ngang.\n"
        if ton_ton > 200:
            txt += f"- Thuê CTV giải tỏa {int(ton_ton)} đơn tồn đọng dài ngày và hỗ trợ chạy tuyến phụ.\n"

        reports.append(txt)
        write_back_list.append((row_in_sheet, txt))

    today_str = datetime.now().strftime("%d/%m/%Y")
    header_title = f"NTB BÁO CÁO BƯU CỤC BẤT ỔN NGÀY {today_str}\n\n"
    final_text_report = header_title + "\n--------------------------------------------------\n\n".join(reports)
    return final_text_report

if __name__ == "__main__":
    rep = generate_report()
    print(rep)
