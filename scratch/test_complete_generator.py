# -*- coding: utf-8 -*-
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"
DEFAULT_HUBS = ["20942000", "22830000"]


def clean_num(val):
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
        for row in vals[1:]:
            if row and row[0].strip():
                target_hubs.append(row[0].strip())

    if not target_hubs:
        target_hubs = DEFAULT_HUBS

    return {
        "client": client,
        "spreadsheet": spreadsheet,
        "ws_input": ws_input,
        "target_hubs": target_hubs,
        "rec": spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values(),
        "data": spreadsheet.worksheet("data").get_all_values(),
        "tn": spreadsheet.worksheet("thu nhập").get_all_values(),
        "ns": spreadsheet.worksheet("năng suất").get_all_values()
    }


def generate_report():
    data_dict = load_all_data()
    target_hubs = data_dict["target_hubs"]

    df_rec = pd.DataFrame(data_dict["rec"])
    
    data_vals = data_dict["data"]
    df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

    tn_vals = data_dict["tn"]
    df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
    df_tn['don_gan_num'] = df_tn['Số đơn gán Giao'].apply(clean_num)
    df_tn['gtc_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
    df_tn['luong_num'] = df_tn['Tổng lương'].apply(clean_num)
    df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].strip())

    ns_vals = data_dict["ns"]
    df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
    df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
    df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)
    df_ns['%GTC_num'] = df_ns['%GTC'].str.replace('%', '').str.replace(',', '.').apply(clean_num)
    df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].strip())

    reports = []

    for hub_idx, hub_query in enumerate(target_hubs, 1):
        query_clean = str(hub_query).strip()

        # Match in Tuyển dụng sheet
        rec_row = None
        for idx, r in df_rec.iterrows():
            row_str = " ".join([str(x) for x in r])
            if query_clean.lower() in row_str.lower():
                rec_row = r
                break

        code = rec_row[2] if rec_row is not None else query_clean
        bc_short_name = rec_row[3] if rec_row is not None else query_clean
        full_bc_name = rec_row[4] if rec_row is not None else query_clean
        
        # Keywords for matching in other sheets
        # Extract main name e.g. "Di Linh" or "Cam Linh"
        main_name = bc_short_name.split(')')[-1].strip() if ')' in bc_short_name else query_clean

        hrbp = rec_row[6] if rec_row is not None else "N/A"
        stt_pttt = rec_row[7] if rec_row is not None else "N/A"
        thieu_pttt = rec_row[8] if rec_row is not None else "0"
        hientai_pttt = rec_row[9] if rec_row is not None else "0"
        dinhbien_pttt = rec_row[10] if rec_row is not None else "0"
        stt_xl = rec_row[18] if rec_row is not None and len(rec_row) > 18 else "Đủ"
        thieu_xl = rec_row[19] if rec_row is not None and len(rec_row) > 19 else "0"

        # Match in Data Sheet
        sub_d = df_data[
            df_data['Chi tiết'].str.contains(code, na=False) |
            df_data['Chi tiết'].str.contains(bc_short_name, regex=False, na=False) |
            df_data['Chi tiết'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ]
        
        if len(sub_d) > 0:
            latest_date = sub_d['Time'].max()
            sub_d_latest = sub_d[sub_d['Time'] == latest_date]

            row_ca1 = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Ca 1', na=False)]
            row_ton = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Tồn', na=False)]

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

        # Match in Thu Nhập Sheet
        sub_tn = df_tn[
            df_tn['Bưu cục'].str.contains(code, na=False) |
            df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
            df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
        ]
        nv_moi_cnt = len(sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', na=False)])
        nv_cu_cnt = len(sub_tn) - nv_moi_cnt

        avg_gan = sub_tn['don_gan_num'].mean() if len(sub_tn) > 0 else 0
        avg_gtc = sub_tn['gtc_num'].mean() if len(sub_tn) > 0 else 0

        luong_g01 = sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', na=False)]['luong_num'].mean() if len(sub_tn) > 0 else 0
        luong_g02 = sub_tn[sub_tn['Thâm niên'].str.contains('6 tháng - 3 năm', na=False)]['luong_num'].mean() if len(sub_tn) > 0 else 0
        luong_g03 = sub_tn[sub_tn['Thâm niên'].str.contains('Trên 3 năm', na=False)]['luong_num'].mean() if len(sub_tn) > 0 else 0

        # Match in Năng Suất Sheet (Match date containing '22')
        sub_ns = df_ns[
            (df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
             df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)) &
            (df_ns['Ngay'].str.contains('22'))
        ]
        if len(sub_ns) == 0:
            sub_ns = df_ns[
                df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
                df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
            ]

        merged_ns = pd.merge(sub_ns, df_tn[['Emp_Code', 'Thâm niên', 'Ngày vào làm']], on='Emp_Code', how='left')

        cnt_good = len(merged_ns[merged_ns['TongDonGTC_num'] >= 50])
        cnt_mid = len(merged_ns[(merged_ns['TongDonGTC_num'] >= 35) & (merged_ns['TongDonGTC_num'] < 50)])
        cnt_low = len(merged_ns[merged_ns['TongDonGTC_num'] < 35])

        low_df = merged_ns[merged_ns['TongDonGTC_num'] < 35]
        low_moi_cnt = len(low_df[low_df['Thâm niên'].str.contains('Dưới 6 tháng', na=False)]) if len(low_df) > 0 else 0
        low_ratio = (low_moi_cnt / len(low_df) * 100) if len(low_df) > 0 else 0

        # BUILD REPORT TEXT
        md = f"### {hub_idx}. BƯU CỤC {full_bc_name}\n\n"
        md += f"**{hub_idx}.1. Định biên & Tuyển dụng:**\n"
        md += f"- **Định biên / Hiện tại:** **{int(clean_num(dinhbien_pttt))} / {int(clean_num(hientai_pttt))} NVPTTT** ({nv_moi_cnt} NV mới <6 tháng - {nv_cu_cnt} NV cũ).\n"
        
        stt_str = f"Thiếu **{thieu_pttt} NV/tuyến**" if stt_pttt == "Thiếu" or clean_num(thieu_pttt) > 0 else "Đủ"
        xl_str = f"Thiếu {thieu_xl} NV" if stt_xl == "Thiếu" or clean_num(thieu_xl) > 0 else "1/1 NV (Đủ)"
        md += f"- **Tuyển dụng:** Trạng thái PTTT: **{stt_pttt}** ({stt_str}) | NVXL: **{xl_str}** | HRBP: {hrbp}.\n\n"

        md += f"**{hub_idx}.2. Vận hành thực tế:**\n"
        md += f"- **Hàng mới Ca 1 ({int(vol_ca1)} đơn):** Gán {int(gan_ca1)} đơn, GTC {int(gtc_ca1)} đơn, Tồn {int(ton_ca1)} đơn, Chưa gán {int(chuagan_ca1)} đơn.\n"
        md += f"- **Hàng tồn ({int(vol_ton)} đơn):** Gán {int(gan_ton)} đơn, GTC {int(gtc_ton)} đơn, Tồn {int(ton_ton)} đơn, Chưa gán {int(chuagan_ton)} đơn.\n"
        md += f"- **Tổng % GTC Bưu cục (Mới + Tồn):** **{tot_gtc_rate:.2f}%** *(GTC {int(tot_gtc)} / Volume {int(tot_vol)} đơn)*\n\n"

        md += f"**{hub_idx}.3. Năng suất & Thu nhập:**\n"
        md += f"- **Năng suất TB (Hôm qua 22/07):** Gán: **{avg_gan:.1f} đơn/NV** | GTC/NV: **{avg_gtc:.1f} đơn**\n"
        md += f"- **Phân nhóm năng suất NV ({len(merged_ns)} NV đi làm):**\n"
        md += f"  - *Gánh tuyến tốt (>50 GTC):* **{cnt_good} NV**.\n"
        md += f"  - *Trung bình (35 - 50 GTC):* **{cnt_mid} NV**.\n"
        md += f"  - *Cần hỗ trợ gấp (GTC thấp):* **{cnt_low} NV** ({low_ratio:.0f}% là NV MỚI thâm niên ngắn ngày).\n"
        md += f"- **Thu nhập cụ thể (TB/ngày):**\n"
        if luong_g01 > 0: md += f"  + NV Dưới 6 tháng: **{luong_g01:,.0f} đ**\n"
        if luong_g02 > 0: md += f"  + NV Trên 6 tháng - 3 năm: **{luong_g02:,.0f} đ**\n"
        if luong_g03 > 0: md += f"  + NV Trên 3 năm: **{luong_g03:,.0f} đ**\n"
        md += "\n"

        md += f"**{hub_idx}.4. Phương án:**\n"
        if clean_num(thieu_pttt) > 0:
            md += f"- HRBP ({hrbp}) đẩy mạnh tuyển bổ sung **{thieu_pttt} NVPTTT** lấp tuyến thiếu.\n"
        md += f"- Phân công kèm cặp 1:1 cho nhóm {cnt_low} NV năng suất thấp (chủ yếu là NV mới) để nâng tỷ lệ GTC.\n"
        if ton_ton > 200:
            md += f"- Thuê CTV giải tỏa {int(ton_ton)} đơn tồn đọng dài ngày và hỗ trợ chạy tuyến phụ.\n"

        reports.append(md)

    final_report = "\n---\n\n".join(reports)
    return final_report


if __name__ == "__main__":
    print("⏳ Đang kết nối Google Sheet và tự động xuất Báo cáo Bưu Cục...")
    report_text = generate_report()
    print("\n" + "=" * 65)
    print(" BÁO CÁO VẬN HÀNH & NĂNG SUẤT BƯU CỤC BẤT ỔN")
    print("=" * 65 + "\n")
    print(report_text)

