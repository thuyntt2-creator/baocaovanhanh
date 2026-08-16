import os
import io
import sys
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Fix encoding cho Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Đường dẫn ghi file báo cáo artifact
ARTIFACT_DIR = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\aa320bb2-4d48-485f-b844-2f40dbbcbca0"
REPORT_PATH = os.path.join(ARTIFACT_DIR, "analysis_duplicates.md")

def check_and_generate_report():
    print(f"📖 Connecting to sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc dữ liệu DRAFT
    ws_draft = sh.worksheet("DRAFT")
    draft_data = ws_draft.get_all_values()
    if len(draft_data) < 2:
        print("Tab DRAFT trống.")
        return
        
    draft_headers = [h.strip() for h in draft_data[0]]
    # Tìm kiếm chỉ số cột động trong DRAFT
    ward_id_col_idx = -1
    prov_idx = -1
    dist_idx = -1
    ward_idx = -1
    hub_idx = -1
    for idx, h in enumerate(draft_headers):
        h_lower = h.lower()
        if "id phường/xã" in h_lower or "id phuong/xa" in h_lower:
            ward_id_col_idx = idx
        elif "tỉnh" in h_lower:
            prov_idx = idx
        elif "quận" in h_lower or "huyện" in h_lower:
            dist_idx = idx
        elif "phường" in h_lower or "xã" in h_lower:
            if "id" not in h_lower:
                ward_idx = idx
        elif "bưu cục" in h_lower:
            hub_idx = idx

    if ward_id_col_idx == -1:
        print("Không tìm thấy cột ID phường/xã trong DRAFT")
        return
        
    draft_wards = {}
    for r_idx, row in enumerate(draft_data[1:]):
        row_num = r_idx + 2
        if len(row) <= ward_id_col_idx:
            continue
        w_id = row[ward_id_col_idx].strip()
        if not w_id:
            continue
            
        ward_name = row[ward_idx].strip() if ward_idx != -1 and len(row) > ward_idx else "Không rõ"
        district = row[dist_idx].strip() if dist_idx != -1 and len(row) > dist_idx else "Không rõ"
        province = row[prov_idx].strip() if prov_idx != -1 and len(row) > prov_idx else "Không rõ"
        hub = row[hub_idx].strip() if hub_idx != -1 and len(row) > hub_idx else "Không rõ"
        
        info = {
            "row_num": row_num,
            "id": w_id,
            "ward": ward_name,
            "district": district,
            "province": province,
            "hub": hub
        }
        if w_id not in draft_wards:
            draft_wards[w_id] = []
        draft_wards[w_id].append(info)

    # Trùng lặp nội bộ DRAFT
    draft_duplicates = {k: v for k, v in draft_wards.items() if len(v) > 1}
    
    # 2. Đọc dữ liệu Đang OFF
    ws_off = sh.worksheet("Đang OFF")
    off_data = ws_off.get_all_values()
    off_headers = [h.strip() for h in off_data[0]]
    
    # Tìm kiếm chỉ số cột động trong Đang OFF
    off_ward_id_idx = -1
    off_prov_idx = -1
    off_dist_idx = -1
    off_ward_idx = -1
    off_hub_idx = -1
    for idx, h in enumerate(off_headers):
        h_lower = h.lower()
        if "id phường/xã" in h_lower or "id phuong/xa" in h_lower:
            off_ward_id_idx = idx
        elif "tỉnh" in h_lower:
            off_prov_idx = idx
        elif "quận" in h_lower or "huyện" in h_lower:
            off_dist_idx = idx
        elif "phường" in h_lower or "xã" in h_lower:
            if "id" not in h_lower:
                off_ward_idx = idx
        elif "bưu cục" in h_lower:
            off_hub_idx = idx

    if off_ward_id_idx == -1:
        # fallback
        for idx, h in enumerate(off_headers):
            if "id" in h.lower():
                off_ward_id_idx = idx
                break
    if off_ward_id_idx == -1:
        off_ward_id_idx = 3 # mặc định nếu không khớp

    off_wards = {}
    for r_idx, row in enumerate(off_data[1:]):
        row_num = r_idx + 2
        if len(row) <= off_ward_id_idx:
            continue
        w_id = row[off_ward_id_idx].strip()
        if not w_id:
            continue
            
        ward_name = row[off_ward_idx].strip() if off_ward_idx != -1 and len(row) > off_ward_idx else "Không rõ"
        district = row[off_dist_idx].strip() if off_dist_idx != -1 and len(row) > off_dist_idx else "Không rõ"
        province = row[off_prov_idx].strip() if off_prov_idx != -1 and len(row) > off_prov_idx else "Không rõ"
        hub = row[off_hub_idx].strip() if off_hub_idx != -1 and len(row) > off_hub_idx else "Không rõ"
        
        info = {
            "row_num": row_num,
            "id": w_id,
            "ward": ward_name,
            "district": district,
            "province": province,
            "hub": hub
        }
        if w_id not in off_wards:
            off_wards[w_id] = []
        off_wards[w_id].append(info)

    # So sánh trùng lặp DRAFT vs Đang OFF
    overlap_with_off = []
    for w_id, draft_infos in draft_wards.items():
        if w_id in off_wards:
            for d_info in draft_infos:
                for o_info in off_wards[w_id]:
                    overlap_with_off.append({
                        "id": w_id,
                        "draft_info": d_info,
                        "off_info": o_info
                    })

    # 3. Tạo file báo cáo Markdown
    md = []
    md.append("# Báo Cáo Kiểm Tra Trùng Lặp ID Phường Xã\n")
    md.append(f"*Thời gian quét: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}*\n")
    md.append(f"*File Google Sheet:* [NTB- FOLLOW OFF Tuyến](https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=488516454)\n")
    
    # Section 1: Trùng lặp trong nội bộ DRAFT
    md.append("## 1. Trùng lặp trong nội bộ tab 'DRAFT'\n")
    if draft_duplicates:
        md.append("> [!IMPORTANT]")
        md.append(f"> Phát hiện **{len(draft_duplicates)}** nhóm ID phường/xã bị trùng lặp nhiều lần ngay trong đề xuất của tab 'DRAFT'.\n")
        
        md.append("| ID Phường/Xã | Tỉnh | Quận/Huyện | Phường/Xã | Bưu Cục | Dòng trong Sheet |")
        md.append("| :--- | :--- | :--- | :--- | :--- | :--- |")
        for w_id, infos in draft_duplicates.items():
            for idx, info in enumerate(infos):
                md.append(f"| `{w_id}` | {info['province']} | {info['district']} | {info['ward']} | {info['hub']} | Dòng {info['row_num']} |")
    else:
        md.append("> [!NOTE]")
        md.append("> Không phát hiện trùng lặp ID phường/xã nào trong nội bộ tab DRAFT.\n")
        
    # Section 2: Trùng lặp DRAFT vs Đang OFF
    md.append("\n## 2. Đề xuất trong 'DRAFT' bị trùng với tuyến đã tắt trong tab 'Đang OFF'\n")
    if overlap_with_off:
        md.append("> [!WARNING]")
        md.append(f"> Phát hiện **{len(overlap_with_off)}** đề xuất tắt tuyến trong 'DRAFT' thực tế **đã có sẵn/đang hoạt động tắt** trong tab 'Đang OFF'. Bạn có thể loại bỏ các đề xuất này vì chúng đã được off từ trước.\n")
        
        md.append("| ID Phường/Xã | Phường/Xã | Quận/Huyện | Tỉnh | Đề xuất (DRAFT) | Hiện trạng (Đang OFF) |")
        md.append("| :--- | :--- | :--- | :--- | :--- | :--- |")
        for overlap in overlap_with_off:
            d = overlap["draft_info"]
            o = overlap["off_info"]
            md.append(f"| `{overlap['id']}` | {d['ward']} | {d['district']} | {d['province']} | Dòng **{d['row_num']}** (BC: {d['hub']}) | Dòng **{o['row_num']}** (BC: {o['hub']}) |")
    else:
        md.append("> [!NOTE]")
        md.append("> Không có đề xuất nào trong DRAFT bị trùng với danh sách đang OFF.\n")

    # Ghi file
    if not os.path.exists(ARTIFACT_DIR):
        os.makedirs(ARTIFACT_DIR)
        
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
    print(f"📊 Báo cáo trùng lặp đã được tạo thành công tại: {REPORT_PATH}")

if __name__ == "__main__":
    check_and_generate_report()
