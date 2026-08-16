# -*- coding: utf-8 -*-
"""
Tạo ảnh báo cáo NVPTT mức thấp (nhóm theo AM) và gửi vào kênh GTalk nội bộ GHN.
Ghép logic vẽ ảnh (make_report_image.py) + logic gửi GTalk (theo mẫu upload_image_to_gtalk
đã có trong script rot_lc_gtalk.py của bạn).
"""
import os
import json
import time
from datetime import datetime

import requests
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont

try:
    from dotenv import load_dotenv
    BASE_DIR_EARLY = os.path.dirname(os.path.abspath(__file__))
    # Thử load .env cùng thư mục script trước, sau đó thử vị trí .env dùng chung với các script khác
    env_candidates = [
        os.path.join(BASE_DIR_EARLY, ".env"),
        r"c:\Users\lap4all\Desktop\New folder\.env",
    ]
    for env_path in env_candidates:
        if os.path.exists(env_path):
            load_dotenv(dotenv_path=env_path, override=True)
            print(f"🔑 Đã load .env từ: {env_path}")
            break
    else:
        load_dotenv(override=True)
except ImportError:
    print("⚠️ Chưa cài python-dotenv (pip install python-dotenv) — sẽ dùng token mặc định trong code.")

# ===== CẤU HÌNH GOOGLE SHEET =====
BAOCAO_SHEET_KEY = "1-p9VUXndK_7BoiT-a81UfTCbUi953XNmVBoXaTGis_c"
BAOCAO_TAB_NAME = "BaoCao"

def get_credentials_path():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json"
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def parse_pct(val):
    """'71.62%' -> 71.62 ; '' hoặc lỗi -> 0.0"""
    if val is None:
        return 0.0
    s = str(val).strip().replace("%", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_num(val):
    if val is None or str(val).strip() == "":
        return 0
    s = str(val).strip().replace(".", "").replace(",", "")
    try:
        return int(float(s))
    except ValueError:
        return 0


def load_data_from_sheet(sheet_key=BAOCAO_SHEET_KEY, tab_name=BAOCAO_TAB_NAME, threshold_pct=80.0):
    """
    Đọc tab BaoCao, tự dò dòng header (chứa 'Nhân Viên'/'Mã NV'),
    lọc các dòng có Đánh Giá == 'Thấp' (hoặc %GTC < threshold_pct nếu cột Đánh Giá không có),
    rồi group thành cấu trúc AM -> Bưu Cục -> [(NhanVien, Gan, TC, Pct, LTC), ...]
    """
    json_path = get_credentials_path()
    if not json_path:
        raise FileNotFoundError("Không tìm thấy credentials.json để đọc Google Sheet.")

    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = Credentials.from_service_account_file(json_path, scopes=scopes)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(sheet_key)
    ws = sh.worksheet(tab_name)

    all_values = ws.get_all_values()

    # Dò dòng header: dòng có chứa cả 'Nhân Viên' (hoặc 'Nhan Vien') và 'Bưu' (hoặc 'BuuCuc'/'Buu Cuc')
    header_idx = None
    for idx, row in enumerate(all_values):
        row_join = " ".join(row).lower()
        if ("nhân viên" in row_join or "nhan vien" in row_join) and ("bưu" in row_join or "buu" in row_join):
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError("Không tìm thấy dòng header trong tab BaoCao (cần cột 'Bưu Cục' và 'Nhân Viên').")

    headers = [h.strip() for h in all_values[header_idx]]
    data_rows = all_values[header_idx + 1:]

    def col_idx(*candidates):
        for cand in candidates:
            for i, h in enumerate(headers):
                if h.lower().replace(" ", "") == cand.lower().replace(" ", ""):
                    return i
        return None

    i_bc = col_idx("Bưu Cục", "BuuCuc")
    i_am = col_idx("AM")
    i_name = col_idx("Nhân Viên", "NhanVien")
    i_gan = col_idx("Gán Giao", "GanGiao")
    i_tc = col_idx("Giao TC", "GiaoTC")
    i_pct = col_idx("%GTC")
    i_ltc = col_idx("LTC")
    i_danhgia = col_idx("Đánh Giá", "DanhGia")

    missing = [name for name, i in [
        ("Bưu Cục", i_bc), ("AM", i_am), ("Nhân Viên", i_name),
        ("%GTC", i_pct)
    ] if i is None]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc trong tab BaoCao: {missing}. Header đọc được: {headers}")

    # Group AM -> BC -> [(name, gan, tc, pct, ltc)], giữ thứ tự xuất hiện trong sheet
    grouped = {}
    for row in data_rows:
        if len(row) <= i_pct:
            continue

        bc = row[i_bc].strip() if i_bc < len(row) else ""
        am = row[i_am].strip() if i_am < len(row) else ""
        name = row[i_name].strip() if i_name < len(row) else ""
        if not bc or not am or not name:
            continue

        pct = parse_pct(row[i_pct]) if i_pct < len(row) else 0.0
        danh_gia = row[i_danhgia].strip() if (i_danhgia is not None and i_danhgia < len(row)) else ""

        is_thap = (danh_gia == "Thấp") if danh_gia else (pct < threshold_pct and pct > 0)
        if not is_thap:
            continue

        gan = parse_num(row[i_gan]) if (i_gan is not None and i_gan < len(row)) else 0
        tc = parse_num(row[i_tc]) if (i_tc is not None and i_tc < len(row)) else 0
        ltc = parse_num(row[i_ltc]) if (i_ltc is not None and i_ltc < len(row)) else 0

        grouped.setdefault(am, {}).setdefault(bc, []).append((name, gan, tc, pct, ltc))

    # Chuyển sang list cấu trúc DATA mà generate_report_image() cần
    result = []
    for am_name, bc_map in grouped.items():
        result.append({
            "am": am_name,
            "bcs": [(bc_name, staff_list) for bc_name, staff_list in bc_map.items()]
        })

    total_staff = sum(len(s) for am in result for _, s in am["bcs"])
    print(f"📊 Đã lọc được {total_staff} NVPTT mức Thấp, thuộc {len(result)} AM.")
    return result

# ===== CẤU HÌNH GTALK =====
# Lấy từ .env nếu có, fallback theo giá trị mặc định (giống pattern rot_lc_gtalk.py)
GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("NVPTT_THAP_GTALK_CHANNEL_ID") or "2076974545807159296"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_PATH = os.path.join(BASE_DIR, "nvptt_thap_theo_am.png")

FONT_CANDIDATES_BOLD = [
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\seguisb.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]
FONT_CANDIDATES_REG = [
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\segoeui.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]

def _first_existing(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return paths[0]  # để lỗi rõ ràng nếu không tìm thấy font nào

F_BOLD = _first_existing(FONT_CANDIDATES_BOLD)
F_REG = _first_existing(FONT_CANDIDATES_REG)


def font(size, bold=False):
    return ImageFont.truetype(F_BOLD if bold else F_REG, size)


# ===== DỮ LIỆU MẪU (fallback nếu không đọc được sheet, hoặc dùng để test layout) =====
DATA_SAMPLE = [
    {
        "am": "Nguyễn Ngọc Khánh",
        "bcs": [
            ("(BTH) Hàm Thắng", [
                ("Nguyễn Thanh Quốc", 74, 53, 71.62, 4),
                ("Lâm Đức Mỹ", 55, 42, 76.36, 3),
                ("Nguyễn Hoàng Duy", 68, 52, 76.47, 18),
                ("Ngô Thanh Dũng", 49, 39, 79.59, 0),
            ]),
            ("(BTH) Lương Sơn", [
                ("Sử Vĩnh Hưng", 79, 58, 73.42, 4),
                ("Đặng Minh Tuấn", 74, 57, 77.03, 27),
            ]),
            ("(BTH) Mũi Né", [
                ("Phạm Xuân Thịnh", 36, 17, 47.22, 0),
                ("Trần Ngọc Sang", 12, 6, 50.00, 0),
                ("Trần Hải Triều", 54, 34, 62.96, 7),
            ]),
        ]
    },
    {
        "am": "Lê Thanh Nhựt",
        "bcs": [
            ("(BTH) Hàm Tân", [
                ("Dương Ngọc Thuận", 89, 48, 53.93, 2),
                ("Nguyễn Hữu Anh Vũ", 40, 29, 72.50, 1),
                ("Nguyễn Trọng Nghĩa", 81, 64, 79.01, 7),
            ]),
        ]
    },
    {
        "am": "Nguyễn Duy Long",
        "bcs": [
            ("(BTH) Liên Hương", [
                ("Đặng Văn Vinh", 37, 28, 75.68, 1),
                ("Lê Thanh Hiếu", 51, 39, 76.47, 0),
            ]),
        ]
    },
]


# ===== MÀU SẮC =====
C_BG = (255, 255, 255)
C_HEADER_BG = (21, 39, 68)
C_HEADER_ACCENT = (56, 130, 220)
C_HEADER_TEXT = (255, 255, 255)
C_CARD_BORDER = (226, 229, 233)
C_TEXT = (40, 42, 46)
C_MUTED = (120, 124, 130)
C_RED_BG = (250, 227, 227)
C_RED_TEXT = (150, 30, 30)
C_AMBER_BG = (250, 236, 210)
C_AMBER_TEXT = (140, 90, 10)
C_ROW_ALT = (247, 248, 250)
C_GRIDLINE = (235, 237, 240)

AM_PALETTE = [
    ((227, 236, 248), (24, 66, 120)),    # pastel blue
    ((228, 245, 235), (20, 110, 80)),    # pastel green
    ((245, 232, 248), (110, 50, 130)),   # pastel purple
    ((252, 232, 236), (150, 40, 80)),    # pastel pink
    ((252, 240, 220), (150, 95, 20)),    # pastel amber
    ((225, 245, 246), (20, 100, 105)),   # pastel teal
]

W = 980
PAD = 26
ROW_H = 30
COL_BC_W = 150
COL_GAN_W = 85
COL_TC_W = 85
COL_PCT_W = 90
COL_LTC_W = 60


def color_for(pct):
    return (C_RED_BG, C_RED_TEXT) if pct < 60 else (C_AMBER_BG, C_AMBER_TEXT)


def measure(draw, text, f):
    bbox = draw.textbbox((0, 0), text, font=f)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def generate_report_image(data, report_date_str, updated_at_str, out_path):
    f_h1 = font(28, bold=True)
    f_h2 = font(16)
    f_am = font(19, bold=True)
    f_bc = font(13, bold=True)
    f_colhead = font(12, bold=True)
    f_name = font(14)
    f_val = font(14, bold=True)

    def compute_height():
        h = 130
        h += 16
        for am_block in data:
            h += 44
            h += 22
            for _, staff in am_block["bcs"]:
                h += len(staff) * ROW_H
            h += 16
        return h + 20

    H = compute_height()
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)

    # ---- Header ----
    d.rectangle([0, 0, W, 118], fill=C_HEADER_BG)
    d.rectangle([0, 118, W, 121], fill=C_HEADER_ACCENT)
    title = "Cập nhật NVPTT mức thấp"
    sub1 = "%GTC < 80%"
    sub2 = f"Ngày báo cáo {report_date_str}   ·   Cập nhật lúc {updated_at_str}"
    tw, th = measure(d, title, f_h1)
    d.text(((W - tw) / 2, 22), title, font=f_h1, fill=C_HEADER_TEXT)
    sw, sh = measure(d, sub1, f_h2)
    d.text(((W - sw) / 2, 58), sub1, font=f_h2, fill=(150, 190, 235))
    sw2, sh2 = measure(d, sub2, font(13))
    d.text(((W - sw2) / 2, 86), sub2, font=font(13), fill=(190, 205, 225))

    y = 146
    card_x0 = PAD
    card_x1 = W - PAD

    for am_idx, am_block in enumerate(data):
        bar_bg, bar_txt = AM_PALETTE[am_idx % len(AM_PALETTE)]
        total_rows = sum(len(staff) for _, staff in am_block["bcs"])
        card_h = 44 + 22 + total_rows * ROW_H + 16

        d.rounded_rectangle(
            [card_x0, y, card_x1, y + card_h],
            radius=14, fill=(252, 252, 253), outline=C_CARD_BORDER, width=1
        )

        inner_x = card_x0 + 20
        inner_right = card_x1 - 20
        row_y = y + 14

        d.rounded_rectangle([inner_x, row_y, inner_right, row_y + 26], radius=6, fill=bar_bg)
        d.text((inner_x + 10, row_y + 4), f"AM: {am_block['am']}", font=f_am, fill=bar_txt)
        row_y += 38

        inner_w = inner_right - inner_x
        col_name_w = inner_w - (COL_BC_W + COL_GAN_W + COL_TC_W + COL_PCT_W + COL_LTC_W)

        x_bc = inner_x
        x_name = x_bc + COL_BC_W
        x_gan = x_name + col_name_w
        x_tc = x_gan + COL_GAN_W
        x_pct = x_tc + COL_TC_W
        x_ltc = x_pct + COL_PCT_W

        d.text((x_bc, row_y), "Bưu cục", font=f_colhead, fill=(0, 0, 0))
        d.text((x_name, row_y), "Nhân viên", font=f_colhead, fill=(0, 0, 0))
        d.text((x_gan + COL_GAN_W - measure(d, "Gán giao", f_colhead)[0], row_y), "Gán giao", font=f_colhead, fill=(0, 0, 0))
        d.text((x_tc + COL_TC_W - measure(d, "Giao TC", f_colhead)[0], row_y), "Giao TC", font=f_colhead, fill=(0, 0, 0))
        d.text((x_pct + COL_PCT_W - measure(d, "%GTC", f_colhead)[0], row_y), "%GTC", font=f_colhead, fill=(0, 0, 0))
        d.text((x_ltc + COL_LTC_W - measure(d, "LTC", f_colhead)[0], row_y), "LTC", font=f_colhead, fill=(0, 0, 0))
        row_y += 22
        table_top = row_y

        global_row_idx = 0
        for bc_idx, (bc_name, staff) in enumerate(am_block["bcs"]):
            bc_block_top = row_y
            # Shading theo từng cụm Bưu Cục (thay vì so le từng dòng) - đỡ rối mắt hơn
            if bc_idx % 2 == 1:
                d.rectangle(
                    [inner_x - 4, row_y - 2, inner_right + 4, row_y + len(staff) * ROW_H - 4],
                    fill=C_ROW_ALT
                )

            for name, gan, tc, pct, ltc in staff:
                d.text((x_name, row_y), name, font=f_name, fill=C_TEXT)

                t = str(gan)
                d.text((x_gan + COL_GAN_W - measure(d, t, f_val)[0], row_y), t, font=f_val, fill=C_TEXT)
                t = str(tc)
                d.text((x_tc + COL_TC_W - measure(d, t, f_val)[0], row_y), t, font=f_val, fill=C_TEXT)

                pct_text = f"{pct:.2f}%"
                bg_c, txt_c = color_for(pct)
                pw, ph = measure(d, pct_text, f_val)
                badge_w = pw + 18
                badge_x1 = x_pct + COL_PCT_W
                badge_x0 = badge_x1 - badge_w
                d.rounded_rectangle([badge_x0, row_y - 3, badge_x1, row_y + ph + 3], radius=10, fill=bg_c)
                d.text((badge_x0 + 9, row_y - 1), pct_text, font=f_val, fill=txt_c)

                t = str(ltc)
                d.text((x_ltc + COL_LTC_W - measure(d, t, f_val)[0], row_y), t, font=f_val, fill=C_TEXT)

                row_y += ROW_H
                global_row_idx += 1

            bc_block_bottom = row_y
            bc_center_y = (bc_block_top + bc_block_bottom) / 2
            bw, bh = measure(d, bc_name, f_bc)
            if bw > COL_BC_W - 8:
                parts = bc_name.split(") ", 1)
                line1 = parts[0] + ")" if len(parts) > 1 else bc_name
                line2 = parts[1] if len(parts) > 1 else ""
                d.text((x_bc, bc_center_y - 14), line1, font=f_bc, fill=bar_txt)
                d.text((x_bc, bc_center_y + 2), line2, font=f_bc, fill=bar_txt)
            else:
                d.text((x_bc, bc_center_y - bh / 2), bc_name, font=f_bc, fill=bar_txt)

            if bc_block_bottom < table_top + total_rows * ROW_H:
                d.line([inner_x - 4, bc_block_bottom, inner_right + 4, bc_block_bottom], fill=C_GRIDLINE, width=1)

        d.line([x_name - 10, table_top - 4, x_name - 10, row_y - 2], fill=C_GRIDLINE, width=1)
        y += card_h + 16

    img = img.crop((0, 0, W, y + 10))
    img.save(out_path)
    print(f"✅ Đã tạo ảnh: {out_path}")
    return out_path


# ===== GỬI ẢNH VÀO GTALK (theo đúng pattern initiate-upload / put / complete-upload / send-message) =====

def upload_image_to_gtalk(image_path, channel_id, oa_token):
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

    resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload)
    if resp_init.status_code != 200:
        print(f"⚠️ GTalk initiate-upload HTTP error {resp_init.status_code}: {resp_init.text}")
        return None, None

    init_data = resp_init.json()
    if init_data.get("errorCode") != "success":
        print(f"⚠️ GTalk initiate-upload logic error: {init_data}")
        return None, None

    presigned_url = init_data["data"]["PresignedURL"]
    upload_id = init_data["data"]["UploadId"]

    resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"})
    if resp_put.status_code != 200:
        print(f"⚠️ GTalk put-file HTTP error {resp_put.status_code}: {resp_put.text}")
        return None, None

    resp_comp = requests.post(
        "https://mbff.ghn.vn/api/gtalk/complete-upload",
        json={"oaToken": oa_token, "UploadId": upload_id}
    )
    if resp_comp.status_code != 200:
        print(f"⚠️ GTalk complete-upload HTTP error {resp_comp.status_code}: {resp_comp.text}")
        return None, None

    comp_data = resp_comp.json()
    if comp_data.get("errorCode") != "success":
        print(f"⚠️ GTalk complete-upload logic error: {comp_data}")
        return None, None

    return comp_data["data"]["Id"], (width, height)


def send_report_to_gtalk(image_path, caption, channel_id=None, oa_token=None):
    channel_id = channel_id or GTALK_CHANNEL_ID
    oa_token = oa_token or GTALK_OA_TOKEN

    print("📡 Đang upload ảnh lên GTalk...")
    file_id, size = upload_image_to_gtalk(image_path, channel_id, oa_token)
    if not file_id:
        print("❌ Upload ảnh lên GTalk thất bại.")
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
    r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
    if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
        print("✅ Đã gửi ảnh vào kênh GTalk thành công!")
        return True
    else:
        print(f"❌ Gửi tin nhắn GTalk thất bại: {r_send.text}")
        return False


def main():
    now = datetime.now()
    report_date_str = now.strftime("%-d thg %-m, %Y") if os.name != "nt" else now.strftime("%d thg %m, %Y")
    updated_at_str = now.strftime("%H:%M · %d/%m/%Y")

    try:
        data = load_data_from_sheet()
        if not data:
            print("⚠️ Không có NVPTT nào ở mức Thấp trong sheet hôm nay — dùng data mẫu để test.")
            data = DATA_SAMPLE
    except Exception as e:
        print(f"⚠️ Lỗi đọc data từ Google Sheet: {e}")
        print("↪️ Dùng data mẫu để test layout.")
        data = DATA_SAMPLE

    print(f"📦 Sẽ tạo và gửi {len(data)} ảnh (mỗi AM 1 ảnh)...")

    for i, am_block in enumerate(data, start=1):
        am_name = am_block["am"]
        total_nv = sum(len(staff) for _, staff in am_block["bcs"])

        safe_name = "".join(c if c.isalnum() else "_" for c in am_name)
        out_path = os.path.join(BASE_DIR, f"nvptt_thap_{safe_name}.png")

        print(f"\n[{i}/{len(data)}] 🖼️ Đang tạo ảnh cho AM: {am_name} ({total_nv} NVPTT thấp)...")
        generate_report_image([am_block], report_date_str, updated_at_str, out_path)

        caption = (
            f"<b>CẬP NHẬT NVPTT MỨC THẤP</b>\n"
            f"AM: <b>{am_name}</b> · {total_nv} NVPTT dưới 80%\n"
            f"Ngày báo cáo {report_date_str} · Cập nhật lúc {updated_at_str}"
        )

        ok = send_report_to_gtalk(out_path, caption)
        if not ok:
            print(f"⚠️ Gửi ảnh AM '{am_name}' thất bại, tiếp tục AM kế tiếp...")

        try:
            os.remove(out_path)
        except Exception:
            pass

        if i < len(data):
            time.sleep(2)  # tránh gửi quá dồn dập, tránh bị rate-limit

    print("\n🎉 Đã xử lý xong toàn bộ AM.")


if __name__ == "__main__":
    main()