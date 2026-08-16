"""
build_gtc_ca1_push.py
----------------------
Push GTC TTS Ca 1: đọc sheet "NTB" (100% đơn chưa gán giao ca 1 TTS),
đếm số lượng theo AM, tách sheet riêng cho từng AM, render ảnh pivot
màu pastel (mốc đầu ngày = baseline, các lần bấm sau so sánh tăng/giảm
so với baseline), rồi gửi ảnh + caption lên group GTalk.

Cách chạy:
    python build_gtc_ca1_push.py              -> chạy bình thường
    python build_gtc_ca1_push.py --clear       -> xoá mốc baseline hôm nay (bấm lại từ đầu)
"""

import sys, io, os, json, unicodedata
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
import requests

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# ================== CẤU HÌNH ==================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SNAPSHOT_FILE = os.path.join(BASE_DIR, 'snapshot_gtc_ca1.json')

SHEET_KEY = '1vZCDHI2yW_LuprC4ShXvbqWRVFm_hNligdvrwauKw6A'
SOURCE_SHEET_NAME = 'NTB'          # tab nguồn chứa toàn bộ đơn ca 1 TTS
AM_COLUMN_NAME = 'AM'              # tên cột chứa AM trong header
STATUS_COLUMN_NAME = 'Trạng thái'  # cột H: xác định đơn đã/chưa gán
CHUA_GAN_VALUE = 'Chưa có chuyến đi trong ngày'  # giá trị = chưa gán giao

GTALK_TOKEN = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL = "2077030639771176960"

CAPTION_MESSAGE = "Anh/Chị AM gán các đơn ca 1 hàng TTS để kéo chỉ số GTC ca 1 TTS lên nhé"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_gspread_client(sheet_key=SHEET_KEY):
    if os.path.exists(JSON_FILE):
        try:
            creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
            gc = gspread.authorize(creds)
            if sheet_key:
                gc.open_by_key(sheet_key)
            return gc
        except Exception as e:
            print(f"⚠️ Service account ({JSON_FILE}) không có quyền truy cập: {e}. Đang chuyển sang authorized_user.json...")
            
    auth_user_file = os.path.join(BASE_DIR, 'authorized_user.json')
    if os.path.exists(auth_user_file):
        from google.oauth2.credentials import Credentials as UserCredentials
        creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=scopes)
        gc = gspread.authorize(creds)
        if sheet_key:
            gc.open_by_key(sheet_key)
        return gc
    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")

# Bảng màu pastel để tô các dòng AM luân phiên
ROW_PASTELS = ["#FFF8E1", "#E8F5E9", "#E3F2FD", "#FCE4EC", "#F3E5F5", "#E0F7FA", "#FFF3E0"]


def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())


# ================== BƯỚC 1: ĐỌC DỮ LIỆU + ĐẾM THEO AM ==================
def read_source_data(sh):
    print(f"📖 Đang đọc dữ liệu từ tab '{SOURCE_SHEET_NAME}'...")
    ws = sh.worksheet(SOURCE_SHEET_NAME)
    data = ws.get_all_values()
    if len(data) < 2:
        print(f"❌ Tab '{SOURCE_SHEET_NAME}' không có dữ liệu.")
        sys.exit(1)

    header = data[0]
    try:
        am_idx = header.index(AM_COLUMN_NAME)
    except ValueError:
        print(f"❌ Không tìm thấy cột '{AM_COLUMN_NAME}' trong header của '{SOURCE_SHEET_NAME}'.")
        sys.exit(1)

    try:
        status_idx = header.index(STATUS_COLUMN_NAME)
    except ValueError:
        print(f"❌ Không tìm thấy cột '{STATUS_COLUMN_NAME}' trong header của '{SOURCE_SHEET_NAME}'.")
        sys.exit(1)

    rows = data[1:]
    return header, rows, am_idx, status_idx


STATUS_DA_GAN_VALUE = 'Đang có chuyến đi trong ngày'  # giá trị = đã gán giao


def count_by_am(rows, am_idx, status_idx):
    """
    Với mỗi AM, đếm:
      - chua_gan: Trạng thái = 'Chưa có chuyến đi trong ngày'
      - da_gan:   Trạng thái = 'Đang có chuyến đi trong ngày'
      - tong:     chua_gan + da_gan (không tính #N/A vì đó là đơn đã xử lý xong)
    groups: chỉ chứa các dòng CHƯA GÁN (để tách sheet cho AM xử lý).
    """
    stats = {}
    groups = {}
    for row in rows:
        if len(row) <= max(am_idx, status_idx):
            continue
        status = normalize_str(row[status_idx])
        am = normalize_str(row[am_idx])
        if not am:
            continue

        if status == CHUA_GAN_VALUE:
            s = stats.setdefault(am, {"tong": 0, "da_gan": 0, "chua_gan": 0})
            s["chua_gan"] += 1
            s["tong"] += 1
            groups.setdefault(am, []).append(row)
        elif status == STATUS_DA_GAN_VALUE:
            s = stats.setdefault(am, {"tong": 0, "da_gan": 0, "chua_gan": 0})
            s["da_gan"] += 1
            s["tong"] += 1
        # #N/A (đã xử lý xong) -> không tính vào tổng

    return stats, groups


# ================== BƯỚC 2: TÁCH SHEET THEO AM ==================
def split_sheets_by_am(sh, header, groups):
    print("📂 Đang tách sheet theo AM...")
    all_ws = {ws.title: ws for ws in sh.worksheets()}
    for am_name, am_rows in groups.items():
        sheet_name = am_name[:31]
        if sheet_name in all_ws:
            ws_am = all_ws[sheet_name]
            ws_am.clear()
        else:
            ws_am = sh.add_worksheet(title=sheet_name, rows=str(max(100, len(am_rows) + 50)), cols=str(len(header)))
        ws_am.update([header] + am_rows)
        ws_am.format(f"A1:{gspread.utils.rowcol_to_a1(1, len(header))}".split("1")[0] + "1", {
            "backgroundColor": {"red": 0.96, "green": 0.65, "blue": 0.51},
            "textFormat": {"bold": True}
        })
    print(f"✔️ Đã tách xong {len(groups)} sheet AM.")


# ================== BƯỚC 3: SNAPSHOT MỐC TRONG NGÀY ==================
def load_or_init_snapshot(current_chua_gan):
    """current_chua_gan: dict {am: so_luong_chua_gan} - dùng để so sánh mốc."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    state = {"date": "", "baseline": {}, "baseline_time": "", "last_time": ""}

    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except Exception:
            pass

    now_time = datetime.now().strftime("%H:%M")
    is_first_run_today = (state.get("date") != today_str)

    if is_first_run_today:
        state = {
            "date": today_str,
            "baseline": current_chua_gan,
            "baseline_time": now_time,
            "last_time": now_time
        }
    else:
        state["last_time"] = now_time

    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    return state, is_first_run_today


# ================== BƯỚC 4: RENDER ẢNH PIVOT PASTEL ==================
def build_rows_html(sorted_ams, stats, baseline_chua_gan, is_first_run):
    rows_html = ""
    for i, am in enumerate(sorted_ams):
        s = stats[am]
        tong, da_gan, chua_gan = s["tong"], s["da_gan"], s["chua_gan"]
        bg = ROW_PASTELS[i % len(ROW_PASTELS)]

        if is_first_run:
            delta_html = "<span class='delta-none'>Mốc đầu ngày</span>"
        else:
            prev = baseline_chua_gan.get(am, 0)
            diff = chua_gan - prev
            if diff == 0:
                delta_html = "<span class='delta-none'>Không đổi</span>"
            elif diff < 0:
                delta_html = f"<span class='delta-badge delta-green'>▼ {abs(diff):,} đơn</span>"
            else:
                delta_html = f"<span class='delta-badge delta-red'>▲ {diff:,} đơn</span>"

        rows_html += f"""
        <tr style="background-color:{bg} !important;">
            <td class="am-name">{am}</td>
            <td class="count-val tong-val">{tong:,}</td>
            <td class="count-val da-gan-val">{da_gan:,}</td>
            <td class="count-val chua-gan-val">{chua_gan:,}</td>
            <td class="delta-col">{delta_html}</td>
        </tr>"""
    return rows_html


def render_and_get_png(sorted_ams, stats, baseline_chua_gan, is_first_run, state):
    total_tong = sum(s["tong"] for s in stats.values())
    total_da_gan = sum(s["da_gan"] for s in stats.values())
    total_chua_gan = sum(s["chua_gan"] for s in stats.values())
    mốc_label = f"Mốc {state['last_time']}" + ("" if is_first_run else f" (so với mốc đầu ngày {state['baseline_time']})")

    rows_html = build_rows_html(sorted_ams, stats, baseline_chua_gan, is_first_run)

    html_content = f"""
<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@600;700;800&family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    body {{ font-family:'Inter',sans-serif; background:linear-gradient(135deg,#fdf6f0,#f0f4f8); margin:0; padding:36px; }}
    #box {{ background:#fff; border-radius:20px; padding:32px; box-shadow:0 20px 25px -5px rgba(0,0,0,.1); width:740px; }}
    h2 {{ font-family:'Plus Jakarta Sans',sans-serif; font-size:24px; margin:0 0 4px 0;
          background:linear-gradient(90deg,#f97316,#ea580c); -webkit-background-clip:text; -webkit-text-fill-color:transparent; }}
    .sub {{ font-size:13px; color:#64748b; margin-bottom:18px; }}
    table {{ width:100%; border-collapse:separate; border-spacing:0; border-radius:12px; overflow:hidden; border:1px solid #eee; }}
    th {{ background:linear-gradient(180deg,#fb923c,#ea580c); color:#fff; font-size:12px; text-transform:uppercase;
          padding:10px 6px; text-align:center; }}
    th.left {{ text-align:left; }}
    td {{ padding:10px 6px; font-size:15px; border-bottom:1px solid #f1f5f9; text-align:center; }}
    td.am-name {{ font-weight:800; color:#1e293b; text-align:left; }}
    td.count-val {{ font-weight:800; font-size:17px; }}
    td.tong-val {{ color:#475569; }}
    td.da-gan-val {{ color:#16a34a; }}
    td.chua-gan-val {{ color:#c2410c; font-size:19px; }}
    .delta-badge {{ display:inline-block; padding:3px 10px; border-radius:8px; font-weight:700; font-size:12px; }}
    .delta-red {{ background:#fef2f2; color:#dc2626; border:1px solid #fecaca; }}
    .delta-green {{ background:#f0fdf4; color:#16a34a; border:1px solid #bbf7d0; }}
    .delta-none {{ color:#94a3b8; font-size:12px; font-weight:600; }}
    tr.total-row td {{ background:#fef3c7 !important; font-weight:800; font-size:17px; border-top:2px solid #f59e0b; }}
</style></head>
<body>
<div id="box">
    <h2>GTC TTS CA 1 — TÌNH HÌNH GÁN GIAO</h2>
    <div class="sub">{mốc_label}</div>
    <table>
        <thead><tr>
            <th class="left">AM</th>
            <th>Tổng đơn</th>
            <th>Đã gán</th>
            <th>Chưa gán</th>
            <th>So với đầu ngày</th>
        </tr></thead>
        <tbody>
            {rows_html}
            <tr class="total-row">
                <td style="text-align:left;">TỔNG CỘNG</td>
                <td>{total_tong:,}</td>
                <td>{total_da_gan:,}</td>
                <td>{total_chua_gan:,}</td>
                <td></td>
            </tr>
        </tbody>
    </table>
</div>
</body></html>
"""

    temp_html = os.path.join(BASE_DIR, "temp_gtc_ca1.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    out_png = os.path.join(BASE_DIR, "gtc_ca1_push.png")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({"width": 860, "height": 900})
        page.goto(f"file:///{temp_html.replace(chr(92), '/')}")
        page.wait_for_timeout(600)
        page.locator("#box").screenshot(path=out_png)
        browser.close()

    try:
        os.remove(temp_html)
    except Exception:
        pass

    return out_png, total_tong, total_da_gan, total_chua_gan


# ================== BƯỚC 5: GỬI GTALK ==================
def send_to_gtalk(image_path, caption):
    print("📡 Đang gửi ảnh lên GTalk...")
    file_name = os.path.basename(image_path)
    file_size = os.path.getsize(image_path)
    with open(image_path, 'rb') as f:
        file_bytes = f.read()

    init_payload = {
        "ChannelId": GTALK_CHANNEL,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 860, "height": 900}),
        "oaToken": GTALK_TOKEN
    }
    resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload)
    if resp_init.status_code != 200:
        print(f"❌ Lỗi initiate-upload: {resp_init.text}")
        return
    init_data = resp_init.json()
    if init_data.get("errorCode") != "success":
        print(f"❌ initiate-upload trả lỗi: {init_data}")
        return

    presigned_url = init_data["data"]["PresignedURL"]
    upload_id = init_data["data"]["UploadId"]

    resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"})
    if resp_put.status_code != 200:
        print(f"❌ Lỗi upload file: {resp_put.status_code}")
        return

    resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload",
                               json={"oaToken": GTALK_TOKEN, "UploadId": upload_id})
    if resp_comp.status_code != 200:
        print(f"❌ Lỗi complete-upload: {resp_comp.text}")
        return
    comp_data = resp_comp.json()
    if comp_data.get("errorCode") != "success":
        print(f"❌ complete-upload trả lỗi: {comp_data}")
        return

    file_id = comp_data["data"]["Id"]
    send_payload = {
        "channelId": GTALK_CHANNEL,
        "clientMsgId": str(int(os.path.getmtime(image_path) * 1000)),
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": caption,
                "items": [{"image": {"fileId": file_id, "width": 860, "height": 900}}]
            }
        },
        "oaToken": GTALK_TOKEN
    }
    r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
    if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
        print("✅ Đã gửi báo cáo GTC Ca 1 sang GTalk group thành công!")
    else:
        print(f"❌ Gửi tin nhắn GTalk lỗi: {r_send.text}")


# ================== MAIN ==================
def main():
    if "--clear" in sys.argv and os.path.exists(SNAPSHOT_FILE):
        os.remove(SNAPSHOT_FILE)
        print("🧹 Đã xoá mốc baseline hôm nay. Lần bấm tới sẽ là mốc đầu ngày mới.")

    gc_client = get_gspread_client(SHEET_KEY)
    sh = gc_client.open_by_key(SHEET_KEY)

    header, rows, am_idx, status_idx = read_source_data(sh)
    stats, groups = count_by_am(rows, am_idx, status_idx)

    if not stats:
        print("❌ Không có AM nào có đơn ca 1 TTS. Dừng.")
        return

    split_sheets_by_am(sh, header, groups)

    current_chua_gan = {am: s["chua_gan"] for am, s in stats.items()}
    state, is_first_run = load_or_init_snapshot(current_chua_gan)
    baseline_chua_gan = state.get("baseline", {})

    sorted_ams = sorted(stats.keys(), key=lambda x: stats[x]["chua_gan"], reverse=True)

    png_path, total_tong, total_da_gan, total_chua_gan = render_and_get_png(
        sorted_ams, stats, baseline_chua_gan, is_first_run, state)

    caption = f"📊 <b>GTC TTS CA 1 — Tình hình gán giao</b>\n"
    caption += f"⏱️ Mốc: {state['last_time']}" + ("" if is_first_run else f" (so với mốc đầu ngày {state['baseline_time']})") + "\n"
    caption += f"Tổng đơn: <b>{total_tong:,}</b> | Đã gán: <b>{total_da_gan:,}</b> | Chưa gán: <b>{total_chua_gan:,}</b>\n\n"
    caption += f"{CAPTION_MESSAGE}\n"
    caption += f'🔗 <a href="https://docs.google.com/spreadsheets/d/1vZCDHI2yW_LuprC4ShXvbqWRVFm_hNligdvrwauKw6A/edit?gid=1077331166#gid=1077331166"><b>Xem chi tiết tại đây</b></a>'

    send_to_gtalk(png_path, caption)
    print("🎉 HOÀN THÀNH!")


if __name__ == "__main__":
    main()