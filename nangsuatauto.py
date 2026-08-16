# -*- coding: utf-8 -*-
"""
EXPORT "3. GIAO HÀNG" TỪ DASHBOARD GHN -> UPLOAD VÀO GOOGLE SHEET.

Luồng thực hiện:
  1. Mở (hoặc bám vào) tab Chrome đang có sẵn trỏ tới dashboard GHN.
  2. Cuộn tới bảng "3. Giao Hàng".
  3. Bấm [+] để mở rộng cột Nhân Viên (hiện chi tiết theo từng nhân viên).
  4. Chuột phải -> Export data -> chọn CSV -> tick "Giữ định dạng giá trị" -> Xuất.
  5. Đọc file CSV vừa tải về thành DataFrame.
  6. Ghi đè (clear + update) vào đúng tab trong Google Sheet đích (xác định bằng gid).

Yêu cầu cài đặt (chạy 1 lần trên máy bạn):
    pip install playwright pandas gspread google-auth
    playwright install chromium

Yêu cầu cấu hình (xác thực Google Sheets — ưu tiên OAuth, fallback Service Account):
  - Nếu có `authorized_user.json` (OAuth) cùng thư mục script: script dùng luôn
    file này để xác thực. Tài khoản Google gắn với token đó PHẢI được share
    quyền "Người chỉnh sửa" (Editor) trên Google Sheet đích.
  - Nếu không có OAuth, script fallback sang `credentials.json` (Service
    Account). Phải SHARE Google Sheet đích cho đúng email trong field
    "client_email" của file đó.
  - Chrome cần đang chạy với cổng debug từ xa, ví dụ khởi động bằng:
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222
    và đã đăng nhập sẵn tài khoản GHN trên baocao.ghn.vn.
    Nếu không kết nối được Chrome debug, script sẽ tự mở 1 trình duyệt Edge
    riêng (profile lưu trong thư mục playwright_profile) và bạn đăng nhập thủ công.
"""

import os
import re
import sys
import time
import unicodedata
from datetime import datetime as _dt

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ========================= CẤU HÌNH =========================
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8ab"

# Google Sheet đích (link bạn cung cấp)
GOOGLE_SHEET_KEY = "1IUWdxN-VEC64OcciE09I_-3DHaTu39XhITblqMKt6Ww"
GOOGLE_SHEET_GID = 491121158          # dùng để xác định đúng tab, kể cả nếu đổi tên
TAB_GIAO_HANG = "giao hàng"           # dùng làm fallback / tên khi cần tạo mới
# ==============================================================


# ---------------------------------------------------------------
# GOOGLE SHEETS
# ---------------------------------------------------------------
def get_credentials_path():
    candidates = [
        os.path.join(SCRIPT_DIR, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def get_gspread_client(sheet_key=None):
    """Xác thực Google Sheets: ưu tiên OAuth (authorized_user.json) nếu có,
    fallback sang Service Account (credentials.json). Giống hệt cơ chế trong
    các script GHN khác của bạn đã chạy ổn định — dùng chung để đảm bảo cùng
    một identity/quyền truy cập với các sheet đang ghi được."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ]
    auth_user_candidates = [
        os.path.join(SCRIPT_DIR, 'authorized_user.json'),
        os.path.join(SCRIPT_DIR, 'credentials_oauth.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json',
    ]
    for auth_file in auth_user_candidates:
        if os.path.exists(auth_file):
            try:
                from google.oauth2.credentials import Credentials as UserCredentials
                creds = UserCredentials.from_authorized_user_file(auth_file, scopes=scopes)
                gc = gspread.authorize(creds)
                if sheet_key:
                    gc.open_by_key(sheet_key)
                print(f"✔️ Đã xác thực thành công qua {auth_file}", flush=True)
                return gc
            except Exception:
                pass

    json_path = get_credentials_path()
    if json_path and os.path.exists(json_path):
        try:
            credentials = ServiceAccountCredentials.from_service_account_file(json_path, scopes=scopes)
            gc = gspread.authorize(credentials)
            if sheet_key:
                gc.open_by_key(sheet_key)
            print(f"✔️ Đã xác thực thành công qua Service Account ({json_path})", flush=True)
            return gc
        except Exception:
            pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


def get_worksheet(sh, tab_name, gid=None):
    """Ưu tiên tìm đúng tab theo gid (chính xác tuyệt đối kể cả khi tab bị đổi tên).
    Nếu không tìm được theo gid thì fallback theo tên; nếu vẫn không có thì tạo mới."""
    if gid is not None:
        for ws in sh.worksheets():
            if str(ws.id) == str(gid):
                return ws
        print(f"⚠️ Không tìm thấy tab với gid={gid}, sẽ thử theo tên '{tab_name}'.", flush=True)

    try:
        return sh.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=30)
        print(f"✅ Đã tạo tab mới: '{tab_name}'", flush=True)
        return ws


def upload_to_sheet(df, sheet_key, tab_name, gid=None):
    try:
        gc = get_gspread_client(sheet_key)
        sh = gc.open_by_key(sheet_key)

        ws = get_worksheet(sh, tab_name, gid=gid)
        print(f"📤 Ghi dữ liệu vào tab '{ws.title}' ({len(df)} dòng)...", flush=True)
        ws.clear()
        data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
        try:
            ws.update(data_to_upload, value_input_option='USER_ENTERED')
        except TypeError:
            ws.update('A1', data_to_upload, value_input_option='USER_ENTERED')
        print(f"🎉 Tab '{ws.title}': {len(df)} dòng, {len(df.columns)} cột — cập nhật thành công!", flush=True)
        return True
    except Exception as e:
        print(f"❌ Lỗi khi ghi Google Sheet: [{type(e).__name__}] {e!r}", flush=True)
        import traceback
        traceback.print_exc()
        return False


# ---------------------------------------------------------------
# ĐỌC FILE CSV VỪA TẢI VỀ
# ---------------------------------------------------------------
def read_csv_smart(filepath):
    if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
        print(f"❌ File CSV rỗng hoặc không tồn tại: {filepath}", flush=True)
        return None

    combos = [
        ('utf-8-sig', ','), ('utf-8', ','), ('utf-16', '\t'),
        ('utf-16-le', '\t'), ('utf-8-sig', '\t'), ('utf-8-sig', ';'),
        ('utf-8', '\t'), ('latin1', ','), ('latin1', '\t'),
    ]
    for enc, sep in combos:
        try:
            tmp = pd.read_csv(filepath, encoding=enc, sep=sep, dtype=str, on_bad_lines='skip').fillna("")
            if len(tmp.columns) >= 1 and len(tmp) > 0:
                print(f"✅ Đọc CSV thành công: encoding={enc}, sep='{sep}', {len(tmp)} dòng, {len(tmp.columns)} cột.", flush=True)
                return tmp
        except Exception:
            continue

    try:
        tmp = pd.read_csv(filepath, encoding='utf-8-sig', sep=None, engine='python', dtype=str).fillna("")
        print(f"✅ Đọc CSV thành công (engine python): {len(tmp)} dòng, {len(tmp.columns)} cột.", flush=True)
        return tmp
    except Exception as ex:
        print(f"❌ Lỗi đọc CSV fallback: {ex}", flush=True)
        return None


# ---------------------------------------------------------------
# CHUẨN HÓA ĐỊNH DẠNG NGÀY / SỐ THẬP PHÂN KIỂU VIỆT NAM
# ---------------------------------------------------------------
def _convert_date_to_vn_format(val):
    s = str(val).strip()
    if not s:
        return s
    formats_to_try = ["%b %d, %Y", "%B %d, %Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]
    for fmt in formats_to_try:
        try:
            dt = _dt.strptime(s, fmt)
            return f"{dt.day} thg {dt.month}, {dt.year}"
        except ValueError:
            continue
    return s


def _convert_decimal_dot_to_comma(val):
    s = str(val).strip()
    if not s:
        return s
    m = re.fullmatch(r"(-?\d+)\.(\d+)(%?)", s)
    if m:
        return f"{m.group(1)},{m.group(2)}{m.group(3)}"
    return s


def clean_df_vn_format(df, date_columns=None, decimal_columns=None):
    df = df.copy()

    cols_date = date_columns
    if cols_date is None:
        cols_date = [c for c in df.columns if 'ngày' in c.lower() or 'date' in c.lower()]
        if not cols_date:
            for c in df.columns:
                sample = df[c].astype(str).head(5)
                if sample.str.match(r'^[A-Za-z]{3,9} \d{1,2},? \d{4}$').any():
                    cols_date.append(c)

    for c in cols_date:
        if c in df.columns:
            df[c] = df[c].apply(_convert_date_to_vn_format)
            print(f"📅 Đã chuyển định dạng ngày cột '{c}' sang kiểu Việt Nam.", flush=True)

    cols_decimal = decimal_columns
    if cols_decimal is None:
        cols_decimal = []
        for c in df.columns:
            if c in cols_date:
                continue
            sample = df[c].astype(str).head(10)
            if sample.str.match(r'^-?\d+\.\d+%?$').any():
                cols_decimal.append(c)

    for c in cols_decimal:
        df[c] = df[c].apply(_convert_decimal_dot_to_comma)
    if cols_decimal:
        print(f"🔢 Đã đổi dấu thập phân (. -> ,) cho các cột: {cols_decimal}", flush=True)

    return df


# ---------------------------------------------------------------
# CÁC HÀM HỖ TRỢ TÌM PHẦN TỬ TRÊN DASHBOARD (LOOKER STUDIO)
# ---------------------------------------------------------------
def debug_dump_near_anchor(report_frame, anchor_el, ghn_page, label, max_items=40):
    print(f"\n🕵️ [DEBUG] Đang quét DOM quanh section '{label}'...", flush=True)
    try:
        anchor_box = anchor_el.bounding_box()
    except Exception:
        anchor_box = None

    broad_selectors = [
        'div[class*="header-cell"]', 'div[class*="pivot"]', 'div[class*="cell"]',
        'div[class*="expand"]', 'div[class*="toggle"]', 'button', '[role="button"]', 'svg',
    ]

    seen = 0
    for sel in broad_selectors:
        try:
            els = report_frame.locator(sel).all()
        except Exception:
            continue
        for el in els:
            if seen >= max_items:
                break
            try:
                if not el.is_visible():
                    continue
                box = el.bounding_box()
                if not box:
                    continue
                if anchor_box and (box['y'] < anchor_box['y'] - 20 or box['y'] > anchor_box['y'] + 400):
                    continue
                cls = el.get_attribute("class") or ""
                text = ""
                try:
                    text = el.inner_text().strip()[:40]
                except Exception:
                    pass
                print(f"  [{sel}] class='{cls[:80]}' text='{text}' y={box['y']:.0f}", flush=True)
                seen += 1
            except Exception:
                continue
        if seen >= max_items:
            break
    if seen == 0:
        print("  (Không tìm thấy phần tử nào phù hợp trong vùng quét)", flush=True)
    print("🕵️ [DEBUG] Hết phần quét.\n", flush=True)


def find_section_anchor(report_frame, ghn_page, keywords):
    for kw in keywords:
        try:
            el = report_frame.locator(f'text="{kw}"').first
            if el.count() > 0:
                el.wait_for(state="attached", timeout=3000)
                print(f"✅ Tìm thấy anchor section: '{kw}'", flush=True)
                return el
        except Exception:
            continue
    return None


def get_elements_below_anchor(report_frame, anchor_el, selector, ghn_page):
    try:
        anchor_box = anchor_el.bounding_box()
        if not anchor_box:
            return []
        anchor_bottom = anchor_box['y'] + anchor_box['height']
    except Exception:
        return []

    all_els = report_frame.locator(selector).all()
    below = []
    for el in all_els:
        try:
            box = el.bounding_box()
            if box and box['y'] >= anchor_bottom - 10:
                below.append((box['y'], el))
        except Exception:
            continue

    below.sort(key=lambda x: x[0])
    return [el for _, el in below]


# ---------------------------------------------------------------
# HÀM CHÍNH: MỞ RỘNG [+] -> CHUỘT PHẢI -> EXPORT -> ĐỌC CSV
# ---------------------------------------------------------------
def export_table(ghn_page, report_frame, label, section_keywords, expand_details=True, expand_keywords=None):
    print(f"\n{'='*50}", flush=True)
    print(f"📋 BẮT ĐẦU XUẤT: {label}", flush=True)
    print(f"{'='*50}", flush=True)

    try:
        ghn_page.keyboard.press("Escape")
        ghn_page.wait_for_timeout(500)
        ghn_page.keyboard.press("Escape")
        ghn_page.wait_for_timeout(500)
        for ctx in [report_frame, ghn_page]:
            try:
                bd = ctx.locator('.cdk-overlay-backdrop').first
                if bd.count() > 0 and bd.is_visible():
                    bd.evaluate("el => el.click()")
            except Exception:
                pass
    except Exception:
        pass

    def get_looker_frame_helper(p):
        p.wait_for_timeout(1000)
        for f in p.frames:
            if f != p.main_frame and f.url and (
                "lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url
            ):
                return f
        for f in p.frames:
            if f != p.main_frame:
                return f
        return p.main_frame

    def clean_compare(s):
        s_norm = unicodedata.normalize('NFC', s).strip().lower()
        return s_norm.rstrip('.…')

    selectors = [
        '[role="menuitem"]', '.mat-mdc-menu-item', '.mat-menu-item',
        '.goog-menuitem', '.mat-mdc-menu-item-text',
    ]

    print(f"🔍 Tìm heading section '{label}'...", flush=True)
    anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

    if anchor:
        anchor.scroll_into_view_if_needed()
        ghn_page.wait_for_timeout(1500)
        print(f"📌 Đã scroll đến section '{label}'.", flush=True)
    else:
        print(f"⚠️ Không tìm thấy heading cho '{label}'.", flush=True)

    if expand_details:
        print("➕ Bấm [+] mở rộng bảng...", flush=True)
        hc = None
        matched_text = None

        def find_expand_target():
            candidates = []
            if anchor:
                candidates = get_elements_below_anchor(report_frame, anchor, 'div.header-cell', ghn_page)
            if not candidates:
                candidates = report_frame.locator('div.header-cell').all()

            # Thu thập text của TẤT CẢ candidate trước để debug + để chọn khớp
            # chính xác nhất (bảng phân cấp BuuCuc->NhanVien có thể khiến
            # header cell cha chứa cả text của header cell con lồng bên trong,
            # nên so khớp kiểu "chứa chuỗi con" dễ bắt nhầm cha thay vì con).
            cand_info = []  # list[(cand, text, text_norm)]
            for cand in candidates:
                try:
                    if not cand.is_visible():
                        continue
                    text = cand.inner_text().strip()
                except Exception:
                    continue
                text_norm = unicodedata.normalize('NFC', text).lower()
                cand_info.append((cand, text, text_norm))

            if cand_info:
                preview = " | ".join(f"'{t}'" for _, t, _ in cand_info[:15])
                print(f"🔎 [DEBUG] {len(cand_info)} header-cell hiện ra gần section: {preview}", flush=True)

            if expand_keywords:
                kw_norms = [unicodedata.normalize('NFC', kw).lower() for kw in expand_keywords]

                # 1) Ưu tiên khớp CHÍNH XÁC (text header == keyword) — tránh bắt
                #    nhầm ô cha có text lồng ghép của nhiều cột con.
                for cand, text, text_norm in cand_info:
                    if text_norm in kw_norms:
                        return cand, text

                # 2) Nếu không có khớp chính xác, chọn candidate chứa keyword
                #    NHƯNG có độ dài text ngắn nhất (ô càng cụ thể, text càng
                #    ngắn; ô cha lồng nhiều cột con sẽ dài hơn).
                substring_matches = []
                for cand, text, text_norm in cand_info:
                    for kw_norm in kw_norms:
                        if kw_norm in text_norm:
                            substring_matches.append((len(text_norm), cand, text))
                            break
                if substring_matches:
                    substring_matches.sort(key=lambda x: x[0])
                    _, best_cand, best_text = substring_matches[0]
                    return best_cand, best_text

            for cand, text, _ in cand_info:
                try:
                    btn = cand.locator('.expand-button, .interaction-button-wrapper, [class*="expand"]').first
                    if btn.count() > 0:
                        return cand, text
                except Exception:
                    continue

            return None, None

        for attempt in range(3):
            try:
                if report_frame.is_detached():
                    raise Exception("Frame is detached")

                hc, matched_text = find_expand_target()

                if hc:
                    print(f"🎯 Header cell khớp để bấm [+]: '{matched_text}'", flush=True)
                    hc.scroll_into_view_if_needed(timeout=5000)
                    ghn_page.wait_for_timeout(500)
                    hc.hover(timeout=5000)
                    ghn_page.wait_for_timeout(1500)

                    expand_btn = hc.locator('.expand-button, .interaction-button-wrapper, [class*="expand"]').first
                    if expand_btn.count() == 0:
                        expand_btn = hc
                    expand_btn.click(force=True, timeout=5000)
                    print(f"✅ Đã bấm [+] trên cột '{matched_text}'.", flush=True)
                    break
                else:
                    print(f"⚠️ Attempt {attempt+1}: chưa tìm thấy header cell khớp keyword mở rộng.", flush=True)
                    ghn_page.wait_for_timeout(1000)
            except Exception as e:
                print(f"⚠️ Thử {attempt+1}: Lỗi bấm [+] ({e}). Đang lấy lại frame...", flush=True)
                ghn_page.wait_for_timeout(2000)
                report_frame = get_looker_frame_helper(ghn_page)
                anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

        if hc is None:
            debug_dump_near_anchor(report_frame, anchor, ghn_page, label)

        print("⏳ Chờ bảng tải lại dữ liệu sau khi bấm [+] (8 giây)...", flush=True)
        ghn_page.wait_for_timeout(8000)

    print("🧹 Dismiss overlays trước khi click...", flush=True)
    try:
        ghn_page.keyboard.press("Escape")
        ghn_page.wait_for_timeout(500)
        bd = report_frame.locator('.cdk-overlay-backdrop').first
        if bd.count() > 0 and bd.is_visible():
            bd.evaluate("el => el.click()")
            ghn_page.wait_for_timeout(500)
    except Exception:
        pass
    print("🖱️ Click chuột phải để mở menu xuất...", flush=True)

    menu_kws = [
        "Export data", "Xuất dữ liệu", "Xuất biểu đồ...", "Export chart...",
        "Xuất biểu đồ", "Export chart", "Export", "Xuất",
    ]

    def is_menu_visible():
        def check_frame(frame_obj):
            for kw in menu_kws:
                kw_clean = clean_compare(kw)
                for sel in selectors:
                    try:
                        locator = frame_obj.locator(sel)
                        count = locator.count()
                        for idx in range(count):
                            el = locator.nth(idx)
                            if el.is_visible():
                                text = el.inner_text().strip()
                                lines = [clean_compare(line) for line in text.split('\n') if line.strip()]
                                if any(line == kw_clean for line in lines):
                                    return el, kw
                    except Exception:
                        continue
            return None, None

        res_el, res_kw = check_frame(report_frame)
        if res_el:
            return res_el, res_kw
        return check_frame(ghn_page)

    menu_el = None
    matched_kw = None
    click_targets = []

    try:
        if anchor:
            header_cells_below = get_elements_below_anchor(
                report_frame, anchor,
                'div.header-cell, .ng-ko-table, mat-table, table, .visual-container, .chart-container',
                ghn_page,
            )
            click_targets.extend(header_cells_below)
    except Exception as e:
        print(f"⚠️ Lỗi tìm phần tử dưới anchor: {e}", flush=True)

    if not click_targets:
        try:
            click_targets = report_frame.locator('div.header-cell').all()
        except Exception:
            pass

    print(f"📌 Tìm thấy {len(click_targets)} đối tượng tiềm năng để click chuột phải.", flush=True)

    if not click_targets and anchor:
        debug_dump_near_anchor(report_frame, anchor, ghn_page, label)

    for i, target in enumerate(click_targets):
        try:
            if report_frame.is_detached():
                report_frame = get_looker_frame_helper(ghn_page)

            try:
                bd = report_frame.locator('.cdk-overlay-backdrop').first
                if bd.count() > 0 and bd.is_visible():
                    bd.evaluate("el => el.click()")
                    ghn_page.wait_for_timeout(500)
                ghn_page.keyboard.press("Escape")
                ghn_page.wait_for_timeout(500)
            except Exception:
                pass

            target.scroll_into_view_if_needed(timeout=3000)
            target.hover()
            ghn_page.wait_for_timeout(500)

            print(f"👉 Thử {i+1}.1: Click chuột phải lên đối tượng...", flush=True)
            target.click(button="right", timeout=3000)
            ghn_page.wait_for_timeout(1500)

            menu_el, matched_kw = is_menu_visible()
            if menu_el:
                print("✅ Menu đã xuất hiện thành công bằng click chuột phải!", flush=True)
                break

            print(f"👉 Thử {i+1}.2: Dispatch event 'contextmenu'...", flush=True)
            target.dispatch_event("contextmenu")
            ghn_page.wait_for_timeout(1500)

            menu_el, matched_kw = is_menu_visible()
            if menu_el:
                print("✅ Menu đã xuất hiện thành công bằng dispatch_event!", flush=True)
                break
        except Exception as e:
            print(f"⚠️ Thử click chuột phải lên đối tượng {i+1} lỗi: {e}", flush=True)
            try:
                ghn_page.keyboard.press("Escape")
            except Exception:
                pass
            continue

    if not menu_el:
        print("❌ Chưa mở được menu, thử header cell bất kỳ...", flush=True)
        try:
            any_hc = report_frame.locator('div.header-cell').first
            if any_hc.count() > 0:
                any_hc.hover()
                any_hc.click(button="right", timeout=3000)
                ghn_page.wait_for_timeout(1500)
                menu_el, matched_kw = is_menu_visible()
        except Exception:
            pass

    if not menu_el:
        if anchor:
            debug_dump_near_anchor(report_frame, anchor, ghn_page, label)
        raise Exception(f"Không thể mở được menu ngữ cảnh để xuất '{label}'!")

    found_text = menu_el.inner_text().strip()
    print(f"✅ Đang chọn menu item: '{found_text}' (khớp: '{matched_kw}')", flush=True)

    if matched_kw in ["Export data", "Xuất dữ liệu"]:
        try:
            menu_el.hover()
            ghn_page.wait_for_timeout(500)
            menu_el.click(timeout=3000)
        except Exception:
            menu_el.evaluate("el => el.click()")
    else:
        sub_el = None
        for attempt in range(5):
            try:
                menu_el.hover()
                ghn_page.wait_for_timeout(800)
            except Exception:
                pass

            for sub_kw in ["Export data", "Xuất dữ liệu", "Export", "Xuất"]:
                sub_kw_clean = clean_compare(sub_kw)
                for ctx in [ghn_page, report_frame]:
                    for sel in selectors:
                        try:
                            locator = ctx.locator(sel)
                            count = locator.count()
                            for idx in range(count):
                                el = locator.nth(idx)
                                if el.is_visible():
                                    text = el.inner_text().strip()
                                    clean_t = clean_compare(text)
                                    if any(b in clean_t for b in ["chart", "biểu đồ", "explore"]):
                                        continue
                                    if clean_t in ["export data", "xuất dữ liệu", "export", "xuất"]:
                                        sub_el = el
                                        break
                            if sub_el:
                                break
                        except Exception:
                            continue
                    if sub_el:
                        break
                if sub_el:
                    break

            if sub_el:
                break

            if attempt == 2:
                try:
                    print("👉 Thử click nhẹ menu 'Export chart...' để kích hoạt submenu...", flush=True)
                    menu_el.click(force=True)
                except Exception:
                    pass

            ghn_page.wait_for_timeout(500)

        if sub_el:
            sub_text = sub_el.inner_text().strip()
            print(f"🎯 Click chọn từ submenu: '{sub_text}'", flush=True)
            try:
                sub_el.click(force=True)
            except Exception:
                sub_el.evaluate("el => el.click()")
        else:
            print("⚠️ Không tìm thấy submenu 'Export data', thử click trực tiếp menu 'Export chart...'...", flush=True)
            menu_el.click(force=True)

    ghn_page.wait_for_timeout(2000)

    dialog = None
    dialog_selector = 'mat-dialog-container, mat-mdc-dialog-container, [role="dialog"]'
    print("⏳ Chờ dialog 'Export data' xuất hiện...", flush=True)
    for _ in range(10):
        for ctx in [ghn_page, report_frame]:
            try:
                d = ctx.locator(dialog_selector).last
                if d.count() > 0 and d.is_visible():
                    dialog = d
                    break
            except Exception:
                continue
        if dialog:
            break
        ghn_page.wait_for_timeout(500)

    if not dialog:
        print("❌ Dialog 'Export data' chưa xuất hiện!", flush=True)
        raise Exception("Dialog 'Export data' không xuất hiện sau khi click menu xuất.")

    print("👉 Chọn định dạng CSV trong dialog...", flush=True)
    try:
        csv_opt = None
        opts = dialog.locator('mat-radio-button, mat-mdc-radio-button, label, span').all()
        for opt in opts:
            if opt.is_visible() and opt.inner_text().strip() == "CSV":
                csv_opt = opt
                break

        if csv_opt:
            try:
                csv_opt.click(timeout=3000)
                print("✅ Đã chọn định dạng CSV.", flush=True)
            except Exception:
                csv_opt.evaluate("el => el.click()")
            ghn_page.wait_for_timeout(500)
        else:
            print("ℹ️ Mặc định đã chọn CSV.", flush=True)
    except Exception as e:
        print(f"⚠️ Lỗi chọn CSV: {e}", flush=True)

    print("👉 Tích 'Keep value formatting' (Giữ định dạng giá trị)...", flush=True)
    try:
        keep_cb_selector = (
            '[formcontrolname="keepFormat"], mat-checkbox:has-text("Keep value formatting"), '
            'mat-checkbox:has-text("Giữ định dạng giá trị"), mat-mdc-checkbox:has-text("Keep value formatting"), '
            'mat-mdc-checkbox:has-text("Giữ định dạng giá trị"), label:has-text("Keep value formatting"), '
            'label:has-text("Giữ định dạng giá trị")'
        )
        keep_cb = dialog.locator(keep_cb_selector).last
        if keep_cb.count() > 0 and keep_cb.is_visible():
            is_checked = False
            try:
                aria_checked = keep_cb.get_attribute("aria-checked")
                if aria_checked is not None:
                    is_checked = (aria_checked == "true")
                else:
                    inner_input = keep_cb.locator('input')
                    if inner_input.count() > 0:
                        is_checked = inner_input.is_checked()
            except Exception:
                pass

            if not is_checked:
                try:
                    keep_cb.click(force=True, timeout=3000)
                    print("✅ Đã tích 'Giữ định dạng giá trị'.", flush=True)
                except Exception:
                    keep_cb.evaluate("el => el.click()")
                    print("✅ Đã tích 'Giữ định dạng giá trị' via JS.", flush=True)
            else:
                print("ℹ️ Checkbox đã tích sẵn.", flush=True)
        else:
            print("⚠️ Không tìm thấy checkbox 'Giữ định dạng giá trị' trong dialog, bỏ qua...", flush=True)
    except Exception as e:
        print(f"⚠️ Không tick được checkbox: {e}", flush=True)

    print("👉 Bấm Xuất để tải file...", flush=True)
    confirm_btn = None
    if dialog:
        try:
            c = dialog.locator('button:has-text("Export"), button:has-text("Xuất")').last
            if c.count() > 0 and c.is_visible():
                confirm_btn = c
        except Exception:
            pass

    if not confirm_btn:
        confirm_btn_selector = (
            'mat-mdc-dialog-container button:has-text("Export"), mat-mdc-dialog-container button:has-text("Xuất"), '
            'mat-dialog-container button:has-text("Export"), mat-dialog-container button:has-text("Xuất"), '
            'button:has-text("Export"), button:has-text("Xuất")'
        )
        for ctx in [ghn_page, report_frame]:
            try:
                c = ctx.locator(confirm_btn_selector).last
                if c.count() > 0 and c.is_visible():
                    confirm_btn = c
                    break
            except Exception:
                continue

    if not confirm_btn:
        print("❌ Không tìm thấy nút Xuất nào trong dialog!", flush=True)
        raise Exception("Không tìm thấy nút Xuất trong dialog!")

    # Dùng đúng cơ chế page.expect_download() bình thường của Playwright —
    # giống các script GHN khác của bạn đang chạy ổn định. KHÔNG can thiệp CDP
    # "Browser.setDownloadBehavior" (lệnh đó ở cấp toàn trình duyệt, set sai sẽ
    # làm hỏng cả tải file thủ công của Chrome cho tới khi restart trình duyệt).
    #
    # Vẫn giữ thêm phần lắng nghe network request/response export bên dưới —
    # không phải để xác định download, mà để CHẨN ĐOÁN thêm nếu lần sau vẫn lỗi
    # (biết chắc export có thực sự chạy tới server hay không, và server trả gì).

    # Theo dõi network request để biết CHẮC CHẮN cú bấm có thực sự kích hoạt
    # export hay không, thay vì chỉ suy đoán qua việc dialog đóng hay chưa.
    # Lý do: dialog của Material/Angular có thể đóng lại nếu click "force=True"
    # vô tình trúng lớp backdrop phía sau nút (giống bấm ra ngoài / Esc) —
    # trường hợp đó dialog vẫn đóng nhưng KHÔNG có request export nào được gửi.
    captured_export_requests = []

    def _on_request(req):
        try:
            u = req.url.lower()
            if "export" in u or "download" in u:
                captured_export_requests.append(req.url)
        except Exception:
            pass

    # Bắt luôn RESPONSE của request export/download — request "bắn đi" không
    # đồng nghĩa với "thành công". Nếu server trả lỗi (quyền, quá tải, rate
    # limit...), Chrome sẽ không có gì để tải về, dù request rõ ràng đã chạy.
    def _on_response(res):
        try:
            u = res.url.lower()
            if "export" not in u and "download" not in u:
                return
            status = res.status
            headers = res.headers
            content_type = headers.get("content-type", "")
            content_disposition = headers.get("content-disposition", "")
            content_length = headers.get("content-length", "")
            print(f"📡 [DEBUG] Response export: status={status} content-type='{content_type}' "
                  f"content-disposition='{content_disposition}' content-length='{content_length}' "
                  f"url={res.url}", flush=True)
            # Nếu là JSON nhỏ (thường là báo lỗi), in luôn nội dung để biết lý do thất bại.
            if "json" in content_type.lower() or (not content_disposition and status != 200):
                try:
                    body_text = res.text()
                    print(f"📡 [DEBUG] Nội dung response (rút gọn): {body_text[:500]}", flush=True)
                except Exception:
                    pass
        except Exception as e:
            print(f"⚠️ Lỗi đọc response export (bỏ qua): {e}", flush=True)

    ghn_page.on("request", _on_request)
    ghn_page.on("response", _on_response)

    export_request_seen = False
    download = None
    try:
        with ghn_page.expect_download(timeout=120000) as dl_info:
            for click_attempt in range(1, 5):
                print(f"👉 Bấm Xuất để tải file (lần {click_attempt})...", flush=True)

                # Ưu tiên click BÌNH THƯỜNG (không force) trước — nếu có phần tử nào
                # đang che khuất nút (overlay/backdrop còn sót), Playwright sẽ báo lỗi
                # rõ ràng "... intercepts pointer events" thay vì âm thầm click trúng
                # chỗ khác như khi dùng force=True ngay từ đầu.
                try:
                    confirm_btn.click(timeout=4000)
                except Exception as e:
                    print(f"⚠️ Click bình thường lần {click_attempt} không được ({e}); thử click force...", flush=True)
                    try:
                        confirm_btn.click(force=True, timeout=4000)
                    except Exception as e2:
                        print(f"⚠️ Click force lần {click_attempt} cũng lỗi: {e2}", flush=True)
                        try:
                            confirm_btn.evaluate("el => el.click()")
                        except Exception as e3:
                            print(f"⚠️ Click JS lần {click_attempt} cũng lỗi: {e3}", flush=True)

                ghn_page.wait_for_timeout(3000)

                if captured_export_requests:
                    export_request_seen = True
                    print(f"🌐 Đã thấy {len(captured_export_requests)} network request export/download.", flush=True)
                    for u in captured_export_requests[:5]:
                        print(f"    - {u}", flush=True)

                dialog_still_open = False
                try:
                    dialog_still_open = dialog.is_visible()
                except Exception:
                    dialog_still_open = False

                if not dialog_still_open:
                    print(f"✅ Dialog đã đóng sau lần bấm {click_attempt}, đang chờ Playwright bắt sự kiện download...", flush=True)
                    break

                print(f"⚠️ Dialog vẫn còn mở sau lần bấm {click_attempt}, thử tìm lại nút và bấm lại...", flush=True)
                try:
                    new_btn = dialog.locator('button:has-text("Export"), button:has-text("Xuất")').last
                    if new_btn.count() > 0:
                        confirm_btn = new_btn
                except Exception:
                    pass
            else:
                print("⚠️ Đã bấm Xuất 4 lần nhưng dialog vẫn không đóng — vẫn tiếp tục chờ download phòng khi export đã thực sự được kích hoạt.", flush=True)

        download = dl_info.value
    finally:
        try:
            ghn_page.remove_listener("request", _on_request)
            ghn_page.remove_listener("response", _on_response)
        except Exception:
            pass

    if not export_request_seen and not captured_export_requests:
        print("⚠️ Không thấy request export/download nào trong network — có thể tên endpoint khác "
              "'export'/'download' nên bị lọc trượt (không nhất thiết là lỗi).", flush=True)

    fname = download.suggested_filename
    print(f"✅ File tải về: {fname}", flush=True)

    local_path = os.path.join(SCRIPT_DIR, fname)
    download.save_as(local_path)

    df = None
    if local_path and os.path.exists(local_path):
        size_bytes = os.path.getsize(local_path)
        print(f"📄 Đang đọc dữ liệu từ file: {os.path.basename(local_path)} ({size_bytes} bytes)...", flush=True)
        if local_path.endswith('.csv'):
            df = read_csv_smart(local_path)
        else:
            try:
                df = pd.read_excel(local_path, dtype=str).fillna("")
                print(f"✅ Đọc Excel thành công: {len(df)} dòng, {len(df.columns)} cột.", flush=True)
            except Exception as ex:
                print(f"❌ Lỗi đọc Excel: {ex}", flush=True)

    try:
        if local_path and os.path.exists(local_path):
            os.remove(local_path)
            print("🧹 Đã xóa file tạm.", flush=True)
    except Exception:
        pass

    return df


# ---------------------------------------------------------------
# MAIN JOB
# ---------------------------------------------------------------
def run_job():
    print(f"\n🚀 BẮT ĐẦU CHẠY LÚC: {time.strftime('%H:%M:%S')}", flush=True)
    print("🔌 Đang kết nối tới Chrome debug (127.0.0.1:9222)...", flush=True)

    with sync_playwright() as p:
        is_cdp = True
        browser = None
        context = None
        ghn_page = None

        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            print("✅ Đã kết nối với Chrome debug thành công!", flush=True)
            if len(browser.contexts) == 0:
                print("❌ Không tìm thấy Chrome context.", flush=True)
                browser.close()
                return
            context = browser.contexts[0]

            for pg in context.pages:
                if "baocao.ghn.vn" in pg.url:
                    ghn_page = pg
                    break

            if not ghn_page:
                print("ℹ️ Không tìm thấy tab baocao.ghn.vn, mở tab mới trên Chrome debug...", flush=True)
                ghn_page = context.new_page()
                ghn_page.goto(DASHBOARD_URL)
        except Exception as e:
            print(f"⚠️ Không kết nối được Chrome debug. Lỗi: {e}", flush=True)
            print("🌐 Tự động mở trình duyệt độc lập mới (Persistent Browser)...", flush=True)
            is_cdp = False
            profile_dir = os.path.join(SCRIPT_DIR, "playwright_profile")
            os.makedirs(profile_dir, exist_ok=True)
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                channel="msedge",
                headless=False,
                args=["--start-maximized"],
            )
            ghn_page = context.new_page()
            print(f"👉 Mở trang báo cáo: {DASHBOARD_URL}", flush=True)
            ghn_page.goto(DASHBOARD_URL)

        print(f"✅ Đã tìm thấy tab: '{ghn_page.title()}'", flush=True)

        context.set_default_timeout(180000)
        ghn_page.set_default_timeout(180000)

        # LƯU Ý: KHÔNG gọi CDP "Browser.setDownloadBehavior" ở đây. Đây là lệnh
        # ở cấp TOÀN BỘ TRÌNH DUYỆT (không riêng tab này) — set sai có thể làm
        # hỏng luôn cơ chế tải file mặc định của Chrome cho MỌI tab, kể cả khi
        # bạn tự bấm tải tay, và trạng thái hỏng đó tồn tại cho tới khi Chrome
        # được khởi động lại. Cứ để Playwright tự quản lý download bình thường
        # qua page.expect_download() bên dưới (giống các script khác của bạn
        # đang chạy ổn định).

        try:
            print("⏳ Đang kiểm tra trạng thái đăng nhập...", flush=True)
            ghn_page.wait_for_selector("iframe", timeout=10000)
        except Exception:
            print("🔑 Vui lòng đăng nhập tài khoản GHN trên cửa sổ trình duyệt vừa mở...", flush=True)
            try:
                ghn_page.wait_for_selector("iframe", timeout=120000)
                print("✅ Đăng nhập thành công!", flush=True)
            except Exception as err:
                print(f"❌ Chi tiết lỗi khi chờ đăng nhập: {err}", flush=True)
                print("❌ Hết thời gian chờ đăng nhập (3 phút). Vui lòng chạy lại script.", flush=True)
                if is_cdp:
                    browser.close()
                else:
                    context.close()
                return

        try:
            ghn_page.keyboard.press("Escape")
            ghn_page.wait_for_timeout(1000)

            def get_looker_frame():
                ghn_page.wait_for_timeout(1000)
                for f in ghn_page.frames:
                    if f != ghn_page.main_frame and (
                        "lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url
                    ):
                        return f
                for f in ghn_page.frames:
                    if f != ghn_page.main_frame:
                        return f
                return ghn_page.main_frame

            report_frame = get_looker_frame()
            print(f"ℹ️ Frame: {report_frame.url[:80] if report_frame.url else 'Dynamic Frame'}", flush=True)

            page_name = ""
            try:
                pn = report_frame.locator("span.pageName, .pageName").first
                pn.wait_for(state="attached", timeout=10000)
                page_name = pn.text_content(timeout=3000).strip()
                print(f"📊 Trang hiện tại: '{page_name}'", flush=True)
            except Exception as e:
                print(f"⚠️ Không đọc được tên trang: {e}", flush=True)

            if "MỤC LỤC" in page_name:
                print("📑 Chuyển sang trang báo cáo (Click Next Page)...", flush=True)
                try:
                    next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                    next_btn.wait_for(state="attached", timeout=15000)
                    next_btn.hover()
                    next_btn.click(force=True)
                    print("✅ Đã click Next Page button.", flush=True)
                except Exception as e:
                    print(f"❌ Lỗi click Next Page button: {e}", flush=True)

                found = False
                for attempt in range(30):
                    ghn_page.wait_for_timeout(1000)
                    for f in ghn_page.frames:
                        if f != ghn_page.main_frame:
                            try:
                                pn2 = f.locator("span.pageName, .pageName").first
                                if pn2.count() > 0:
                                    cur = pn2.text_content(timeout=500).strip()
                                    if cur and "MỤC LỤC" not in cur:
                                        print(f"📊 Đã chuyển sang: '{cur}' (giây {attempt+1})", flush=True)
                                        report_frame = f
                                        found = True
                                        break
                            except Exception:
                                continue
                    if found:
                        break

                if not found:
                    print("❌ Không chuyển được trang.", flush=True)
                    if is_cdp:
                        browser.close()
                    else:
                        context.close()
                    return

                print("⏳ Chờ trang ổn định (5 giây)...", flush=True)
                ghn_page.wait_for_timeout(5000)
            else:
                print("ℹ️ Đã ở sẵn trên trang báo cáo.", flush=True)
                ghn_page.wait_for_timeout(2000)

            # ══════════════════════════════════════════
            # 3. GIAO HÀNG
            # ══════════════════════════════════════════
            df_giao = None
            try:
                df_giao = export_table(
                    ghn_page, report_frame,
                    label="3. GIAO HÀNG",
                    section_keywords=["3. Giao Hàng", "Giao Hàng", "3. GIAO HÀNG"],
                    expand_details=True,
                    expand_keywords=["NhanVien", "Nhân Viên", "TenNV", "Tên NV"],
                )
            except Exception as e:
                print(f"❌ Lỗi khi xuất '3. Giao Hàng': {e}", flush=True)

            if df_giao is not None:
                df_giao = clean_df_vn_format(df_giao)
                upload_to_sheet(df_giao, GOOGLE_SHEET_KEY, TAB_GIAO_HANG, gid=GOOGLE_SHEET_GID)
            else:
                print("❌ Không xuất được bảng '3. Giao Hàng'.", flush=True)

            print(f"\n🏁 HOÀN THÀNH LÚC: {time.strftime('%H:%M:%S')}", flush=True)

        except Exception as e:
            print(f"❌ Lỗi tổng quát: {e}", flush=True)
        finally:
            if is_cdp:
                if browser:
                    browser.close()
            else:
                if context:
                    context.close()


if __name__ == "__main__":
    run_job()