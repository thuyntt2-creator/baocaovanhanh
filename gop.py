# -*- coding: utf-8 -*-
"""
PIPELINE GỘP: Tải/replace Lấy Hàng-Giao Hàng -> Đọc BaoCao -> Tạo ảnh theo AM -> Gửi GTalk.
Gộp từ 2 file: export_lay_giao_hang.py (Playwright export) + send_nvptt_report_to_gtalk.py.
"""

# ======================================================================
# PHẦN 1: EXPORT LẤY HÀNG / GIAO HÀNG (từ export_lay_giao_hang.py)
# ======================================================================
import os
import sys
import time
import unicodedata
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

# Force stdout encoding to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Get absolute path of the script directory to handle Task Scheduler working directory issues
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ========== CONFIGURATION ==========
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8ab"
GOOGLE_SHEET_KEY = "1-p9VUXndK_7BoiT-a81UfTCbUi953XNmVBoXaTGis_c"

TAB_LAY_HANG = "lấy hàng"
TAB_GIAO_HANG = "giao hàng"
# ===================================


def get_credentials_path():
    candidates = [
        os.path.join(SCRIPT_DIR, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json"
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def get_or_create_worksheet(sh, tab_name):
    try:
        ws = sh.worksheet(tab_name)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=30)
        print(f"✅ Đã tạo tab mới: '{tab_name}'")
        return ws


def upload_to_sheet_tab(df, sheet_key, tab_name):
    json_path = get_credentials_path()
    if not json_path:
        print("❌ Không tìm thấy credentials.json.")
        return False
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = Credentials.from_service_account_file(json_path, scopes=scopes)
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(sheet_key)

        ws = get_or_create_worksheet(sh, tab_name)
        print(f"📤 Ghi dữ liệu vào tab '{tab_name}' ({len(df)} dòng)...", flush=True)
        ws.clear()
        data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
        try:
            ws.update(data_to_upload, value_input_option='USER_ENTERED')
        except TypeError:
            ws.update('A1', data_to_upload, value_input_option='USER_ENTERED')
        print(f"🎉 Tab '{tab_name}': {len(df)} dòng, {len(df.columns)} cột — cập nhật thành công!", flush=True)
        return True
    except Exception as e:
        print(f"❌ Lỗi khi ghi tab '{tab_name}': [{type(e).__name__}] {e!r}")
        try:
            import traceback
            traceback.print_exc()
        except Exception:
            pass
        return False


def read_csv_smart(filepath):
    if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
        print(f"❌ File CSV rỗng hoặc không tồn tại: {filepath}")
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
                print(f"✅ Đọc CSV thành công: encoding={enc}, sep='{sep}', {len(tmp)} dòng, {len(tmp.columns)} cột.")
                return tmp
        except Exception:
            continue

    try:
        tmp = pd.read_csv(filepath, encoding='utf-8-sig', sep=None, engine='python', dtype=str).fillna("")
        print(f"✅ Đọc CSV thành công (engine python): {len(tmp)} dòng, {len(tmp.columns)} cột.")
        return tmp
    except Exception as ex:
        print(f"❌ Lỗi đọc CSV fallback: {ex}")
        return None


def debug_dump_near_anchor(report_frame, anchor_el, ghn_page, label, max_items=40):
    print(f"\n🕵️ [DEBUG] Đang quét DOM quanh section '{label}' để tìm nút [+] / header cell...", flush=True)
    try:
        anchor_box = anchor_el.bounding_box()
    except Exception:
        anchor_box = None

    broad_selectors = [
        'div[class*="header-cell"]',
        'div[class*="pivot"]',
        'div[class*="cell"]',
        'div[class*="expand"]',
        'div[class*="toggle"]',
        'button',
        '[role="button"]',
        'svg',
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
                if anchor_box:
                    if box['y'] < anchor_box['y'] - 20 or box['y'] > anchor_box['y'] + 400:
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


def export_table(ghn_page, report_frame, label, section_keywords, filter_btn_index_fallback=0, expand_details=True, expand_keywords=None):
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
        '[role="menuitem"]',
        '.mat-mdc-menu-item',
        '.mat-menu-item',
        '.goog-menuitem',
        '.mat-mdc-menu-item-text'
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
                candidates = get_elements_below_anchor(
                    report_frame, anchor, 'div.header-cell', ghn_page
                )
            if not candidates:
                candidates = report_frame.locator('div.header-cell').all()

            if expand_keywords:
                for cand in candidates:
                    try:
                        if not cand.is_visible():
                            continue
                        text = cand.inner_text().strip()
                    except Exception:
                        continue
                    text_norm = unicodedata.normalize('NFC', text).lower()
                    for kw in expand_keywords:
                        kw_norm = unicodedata.normalize('NFC', kw).lower()
                        if kw_norm in text_norm:
                            return cand, text

            for cand in candidates:
                try:
                    if not cand.is_visible():
                        continue
                    btn = cand.locator('.expand-button, .interaction-button-wrapper, [class*="expand"]').first
                    if btn.count() > 0:
                        text = cand.inner_text().strip()
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
        "Xuất biểu đồ", "Export chart", "Export", "Xuất"
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
                ghn_page
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

    print("👉 Tích 'Keep value formatting'...", flush=True)
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
                    print("✅ Đã tích 'Keep value formatting'.", flush=True)
                except Exception:
                    keep_cb.evaluate("el => el.click()")
                    print("✅ Đã tích 'Keep value formatting' via JS.", flush=True)
            else:
                print("ℹ️ Checkbox đã tích sẵn.", flush=True)
        else:
            print("⚠️ Không tìm thấy checkbox 'Keep value formatting' trong dialog, bỏ qua...", flush=True)
    except Exception as e:
        print(f"⚠️ Không tick được checkbox: {e}", flush=True)

    print("👉 Bấm Export để tải file...", flush=True)
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
        print("❌ Không tìm thấy nút Export nào trong dialog!", flush=True)
        raise Exception("Không tìm thấy nút Export trong dialog!")

    with ghn_page.expect_download(timeout=120000) as dl_info:
        try:
            confirm_btn.click(force=True, timeout=5000)
        except Exception:
            confirm_btn.evaluate("el => el.click()")

    download = dl_info.value
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


import re
from datetime import datetime as _dt


def _convert_date_to_vn_format(val):
    s = str(val).strip()
    if not s:
        return s
    formats_to_try = [
        "%b %d, %Y",   # Jul 2, 2026
        "%B %d, %Y",   # July 2, 2026
        "%Y-%m-%d",    # 2026-07-02
        "%d/%m/%Y",    # 02/07/2026
        "%m/%d/%Y",    # 07/02/2026
    ]
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


def run_job():
    print(f"\n🚀 BẮT ĐẦU CHẠY AUTO DOWNLOAD LÚC: {time.strftime('%H:%M:%S')}", flush=True)
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
                args=["--start-maximized"]
            )
            ghn_page = context.new_page()
            print(f"👉 Mở trang báo cáo: {DASHBOARD_URL}", flush=True)
            ghn_page.goto(DASHBOARD_URL)

        print(f"✅ Đã tìm thấy tab: '{ghn_page.title()}'", flush=True)

        context.set_default_timeout(180000)
        ghn_page.set_default_timeout(180000)

        try:
            print("⏳ Đang kiểm tra trạng thái đăng nhập...", flush=True)
            ghn_page.wait_for_selector("iframe", timeout=10000)
        except Exception:
            print("🔑 Vui lòng thực hiện đăng nhập tài khoản GHN trên cửa sổ trình duyệt vừa mở...", flush=True)
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

            # 2. LẤY HÀNG
            df_lay = None
            try:
                df_lay = export_table(
                    ghn_page, report_frame,
                    label="2. LẤY HÀNG",
                    section_keywords=[
                        "2. Lấy Hàng",
                        "Lấy Hàng",
                        "2. LẤY HÀNG",
                    ],
                    expand_details=True,
                    expand_keywords=["TenNV", "Tên NV", "Nhân Viên", "NhanVien"]
                )
            except Exception as e:
                print(f"❌ Lỗi khi xuất '2. Lấy Hàng': {e}", flush=True)
            if df_lay is not None:
                df_lay = clean_df_vn_format(df_lay)
                upload_to_sheet_tab(df_lay, GOOGLE_SHEET_KEY, TAB_LAY_HANG)
            else:
                print("❌ Không xuất được bảng '2. Lấy Hàng'.", flush=True)

            # 3. GIAO HÀNG
            df_giao = None
            try:
                df_giao = export_table(
                    ghn_page, report_frame,
                    label="3. GIAO HÀNG",
                    section_keywords=[
                        "3. Giao Hàng",
                        "Giao Hàng",
                        "3. GIAO HÀNG",
                    ],
                    expand_details=True,
                    expand_keywords=["NhanVien", "Nhân Viên", "TenNV", "Tên NV"]
                )
            except Exception as e:
                print(f"❌ Lỗi khi xuất '3. Giao Hàng': {e}", flush=True)
            if df_giao is not None:
                df_giao = clean_df_vn_format(df_giao)
                upload_to_sheet_tab(df_giao, GOOGLE_SHEET_KEY, TAB_GIAO_HANG)
            else:
                print("❌ Không xuất được bảng '3. Giao Hàng'.", flush=True)

            print(f"\n🏁 HOÀN THÀNH TOÀN BỘ LÚC: {time.strftime('%H:%M:%S')}", flush=True)

        except Exception as e:
            print(f"❌ Lỗi tổng quát: {e}", flush=True)
        finally:
            if is_cdp:
                if browser:
                    browser.close()
            else:
                if context:
                    context.close()


# ======================================================================
# PHẦN 2: ĐỌC BAOCAO, TẠO ẢNH, GỬI GTALK (từ send_nvptt_report_to_gtalk.py)
# ======================================================================
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
    print("⚠️ Chưa cài python-dotenv — dùng token mặc định trong code.")

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


BASELINE_TAB_NAME = "NVPTT_Baseline_10h"


def _read_baocao_raw(sheet_key=BAOCAO_SHEET_KEY, tab_name=BAOCAO_TAB_NAME):
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

    i_manv = col_idx("Mã NV", "MaNV")
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

    rows_out = []
    for row in data_rows:
        if len(row) <= i_pct:
            continue

        bc = row[i_bc].strip() if i_bc < len(row) else ""
        am = row[i_am].strip() if i_am < len(row) else ""
        name = row[i_name].strip() if i_name < len(row) else ""
        if not bc or not am or not name:
            continue

        ma_nv = row[i_manv].strip() if (i_manv is not None and i_manv < len(row)) else ""
        if not ma_nv:
            ma_nv = f"{name}__{bc}"

        pct = parse_pct(row[i_pct]) if i_pct < len(row) else 0.0
        danh_gia = row[i_danhgia].strip() if (i_danhgia is not None and i_danhgia < len(row)) else ""
        gan = parse_num(row[i_gan]) if (i_gan is not None and i_gan < len(row)) else 0
        tc = parse_num(row[i_tc]) if (i_tc is not None and i_tc < len(row)) else 0
        ltc = parse_num(row[i_ltc]) if (i_ltc is not None and i_ltc < len(row)) else 0

        rows_out.append({
            "ma_nv": ma_nv, "am": am, "bc": bc, "name": name,
            "gan": gan, "tc": tc, "pct": pct, "ltc": ltc, "danh_gia": danh_gia
        })

    return rows_out, sh


def save_baseline_snapshot(rows, sh, today_str):
    try:
        ws = sh.worksheet(BASELINE_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=BASELINE_TAB_NAME, rows=len(rows) + 10, cols=4)

    ws.clear()
    data = [["MaNV", "NhanVien", "PctGTC", "NgayLuu"]]
    for r in rows:
        data.append([r["ma_nv"], r["name"], r["pct"], today_str])
    ws.update(data, value_input_option="USER_ENTERED")
    print(f"💾 Đã lưu baseline 10h ({len(rows)} NVPTT) vào tab '{BASELINE_TAB_NAME}'.")


def load_baseline_snapshot(sh, today_str):
    try:
        ws = sh.worksheet(BASELINE_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"⚠️ Chưa có tab '{BASELINE_TAB_NAME}' (có thể mốc 10h chưa chạy hôm nay).")
        return {}

    values = ws.get_all_values()
    if len(values) < 2:
        return {}

    header = [h.strip() for h in values[0]]
    try:
        i_manv = header.index("MaNV")
        i_pct = header.index("PctGTC")
        i_date = header.index("NgayLuu")
    except ValueError:
        return {}

    baseline = {}
    for row in values[1:]:
        if len(row) <= max(i_manv, i_pct, i_date):
            continue
        if row[i_date].strip() != today_str:
            continue
        try:
            baseline[row[i_manv]] = float(row[i_pct])
        except ValueError:
            continue

    if not baseline:
        print(f"⚠️ Không tìm thấy baseline của ngày {today_str} trong tab '{BASELINE_TAB_NAME}'.")
    return baseline


def load_data_from_sheet(sheet_key=BAOCAO_SHEET_KEY, tab_name=BAOCAO_TAB_NAME, threshold_pct=80.0,
                          is_baseline_run=False, today_str=None):
    if today_str is None:
        today_str = datetime.now().strftime("%Y-%m-%d")

    rows, sh = _read_baocao_raw(sheet_key, tab_name)

    baseline_map = {}
    if is_baseline_run:
        save_baseline_snapshot(rows, sh, today_str)
    else:
        baseline_map = load_baseline_snapshot(sh, today_str)

    grouped = {}
    for r in rows:
        delta = None
        if not is_baseline_run and r["ma_nv"] in baseline_map:
            delta = round(r["pct"] - baseline_map[r["ma_nv"]], 2)

        grouped.setdefault(r["bc"], {"am": r["am"], "staff": []})
        grouped[r["bc"]]["staff"].append(
            (r["name"], r["gan"], r["tc"], r["pct"], r["ltc"], delta)
        )

    result = []
    for bc_name, info in grouped.items():
        info["staff"].sort(key=lambda x: x[3])
        result.append({
            "bc": bc_name,
            "am": info["am"],
            "staff": info["staff"]
        })

    return result

# ===== CẤU HÌNH GTALK =====
# Lấy từ .env nếu có, fallback theo giá trị mặc định (giống pattern rot_lc_gtalk.py)
GTALK_OA_TOKEN = os.environ.get("NVPTT_GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_CHANNEL_ID = os.environ.get("NVPTT_THAP_GTALK_CHANNEL_ID") or "2076974545807159296"  # fallback nếu AM không có trong mapping

# Mỗi AM gửi vào đúng kênh GTalk riêng của họ
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
    "Cao Thị Thanh Thủy": "2083241927281995776",
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    return paths[0]

F_BOLD = _first_existing(FONT_CANDIDATES_BOLD)
F_REG = _first_existing(FONT_CANDIDATES_REG)


def font(size, bold=False):
    return ImageFont.truetype(F_BOLD if bold else F_REG, size)


# ===== DỮ LIỆU MẪU (fallback nếu không đọc được sheet, hoặc dùng để test layout) =====
DATA_SAMPLE_BC = [
    {
        "bc": "(BTH) Hàm Thắng",
        "am": "Nguyễn Ngọc Khánh",
        "staff": [
            ("Nguyễn Thanh Quốc", 74, 53, 71.62, 4, -2.31),
            ("Lâm Đức Mỹ", 55, 42, 76.36, 3, 1.05),
            ("Nguyễn Hoàng Duy", 68, 52, 76.47, 18, 0.0),
            ("Ngô Thanh Dũng", 49, 39, 79.59, 0, None),
        ]
    },
    {
        "bc": "(BTH) Lương Sơn",
        "am": "Nguyễn Ngọc Khánh",
        "staff": [
            ("Sử Vĩnh Hưng", 79, 58, 73.42, 4, 4.12),
            ("Đặng Minh Tuấn", 74, 57, 77.03, 27, -1.50),
        ]
    },
    {
        "bc": "(BTH) Mũi Né",
        "am": "Nguyễn Ngọc Khánh",
        "staff": [
            ("Phạm Xuân Thịnh", 36, 17, 47.22, 0, -8.90),
            ("Trần Ngọc Sang", 12, 6, 50.00, 0, 3.33),
            ("Trần Hải Triều", 54, 34, 62.96, 7, 0.0),
        ]
    },
    {
        "bc": "(BTH) Hàm Tân",
        "am": "Lê Thanh Nhựt",
        "staff": [
            ("Dương Ngọc Thuận", 89, 48, 53.93, 2, -5.60),
            ("Nguyễn Hữu Anh Vũ", 40, 29, 72.50, 1, 2.20),
            ("Nguyễn Trọng Nghĩa", 81, 64, 79.01, 7, None),
        ]
    },
    {
        "bc": "(BTH) Liên Hương",
        "am": "Nguyễn Duy Long",
        "staff": [
            ("Đặng Văn Vinh", 37, 28, 75.68, 1, 1.10),
            ("Lê Thanh Hiếu", 51, 39, 76.47, 0, -0.80),
        ]
    }
]


def color_for(pct):
    """Màu sắc dạng Pill Badge cho cột %GTC: Đạt (Xanh), Cải thiện (Vàng), Yếu (Đỏ)"""
    if pct >= 90.0:
        return (222, 247, 236), (3, 84, 63)     # Xanh lá pastel / chữ xanh đậm
    elif pct >= 80.0:
        return (254, 243, 199), (180, 83, 9)    # Vàng pastel / chữ vàng đậm
    else:
        return (253, 232, 232), (155, 28, 28)    # Đỏ pastel / chữ đỏ đậm


def measure(draw, text, f):
    bbox = draw.textbbox((0, 0), text, font=f)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_gradient_header(d, x0, y0, x1, y1, color_start, color_end):
    """Vẽ dải màu gradient từ color_start sang color_end theo chiều dọc"""
    for y in range(y0, y1 + 1):
        ratio = (y - y0) / (y1 - y0)
        r = int(color_start[0] * (1 - ratio) + color_end[0] * ratio)
        g = int(color_start[1] * (1 - ratio) + color_end[1] * ratio)
        b = int(color_start[2] * (1 - ratio) + color_end[2] * ratio)
        d.line([x0, y, x1, y], fill=(r, g, b))


def generate_report_image(bc_block, report_date_str, updated_at_str, out_path, show_comparison=False):
    bc_name = bc_block["bc"]
    am_name = bc_block["am"]
    staff = bc_block["staff"]
    total_nv = len(staff)

    # Tính toán các chỉ số tổng hợp
    total_gan = sum(x[1] for x in staff)
    total_tc = sum(x[2] for x in staff)
    avg_pct = (sum(x[3] for x in staff) / total_nv) if total_nv > 0 else 0.0

    # Khởi tạo font chữ chuyên nghiệp
    f_title = font(20, bold=True)
    f_bc = font(28, bold=True)
    f_badge = font(14, bold=True)
    f_kpi_title = font(11, bold=True)
    f_kpi_val = font(26, bold=True)
    f_kpi_sub = font(11)
    f_colhead = font(13, bold=True)
    f_name = font(15, bold=True)
    f_val = font(16, bold=True)
    f_delta = font(14, bold=True)

    ROW_H = 46
    W = 1000
    PAD = 30
    GAP = 20

    # Layout dimensions
    header_h = 130
    kpi_h = 110
    table_header_h = 40
    table_card_h = table_header_h + total_nv * ROW_H + 20
    H = header_h + 20 + kpi_h + 20 + table_card_h + 40

    img = Image.new("RGB", (W, H), (244, 245, 247))  # Màu nền xám nhạt dịu mắt (slate-50)
    d = ImageDraw.Draw(img)

    # 1. Header (Nền trắng tinh tế)
    d.rectangle([0, 0, W, header_h], fill=(255, 255, 255))
    d.line([0, 0, W, 0], fill=(234, 88, 12), width=4) # Accent top bar màu Cam GHN
    d.line([0, header_h, W, header_h], fill=(226, 232, 240), width=1) # Đường kẻ dưới mờ

    # Tiêu đề & Tên Bưu Cục
    d.text((PAD, 22), "BÁO CÁO HIỆU SUẤT VẬN HÀNH BƯU CỤC", font=f_title, fill=(100, 116, 139))
    d.text((PAD, 55), bc_name, font=f_bc, fill=(15, 23, 42))
    
    # Tag AM nằm bên cạnh tên bưu cục (nền xanh nhạt viền mờ)
    am_lbl = f"AM: {am_name}"
    aw, ah = measure(d, am_lbl, f_badge)
    bc_w, bc_h = measure(d, bc_name, f_bc)
    
    badge_x = PAD + bc_w + 16
    badge_y = 61
    if badge_x + aw + 20 > W - PAD: # Tự động xuống dòng nếu tên bưu cục quá dài
        badge_x = PAD
        badge_y = 95
    d.rounded_rectangle([badge_x, badge_y, badge_x + aw + 16, badge_y + ah + 8], radius=6, fill=(239, 246, 255), outline=(191, 219, 254), width=1)
    d.text((badge_x + 8, badge_y + 3), am_lbl, font=f_badge, fill=(30, 64, 175))

    # Cập nhật thời gian
    date_lbl = f"Cập nhật: {updated_at_str}"
    dw, dh = measure(d, date_lbl, font(13))
    d.text((W - PAD - dw, 22), date_lbl, font=font(13), fill=(100, 116, 139))

    # 2. KPI Summary Cards (3 Thẻ Tóm Tắt có màu nền pastel riêng biệt)
    card_w = (W - 2 * PAD - 2 * GAP) // 3
    c_y0 = header_h + 20
    c_y1 = c_y0 + kpi_h
    
    # Card 1: Nhân Sự (Màu xanh dương chủ đạo)
    c1_x0 = PAD
    c1_x1 = PAD + card_w
    d.rounded_rectangle([c1_x0, c_y0, c1_x1, c_y1], radius=10, fill=(239, 246, 255), outline=(191, 219, 254), width=1) # xanh dương pastel
    d.rectangle([c1_x0 + 1, c_y0 + 12, c1_x0 + 5, c_y1 - 12], fill=(59, 130, 246)) # Thanh dọc xanh dương sẫm
    d.text((c1_x0 + 18, c_y0 + 15), "NHÂN SỰ HOẠT ĐỘNG", font=f_kpi_title, fill=(30, 64, 175))
    d.text((c1_x0 + 18, c_y0 + 35), str(total_nv), font=f_kpi_val, fill=(29, 78, 216))
    d.text((c1_x0 + 18, c_y0 + 75), "Riders hoạt động trong ngày", font=f_kpi_sub, fill=(71, 85, 105))

    # Card 2: Hiệu suất bưu cục (Màu động theo mức hiệu suất)
    c2_x0 = c1_x1 + GAP
    c2_x1 = c2_x0 + card_w
    
    # Xác định màu sắc động theo hiệu suất
    if avg_pct >= 90:
        eval_lbl = "ĐẠT CHỈ TIÊU"
        card_bg = (240, 253, 244)     # xanh lá pastel
        card_border = (187, 247, 208) # viền xanh lá mờ
        eval_color = (21, 128, 61)    # chữ xanh lá đậm
    elif avg_pct >= 80:
        eval_lbl = "CẦN CẢI THIỆN"
        card_bg = (254, 243, 199)     # vàng pastel
        card_border = (253, 230, 138) # viền vàng mờ
        eval_color = (180, 83, 9)     # chữ vàng đậm
    else:
        eval_lbl = "YẾU / CẦN CẢI THIỆN"
        card_bg = (254, 242, 242)     # đỏ hồng pastel
        card_border = (254, 202, 202) # viền đỏ mờ
        eval_color = (185, 28, 28)    # chữ đỏ đậm

    d.rounded_rectangle([c2_x0, c_y0, c2_x1, c_y1], radius=10, fill=card_bg, outline=card_border, width=1)
    d.rectangle([c2_x0 + 1, c_y0 + 12, c2_x0 + 5, c_y1 - 12], fill=eval_color)
    d.text((c2_x0 + 18, c_y0 + 15), "HIỆU SUẤT GIAO TC", font=f_kpi_title, fill=eval_color)
    d.text((c2_x0 + 18, c_y0 + 35), f"{avg_pct:.2f}%", font=f_kpi_val, fill=eval_color)
    d.text((c2_x0 + 18, c_y0 + 75), eval_lbl, font=f_kpi_sub, fill=eval_color)

    # Card 3: Sản lượng (Màu tím/violet chủ đạo)
    c3_x0 = c2_x1 + GAP
    c3_x1 = c3_x0 + card_w
    d.rounded_rectangle([c3_x0, c_y0, c3_x1, c_y1], radius=10, fill=(250, 245, 255), outline=(233, 213, 255), width=1) # tím pastel
    d.rectangle([c3_x0 + 1, c_y0 + 12, c3_x0 + 5, c_y1 - 12], fill=(139, 92, 246)) # Thanh dọc tím
    d.text((c3_x0 + 18, c_y0 + 15), "SẢN LƯỢNG ĐƠN GIAO", font=f_kpi_title, fill=(109, 40, 217))
    d.text((c3_x0 + 18, c_y0 + 35), f"{total_tc} / {total_gan}", font=f_kpi_val, fill=(126, 34, 206))
    d.text((c3_x0 + 18, c_y0 + 75), "Giao TC / Gán giao (đơn)", font=f_kpi_sub, fill=(71, 85, 105))

    # 3. Employee Performance Table
    card_x0 = PAD
    card_x1 = W - PAD
    table_card_y0 = c_y1 + 20
    
    # Vẽ card nền bao bọc toàn bộ bảng (nền xám slate-50 cực nhạt)
    d.rounded_rectangle([card_x0, table_card_y0, card_x1, table_card_y0 + table_card_h], radius=14, fill=(248, 250, 252), outline=(226, 232, 240), width=1)

    inner_x = card_x0 + 20
    inner_right = card_x1 - 20
    inner_w = inner_right - inner_x

    # Thiết lập tọa độ và chiều rộng các cột
    if show_comparison:
        x_stt = inner_x
        x_name = x_stt + 45
        x_gan = x_name + 210
        x_tc = x_gan + 110
        x_pct = x_tc + 110
        x_ltc = x_pct + 140
        x_delta = x_ltc + 90
        
        col_w_gan = 90
        col_w_tc = 90
        col_w_pct = 120
        col_w_ltc = 70
        col_w_delta = 120
    else:
        x_stt = inner_x
        x_name = x_stt + 50
        x_gan = x_name + 280
        x_tc = x_gan + 130
        x_pct = x_tc + 130
        x_ltc = x_pct + 150
        
        col_w_gan = 100
        col_w_tc = 100
        col_w_pct = 130
        col_w_ltc = 80

    # Vẽ tiêu đề cột
    th_y = table_card_y0 + 15
    d.text((x_stt, th_y), "STT", font=f_colhead, fill=(51, 65, 85))
    d.text((x_name, th_y), "Nhân viên", font=f_colhead, fill=(51, 65, 85))
    
    # Helper vẽ text canh lề phải
    def text_right_align(val, x_start, width, y_pos, font_obj, fill_color):
        tw, th = measure(d, val, font_obj)
        d.text((x_start + width - tw, y_pos), val, font=font_obj, fill=fill_color)

    text_right_align("Gán giao", x_gan, col_w_gan, th_y, f_colhead, (51, 65, 85))
    text_right_align("Giao TC", x_tc, col_w_tc, th_y, f_colhead, (51, 65, 85))
    text_right_align("%GTC", x_pct, col_w_pct, th_y, f_colhead, (51, 65, 85))
    text_right_align("LTC", x_ltc, col_w_ltc, th_y, f_colhead, (51, 65, 85))
    if show_comparison:
        text_right_align("So với 10h", x_delta, col_w_delta, th_y, f_colhead, (51, 65, 85))

    # Đường kẻ dưới header bảng
    d.line([inner_x, th_y + 25, inner_right, th_y + 25], fill=(226, 232, 240), width=1)

    # Duyệt vẽ từng nhân viên
    row_y = th_y + 35
    for idx, row in enumerate(staff):
        name, gan, tc, pct, ltc = row[0], row[1], row[2], row[3], row[4]
        delta = row[5] if len(row) > 5 else None

        # Zebra striping (dòng xen kẽ trắng nổi bật trên nền xám)
        if idx % 2 == 1:
            d.rounded_rectangle([inner_x - 10, row_y - 8, inner_right + 10, row_y + ROW_H - 12], radius=6, fill=(255, 255, 255))

        # Số thứ tự
        d.text((x_stt, row_y), str(idx + 1), font=f_colhead, fill=(148, 163, 184))

        # Tên nhân viên
        d.text((x_name, row_y), name, font=f_name, fill=(30, 41, 59))

        # Số liệu
        text_right_align(str(gan), x_gan, col_w_gan, row_y, f_val, (71, 85, 105))
        text_right_align(str(tc), x_tc, col_w_tc, row_y, f_val, (71, 85, 105))

        # %GTC Badge Pill
        pct_text = f"{pct:.2f}%"
        bg_c, txt_c = color_for(pct)
        pw, ph = measure(d, pct_text, f_val)
        
        pill_w = pw + 20
        pill_h = ph + 6
        pill_x1 = x_pct + col_w_pct
        pill_x0 = pill_x1 - pill_w
        d.rounded_rectangle([pill_x0, row_y - 2, pill_x1, row_y + pill_h - 2], radius=10, fill=bg_c)
        d.text((pill_x0 + 10, row_y + 1), pct_text, font=f_val, fill=txt_c)

        # LTC
        text_right_align(str(ltc), x_ltc, col_w_ltc, row_y, f_val, (71, 85, 105))

        # So sánh 10h (Delta)
        if show_comparison:
            if delta is None:
                delta_text = "—"
                delta_color = (148, 163, 184)
            elif delta > 0:
                delta_text = f"▲ +{delta:.2f}%"
                delta_color = (21, 128, 61) # xanh lá
            elif delta < 0:
                delta_text = f"▼ {delta:.2f}%"
                delta_color = (185, 28, 28) # đỏ
            else:
                delta_text = "0.00%"
                delta_color = (100, 116, 139)
            text_right_align(delta_text, x_delta, col_w_delta, row_y, f_delta, delta_color)

        row_y += ROW_H

    img.save(out_path)
    print(f"✅ Đã tạo ảnh: {out_path}")
    return out_path


# ===== GỬI ẢNH VÀO GTALK =====

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

    print(f"📡 Đang upload ảnh lên GTalk (Channel: {channel_id})...")
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


def get_channel_for_am(am_name):
    norm_target = unicodedata.normalize('NFC', am_name.strip()).lower()
    for name, channel_id in AM_CHANNEL_MAP.items():
        if unicodedata.normalize('NFC', name.strip()).lower() == norm_target:
            return channel_id
    print(f"⚠️ Không tìm thấy kênh GTalk riêng cho AM '{am_name}' — dùng kênh mặc định.")
    return GTALK_CHANNEL_ID


def send_gtalk_report(force_baseline=False):
    now = datetime.now()
    report_date_str = now.strftime("%-d thg %-m, %Y") if os.name != "nt" else now.strftime("%d thg %m, %Y")
    updated_at_str = now.strftime("%H:%M · %d/%m/%Y")
    today_str = now.strftime("%Y-%m-%d")

    # Mốc 10h (chạy trước 12h trưa) = baseline, không so sánh.
    is_baseline_run = force_baseline or now.hour < 12
    show_comparison = not is_baseline_run
    if force_baseline:
        print("⚠️ Đang CHẠY Ở CHẾ ĐỘ ÉP LƯU BASELINE (--force-baseline) — "
              "dùng dữ liệu hiện tại làm mốc so sánh cho phần còn lại trong ngày.")
    print(f"🕐 Giờ hiện tại: {now.strftime('%H:%M')} -> "
          f"{'MỐC BASELINE (không so sánh)' if is_baseline_run else 'MỐC SO SÁNH (có cột so với 10h)'}")

    try:
        data = load_data_from_sheet(is_baseline_run=is_baseline_run, today_str=today_str)
        if not data:
            print("⚠️ Không có nhân viên nào trong sheet hôm nay — dùng dữ liệu mẫu để test.")
            data = DATA_SAMPLE_BC
    except Exception as e:
        print(f"⚠️ Lỗi đọc data từ Google Sheet: {e}")

    print(f"📦 Sẽ tạo và gửi {len(data)} ảnh (mỗi Bưu cục 1 ảnh)...")

    for i, bc_block in enumerate(data, start=1):
        bc_name = bc_block["bc"]
        am_name = bc_block["am"]
        staff = bc_block["staff"]
        total_nv = len(staff)

        safe_name = "".join(c if c.isalnum() else "_" for c in bc_name)
        out_path = os.path.join(BASE_DIR, f"nvptt_{safe_name}.png")

        print(f"\n[{i}/{len(data)}] 🖼️ Đang tạo ảnh cho Bưu cục: {bc_name} (AM: {am_name}, {total_nv} nhân viên)...")
        generate_report_image(bc_block, report_date_str, updated_at_str, out_path, show_comparison=show_comparison)

        caption = (
            f"<b>CẬP NHẬT HIỆU SUẤT BƯU CỤC</b>\n"
            f"Bưu cục: <b>{bc_name}</b> (AM: <b>{am_name}</b>)\n"
            f"Số lượng: <b>{total_nv} nhân viên</b>\n"
            f"Ngày báo cáo {report_date_str} · Cập nhật lúc {updated_at_str}"
        )

        am_channel_id = get_channel_for_am(am_name)
        ok = send_report_to_gtalk(out_path, caption, channel_id=am_channel_id)
        if not ok:
            print(f"⚠️ Gửi ảnh Bưu cục '{bc_name}' thất bại, tiếp tục bưu cục kế tiếp...")

        try:
            os.remove(out_path)
        except Exception:
            pass

        if i < len(data):
            time.sleep(2)  # tránh gửi quá dồn dập, tránh bị rate-limit

    print("\n🎉 Đã xử lý xong toàn bộ Bưu cục.")


# ======================================================================
# MAIN: CHẠY TUẦN TỰ CẢ 2 BƯỚC
# ======================================================================
def main():
    force_baseline = "--force-baseline" in sys.argv

    print("=" * 60)
    print("BƯỚC 1/2: TẢI & REPLACE DỮ LIỆU LẤY HÀNG / GIAO HÀNG")
    print("=" * 60)
    try:
        run_job()
    except Exception as e:
        print(f"❌ Lỗi ở bước tải/replace: {e}")
        print("⚠️ Vẫn tiếp tục bước 2 (đọc BaoCao) dù bước 1 lỗi.")

    print("\n⏳ Đợi 10 giây để Google Sheet tính lại công thức trong tab BaoCao...")
    time.sleep(10)

    print("\n" + "=" * 60)
    print("BƯỚC 2/2: ĐỌC BAOCAO, TẠO ẢNH & GỬI GTALK")
    print("=" * 60)
    send_gtalk_report(force_baseline=force_baseline)

    print("\n🎉 HOÀN THÀNH TOÀN BỘ PIPELINE.")


if __name__ == "__main__":
    main()
