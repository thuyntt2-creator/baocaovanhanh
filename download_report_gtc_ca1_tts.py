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
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"
GOOGLE_SHEET_KEY = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
TAB_BAO_CAO = "rawGTCTTS"
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

def get_gspread_client(sheet_key=None):
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    auth_user_candidates = [
        os.path.join(SCRIPT_DIR, 'authorized_user.json'),
        os.path.join(SCRIPT_DIR, 'credentials_oauth.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for auth_file in auth_user_candidates:
        if os.path.exists(auth_file):
            try:
                from google.oauth2.credentials import Credentials as UserCredentials
                creds = UserCredentials.from_authorized_user_file(auth_file, scopes=scopes)
                gc = gspread.authorize(creds)
                if sheet_key:
                    gc.open_by_key(sheet_key)
                print(f"✔️ Đã xác thực thành công qua {auth_file}")
                return gc
            except Exception as e:
                pass

    json_path = get_credentials_path()
    if json_path and os.path.exists(json_path):
        try:
            credentials = Credentials.from_service_account_file(json_path, scopes=scopes)
            gc = gspread.authorize(credentials)
            if sheet_key:
                gc.open_by_key(sheet_key)
            print(f"✔️ Đã xác thực thành công qua Service Account ({json_path})")
            return gc
        except Exception as e:
            pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")

def get_or_create_worksheet(sh, tab_name):
    try:
        ws = sh.worksheet(tab_name)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=10)
        print(f"✅ Đã tạo tab mới: '{tab_name}'")
        return ws

def to_vn_pct_str(val):
    if val is None or pd.isna(val) or val == "":
        return "0,00%"
    try:
        val_str = str(val).strip().replace('%', '')
        val_str = val_str.replace(',', '.')
        num = float(val_str)
        if num > 1.0:
            num = num / 100.0
        pct_str = f"{num * 100:.2f}%"
        return pct_str.replace('.', ',')
    except Exception:
        val_str = str(val).strip()
        if '%' not in val_str:
            val_str += '%'
        return val_str.replace('.', ',')

def to_clean_volume_str(val):
    if val is None or pd.isna(val) or val == "":
        return "0"
    try:
        val_str = str(val).strip().replace('.', '').replace(',', '')
        return str(int(round(float(val_str))))
    except Exception:
        return "0"

def upload_to_sheet_tab(df, sheet_key, tab_name):
    try:
        gc = get_gspread_client(sheet_key)
        sh = gc.open_by_key(sheet_key)

        ws = get_or_create_worksheet(sh, tab_name)
        print(f"📤 Đang xử lý và ghi dữ liệu vào tab '{tab_name}'...")
        ws.clear()

        # Format columns exactly as raw download shown in user screenshot
        new_df = pd.DataFrame()
        
        # 1. Cấp Quản Lý
        col_quan_ly = next((c for c in df.columns if 'quản lý' in c.lower() or 'quan ly' in c.lower()), 'Cấp Quản Lý')
        new_df['Cấp Quản Lý'] = df[col_quan_ly].fillna("")
        
        # 2. Chi tiết
        col_chi_tiet = next((c for c in df.columns if 'chi tiết' in c.lower() or 'chi tiet' in c.lower() or 'bưu cục' in c.lower()), 'Chi tiết')
        new_df['Chi tiết'] = df[col_chi_tiet].fillna("")
        
        # 3. Time (Date)
        col_time = next((c for c in df.columns if 'time' in c.lower() or 'ngày' in c.lower() or 'thời gian' in c.lower()), 'Time')
        new_df['Time'] = df[col_time].fillna("")
        
        # 4. Volume
        col_vol = next((c for c in df.columns if 'volume' in c.lower() or 'sản lượng' in c.lower() or 'san luong' in c.lower() or 'gtc' in c.lower()), 'Volume')
        new_df['Volume'] = df[col_vol].apply(to_clean_volume_str)
        
        # 5. % Gán
        col_gan = next((c for c in df.columns if 'gán' in c.lower() or 'gan' in c.lower()), '% Gán')
        new_df['% Gán'] = df[col_gan].apply(to_vn_pct_str)
        
        # 6. % GTC
        col_gtc = next((c for c in df.columns if '% gtc' in c.lower() or '% ltc' in c.lower() or 'tỷ lệ gtc' in c.lower()), '% GTC')
        new_df['% GTC'] = df[col_gtc].apply(to_vn_pct_str)
        
        # 7. % Chuyển trả
        col_tra = next((c for c in df.columns if 'chuyển trả' in c.lower() or 'chuyen tra' in c.lower() or 'đóng kiện' in c.lower()), '% Chuyển trả')
        new_df['% Chuyển trả'] = df[col_tra].apply(to_vn_pct_str)
        
        # 8. Leadtime
        col_lt = next((c for c in df.columns if 'leadtime' in c.lower() or 'lead time' in c.lower() or 'lc' in c.lower()), 'Leadtime')
        new_df['Leadtime'] = df[col_lt].apply(lambda x: str(x).strip().replace('.', ','))

        # Upload
        data_to_upload = [new_df.columns.values.tolist()] + new_df.values.tolist()
        ws.update(data_to_upload, value_input_option='USER_ENTERED')
        print(f"🎉 Tab '{tab_name}': {len(new_df)} dòng, {len(new_df.columns)} cột — cập nhật thành công!")
        return True
    except Exception as e:
        print(f"❌ Lỗi khi ghi tab '{tab_name}': {e}")
        return False

def read_csv_smart(filepath):
    combos = [
        ('utf-8-sig', ','), ('utf-16', '\t'), ('utf-16-le', '\t'),
        ('utf-8-sig', '\t'), ('utf-8-sig', ';'), ('utf-8', ','),
        ('utf-8', '\t'), ('latin1', ','), ('latin1', '\t'),
    ]
    for enc, sep in combos:
        try:
            tmp = pd.read_csv(filepath, encoding=enc, sep=sep, dtype=str).fillna("")
            if len(tmp.columns) > 1:
                print(f"✅ Đọc CSV: encoding={enc}, sep='{sep}', {len(tmp)} dòng, {len(tmp.columns)} cột.")
                return tmp
        except Exception:
            continue
    return None

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

def apply_filter(report_frame, ghn_page, anchor, filter_label, target_value):
    print(f"\n🖱️ Click filter '{filter_label}'...", flush=True)
    btn_below = get_elements_below_anchor(
        report_frame, anchor,
        f'button:has-text("{filter_label}")',
        ghn_page
    )
    if btn_below:
        btn = btn_below[0]
        print(f"   ✅ Tìm thấy button '{filter_label}' dưới heading.", flush=True)
    else:
        print(f"   ⚠️ Không tìm thấy button '{filter_label}' theo vị trí, dùng fallback first.", flush=True)
        btn = report_frame.locator(f'button:has-text("{filter_label}")').first

    btn.scroll_into_view_if_needed()
    ghn_page.wait_for_timeout(500)
    btn.click(force=True)
    ghn_page.wait_for_timeout(2500)

    print(f"   👉 Chọn duy nhất '{target_value}' trong dropdown...", flush=True)
    target_opt = None
    all_opts = report_frame.locator('div.item, div.item-single, div.row, div[role="option"]').all()
    for opt in all_opts:
        try:
            if opt.is_visible():
                text = opt.inner_text().strip()
                first_line = text.split('\n')[0].strip()
                if first_line.lower() == target_value.lower():
                    target_opt = opt
                    break
        except Exception:
            continue

    if target_opt:
        print(f"   🖱️ Hover và click chọn 'Chỉ' (Only) cho '{target_value}'...", flush=True)
        target_opt.hover()
        ghn_page.wait_for_timeout(1000)
        
        only_btn = target_opt.locator('span.only, .only').first
        if only_btn.count() == 0:
            only_btn = target_opt.locator('text="Chỉ"').first
        if only_btn.count() == 0:
            only_btn = target_opt.locator('text="Only"').first
            
        if only_btn.count() > 0:
            only_btn.click(force=True)
            print("   ✅ Đã click button 'Chỉ' (Only).", flush=True)
        else:
            print("   ⚠️ Không tìm thấy button 'Chỉ', click trực tiếp...", flush=True)
            target_opt.evaluate("el => el.click()")
        ghn_page.wait_for_timeout(1500)
    else:
        print(f"   ❌ Không tìm thấy option '{target_value}'!", flush=True)

    try:
        if anchor:
            anchor.click(force=True, timeout=3000)
    except Exception:
        pass
    ghn_page.keyboard.press("Escape")
    
    print(f"⏳ Chờ bảng tải lại dữ liệu sau khi lọc '{target_value}' (8 giây)...", flush=True)
    ghn_page.wait_for_timeout(8000)

def export_table(ghn_page, report_frame, label, section_keywords, filter_btn_index_fallback=0, table_index_fallback=0, expand_details=False, filter_tts=False, filter_ca1=False):
    print(f"\n{'='*50}", flush=True)
    print(f"📋 BẮT ĐẦU XUẤT: {label}", flush=True)
    print(f"{'='*50}", flush=True)

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

    # ---- Bước 0: Scroll xuống để load hết DOM ----
    ghn_page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    ghn_page.wait_for_timeout(2000)

    # ---- Bước 1: Tìm anchor heading của section ----
    print(f"🔍 Tìm heading section '{label}'...", flush=True)
    anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

    if anchor:
        anchor.scroll_into_view_if_needed()
        ghn_page.wait_for_timeout(1500)
        print(f"📌 Đã scroll đến section '{label}'.", flush=True)

        # ---- Bước 2: Tìm button Chi tiết NẰM DƯỚI anchor ----
        print("🔍 Tìm button 'Chi tiết' thuộc section này...", flush=True)
        chi_tiet_below = get_elements_below_anchor(
            report_frame, anchor,
            'button:has-text("Chi tiết")',
            ghn_page
        )
        if chi_tiet_below:
            chi_tiet_btn = chi_tiet_below[0]
            print("✅ Dùng button Chi tiết đầu tiên dưới heading.", flush=True)
        else:
            print(f"⚠️ Không tìm được button theo vị trí, fallback index={filter_btn_index_fallback}", flush=True)
            all_btns = report_frame.locator('button:has-text("Chi tiết")').all()
            chi_tiet_btn = all_btns[filter_btn_index_fallback] if len(all_btns) > filter_btn_index_fallback else all_btns[0]
    else:
        print(f"⚠️ Không tìm thấy heading, fallback index={filter_btn_index_fallback}", flush=True)
        all_btns = report_frame.locator('button:has-text("Chi tiết")').all()
        chi_tiet_btn = all_btns[filter_btn_index_fallback] if len(all_btns) > filter_btn_index_fallback else all_btns[0]

    # ---- Bước 3: Click filter Chi tiết → chọn Bưu Cục - Tỉnh ----
    print("🖱️ Click 'Chi tiết'...", flush=True)
    chi_tiet_btn.scroll_into_view_if_needed()
    ghn_page.wait_for_timeout(500)
    chi_tiet_btn.click(force=True)
    ghn_page.wait_for_timeout(2500)

    print("👉 Xử lý chọn 'Bưu Cục - Tỉnh' trong dropdown...", flush=True)
    target_opt = None
    all_opts = report_frame.locator('div.item, div.item-single').all()
    for opt in all_opts:
        try:
            if opt.is_visible():
                text = opt.inner_text().strip()
                if text.lower() == "bưu cục - tỉnh":
                    target_opt = opt
                    break
        except Exception:
            continue

    if target_opt:
        print("   🖱️ Click chọn 'Bưu Cục - Tỉnh'...", flush=True)
        target_opt.evaluate("el => el.click()")
        ghn_page.wait_for_timeout(1000)
    else:
        print("   ❌ Không tìm thấy option 'Bưu Cục - Tỉnh'!", flush=True)

    try:
        if anchor:
            anchor.click(force=True, timeout=3000)
    except Exception:
        pass
    ghn_page.keyboard.press("Escape")
    
    print("⏳ Chờ bảng tải lại dữ liệu sau khi đổi bộ lọc (8 giây)...", flush=True)
    ghn_page.wait_for_timeout(8000)

    # ---- Bước 3b: Click filter Loại khách hàng → chọn duy nhất TTS (nếu filter_tts=True) ----
    if filter_tts:
        apply_filter(report_frame, ghn_page, anchor, "Loại khách hàng", "TTS")

    # ---- Bước 3c: Click filter Loại hàng → chọn duy nhất Hàng Mới Ca 1 (nếu filter_ca1=True) ----
    if filter_ca1:
        apply_filter(report_frame, ghn_page, anchor, "Loại hàng", "Hàng Mới Ca 1")

    # ---- Bước 4: Click [+] mở rộng cột Chi tiết (nếu expand_details=True) ----
    if expand_details:
        is_already_expanded = False
        try:
            lh_count = report_frame.locator('div.header-cell:has-text("Loại Hàng")').count()
            if lh_count > 0:
                is_already_expanded = True
                print("ℹ️ Bảng đã được mở rộng sẵn (tìm thấy header 'Loại Hàng'). Bỏ qua bấm [+].", flush=True)
        except Exception as e:
            print(f"⚠️ Lỗi khi kiểm tra trạng thái mở rộng: {e}", flush=True)

        if not is_already_expanded:
            print("➕ Bấm [+] mở rộng bảng...", flush=True)
            
            hc = None
            for attempt in range(3):
                try:
                    if report_frame.is_detached():
                        raise Exception("Frame is detached")

                    if anchor:
                        header_cells_below = get_elements_below_anchor(
                            report_frame, anchor,
                            'div.header-cell:has-text("Chi tiết")',
                            ghn_page
                        )
                        hc = header_cells_below[0] if header_cells_below else None
                    else:
                        hc = None

                    if not hc:
                        all_hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').all()
                        if all_hc:
                            hc = all_hc[filter_btn_index_fallback] if len(all_hc) > filter_btn_index_fallback else all_hc[0]
                        else:
                            hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').first

                    hc.scroll_into_view_if_needed(timeout=5000)
                    ghn_page.wait_for_timeout(500)
                    hc.hover(timeout=5000)
                    ghn_page.wait_for_timeout(1500)

                    expand_btn = hc.locator('.expand-button, .interaction-button-wrapper').first
                    expand_btn.wait_for(state="attached", timeout=5000)
                    
                    expand_btn.click(force=True, timeout=5000)
                    print("✅ Đã bấm [+] bằng force click.", flush=True)
                    break
                except Exception as e:
                    print(f"⚠️ Thử {attempt+1}: Lỗi bấm [+] ({e}). Đang lấy lại frame...", flush=True)
                    ghn_page.wait_for_timeout(2000)
                    report_frame = get_looker_frame_helper(ghn_page)
                    anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

            print("⏳ Chờ bảng tải lại dữ liệu sau khi bấm [+] (8 giây)...", flush=True)
            ghn_page.wait_for_timeout(8000)
        else:
            print("⏭️ Bỏ qua bước bấm [+] vì đã mở rộng sẵn.", flush=True)
    else:
        print("⏭️ Bỏ qua bước bấm [+].", flush=True)

    # ---- Bước 5: Right-click table header cell → Export ----
    print("🧹 Dismiss any active overlays before clicking...", flush=True)
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
        "Export data",
        "Xuất dữ liệu",
        "Xuất biểu đồ...",
        "Export chart...",
        "Xuất biểu đồ",
        "Export chart",
        "Export",
        "Xuất"
    ]
    
    def is_menu_visible():
        scanned_info = []

        def check_frame(frame_obj, frame_name):
            for kw in menu_kws:
                kw_clean = clean_compare(kw)
                for sel in selectors:
                    try:
                        locator = frame_obj.locator(sel)
                        count = locator.count()
                        for idx in range(count):
                            el = locator.nth(idx)
                            is_vis = el.is_visible()
                            text = el.inner_text().strip()
                            
                            scanned_info.append(f"[{frame_name}] {sel} #{idx}: visible={is_vis}, text={repr(text)}")
                            
                            if is_vis:
                                lines = [clean_compare(line) for line in text.split('\n') if line.strip()]
                                if any(line == kw_clean for line in lines):
                                    return el, kw
                    except Exception as ex:
                        scanned_info.append(f"[{frame_name}] Error querying {sel}: {ex}")
            return None, None

        res_el, res_kw = check_frame(report_frame, "Frame")
        if res_el:
            return res_el, res_kw
            
        res_el, res_kw = check_frame(ghn_page, "Main")
        if res_el:
            return res_el, res_kw
            
        return None, None

    menu_el = None
    matched_kw = None
    click_targets = []
    
    try:
        if anchor:
            header_cells_below = get_elements_below_anchor(
                report_frame, anchor,
                'div.header-cell:has-text("Chi tiết")',
                ghn_page
            )
            if header_cells_below:
                click_targets.extend(header_cells_below)
    except Exception as e:
        print(f"⚠️ Lỗi tìm header cells dưới anchor: {e}", flush=True)

    try:
        all_hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').all()
        for hc in all_hc:
            if hc not in click_targets:
                click_targets.append(hc)
    except Exception:
        pass
        
    try:
        if anchor:
            tables_below = get_elements_below_anchor(
                report_frame, anchor,
                '.ng-ko-table, mat-table, table, .visual-container, .chart-container',
                ghn_page
            )
            for tb in tables_below:
                if tb not in click_targets:
                    click_targets.append(tb)
    except Exception:
        pass

    print(f"📌 Tìm thấy {len(click_targets)} đối tượng tiềm năng để click chuột phải.", flush=True)

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
            
            print(f"👉 Thử {i+1}.1: Click chuột phải (button='right') lên đối tượng...", flush=True)
            target.click(button="right", timeout=3000)
            ghn_page.wait_for_timeout(1500)
            
            menu_el, matched_kw = is_menu_visible()
            if menu_el:
                print(f"✅ Menu đã xuất hiện thành công bằng cách click chuột phải!", flush=True)
                break
                
            print(f"👉 Thử {i+1}.2: Dispatch event 'contextmenu' lên đối tượng...", flush=True)
            target.dispatch_event("contextmenu")
            ghn_page.wait_for_timeout(1500)
            
            menu_el, matched_kw = is_menu_visible()
            if menu_el:
                print(f"✅ Menu đã xuất hiện thành công bằng dispatch_event!", flush=True)
                break
        except Exception as e:
            print(f"⚠️ Thử click chuột phải lên đối tượng {i+1} lỗi: {e}", flush=True)
            try:
                ghn_page.keyboard.press("Escape")
            except:
                pass
            continue

    if not menu_el:
        print("❌ Chưa mở được menu bằng các đối tượng đã chọn. Thử tìm kiếm phần tử chung...", flush=True)
        try:
            try:
                bd = report_frame.locator('.cdk-overlay-backdrop').first
                if bd.count() > 0 and bd.is_visible():
                    bd.evaluate("el => el.click()")
                    ghn_page.wait_for_timeout(500)
                ghn_page.keyboard.press("Escape")
                ghn_page.wait_for_timeout(500)
            except Exception:
                pass

            any_hc = report_frame.locator('div.header-cell').first
            if any_hc.count() > 0:
                print("👉 Thử click chuột phải lên header cell bất kỳ...", flush=True)
                any_hc.hover()
                any_hc.click(button="right", timeout=3000)
                ghn_page.wait_for_timeout(1500)
                menu_el, matched_kw = is_menu_visible()
        except Exception:
            pass

    if not menu_el:
        raise Exception("Không thể mở được menu ngữ cảnh (context menu) để chọn xuất!")

    found_text = menu_el.inner_text().strip()
    print(f"✅ Đang chọn menu item: '{found_text}' (khớp từ khóa: '{matched_kw}')", flush=True)
    
    if matched_kw in ["Export data", "Xuất dữ liệu"]:
        print(f"   🖱️ Click trực tiếp vào '{matched_kw}'...", flush=True)
        try:
            menu_el.hover()
            ghn_page.wait_for_timeout(500)
            menu_el.click(timeout=3000)
        except Exception as ex:
            print(f"   ⚠️ Lỗi click thường, thử evaluate click: {ex}", flush=True)
            menu_el.evaluate("el => el.click()")
    else:
        print(f"   🖱️ Hover '{found_text}' để hiển thị submenu...", flush=True)
        menu_el.hover()
        ghn_page.wait_for_timeout(1500)
        
        print("   🔍 Tìm 'Export data' hoặc 'Xuất dữ liệu' trong submenu...", flush=True)
        sub_el = None
        for attempt in range(5):
            for sub_kw in ["Export data", "Xuất dữ liệu"]:
                sub_kw_clean = clean_compare(sub_kw)
                for ctx in [report_frame, ghn_page]:
                    for sel in selectors:
                        try:
                            locator = ctx.locator(sel)
                            count = locator.count()
                            for idx in range(count):
                                el = locator.nth(idx)
                                if el.is_visible():
                                    text = el.inner_text().strip()
                                    lines = [clean_compare(line) for line in text.split('\n') if line.strip()]
                                    if any(line == sub_kw_clean for line in lines):
                                        sub_el = el
                                        break
                            if sub_el: break
                        except Exception:
                            continue
                    if sub_el: break
                if sub_el: break
            if sub_el: break
            ghn_page.wait_for_timeout(500)
            
        if not sub_el:
            print(f"   🖱️ Hover chưa mở submenu, thử Click '{found_text}'...", flush=True)
            menu_el.click(force=True)
            ghn_page.wait_for_timeout(1500)
            
            for attempt in range(5):
                for sub_kw in ["Export data", "Xuất dữ liệu"]:
                    sub_kw_clean = clean_compare(sub_kw)
                    for ctx in [report_frame, ghn_page]:
                        for sel in selectors:
                            try:
                                locator = ctx.locator(sel)
                                count = locator.count()
                                for idx in range(count):
                                    el = locator.nth(idx)
                                    if el.is_visible():
                                        text = el.inner_text().strip()
                                        lines = [clean_compare(line) for line in text.split('\n') if line.strip()]
                                        if any(line == sub_kw_clean for line in lines):
                                            sub_el = el
                                            break
                                if sub_el: break
                            except Exception:
                                continue
                        if sub_el: break
                    if sub_el: break
                    
        if sub_el:
            try:
                sub_text = sub_el.inner_text().strip()
            except Exception:
                sub_text = "Xuất dữ liệu"
            print(f"   🖱️ Click '{sub_text}' từ submenu...", flush=True)
            sub_el.click(force=True)
        else:
            print("⚠️ Không tìm thấy submenu 'Export data' / 'Xuất dữ liệu'.", flush=True)
            
    ghn_page.wait_for_timeout(3000)

    # ---- Chọn định dạng CSV ----
    print("👉 Chọn định dạng CSV...", flush=True)
    try:
        dialog_selector = 'mat-dialog-container, mat-mdc-dialog-container, [role="dialog"]'
        try:
            report_frame.locator(dialog_selector).first.wait_for(state="attached", timeout=3000)
        except:
            try:
                ghn_page.locator(dialog_selector).first.wait_for(state="attached", timeout=2000)
            except:
                pass

        dialog = report_frame.locator(dialog_selector).first
        if dialog.count() == 0 or not dialog.is_visible():
            dialog = ghn_page.locator(dialog_selector).first

        csv_opt = None
        try:
            opt = dialog.locator('text="CSV"').first
            if opt.count() > 0 and opt.is_visible():
                inner_txt = opt.inner_text().strip()
                if inner_txt == "CSV":
                    csv_opt = opt
        except Exception:
            pass

        if csv_opt:
            print(f"   ✅ Click chọn định dạng: '{csv_opt.inner_text().strip()}'", flush=True)
            csv_opt.click(timeout=5000)
            ghn_page.wait_for_timeout(1000)
        else:
            print("   ⚠️ Không tìm thấy tùy chọn CSV trong dialog, giữ mặc định.", flush=True)
    except Exception as e:
        print(f"❌ Lỗi chọn định dạng CSV: {e}", flush=True)
        try:
            ghn_page.keyboard.press("Escape")
        except:
            pass
        raise e

    # ---- Bước 6: Tick Keep value formatting ----
    print("👉 Tích 'Keep value formatting'...", flush=True)
    try:
        keep_cb_selector = (
            '[formcontrolname="keepFormat"], mat-checkbox:has-text("Keep value formatting"), mat-checkbox:has-text("Giữ định dạng giá trị")'
        )
        keep_cb = report_frame.locator(keep_cb_selector).first
        if keep_cb.count() == 0 or not keep_cb.is_visible():
            keep_cb = ghn_page.locator(keep_cb_selector).first
            
        keep_cb.wait_for(state="visible", timeout=5000)
        
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
            keep_cb.click(timeout=5000)
            print("   ✅ Đã tích 'Keep value formatting'.", flush=True)
        else:
            print("   ℹ️ Checkbox đã tích sẵn.", flush=True)
    except Exception as e:
        print(f"⚠️ Không tick được checkbox: {e}", flush=True)

    # ---- Bước 7: Bấm Export → tải file ----
    print("👉 Bấm Export để tải file...", flush=True)
    confirm_btn_selector = (
        'mat-mdc-dialog-container button:has-text("Export"), mat-mdc-dialog-container button:has-text("Xuất"), '
        'mat-dialog-container button:has-text("Export"), mat-dialog-container button:has-text("Xuất"), '
        'button:has-text("Export"), button:has-text("Xuất")'
    )
    confirm_btn = report_frame.locator(confirm_btn_selector).first
    if confirm_btn.count() == 0 or not confirm_btn.is_visible():
        confirm_btn = ghn_page.locator(confirm_btn_selector).first

    with ghn_page.expect_download() as dl_info:
        confirm_btn.click(timeout=5000)

    download = dl_info.value
    fname = download.suggested_filename
    print(f"✅ File tải về: {fname}", flush=True)

    local_path = os.path.join(SCRIPT_DIR, fname)
    download.save_as(local_path)

    # ---- Bước 8: Đọc file ----
    df = None
    if local_path.endswith('.csv'):
        df = read_csv_smart(local_path)
    else:
        try:
            df = pd.read_excel(local_path, dtype=str).fillna("")
        except Exception as ex:
            print(f"❌ Lỗi đọc Excel: {ex}", flush=True)

    if df is not None:
        try:
            # Clean numeric columns to handle Vietnamese dotted thousands separator formats
            if 'Volume' in df.columns:
                df['Volume'] = df['Volume'].astype(str).str.replace('.', '', regex=False).str.replace(',', '', regex=False)
                df['Volume'] = pd.to_numeric(df['Volume'], errors='ignore')
                
            if 'Leadtime' in df.columns:
                df['Leadtime'] = df['Leadtime'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                df['Leadtime'] = pd.to_numeric(df['Leadtime'], errors='ignore')
            print("📊 Đã chuẩn hóa định dạng số cho cột Volume và Leadtime thành công.", flush=True)

            # Lọc bỏ các dòng Grand Total
            initial_len = len(df)
            for col in ['Chi tiết', 'Cấp quản lý']:
                if col in df.columns:
                    df = df[~df[col].astype(str).str.lower().str.strip().isin(['grand total', 'grandtotal', 'tổng cộng', 'tong cong'])]
                    df = df[~df[col].astype(str).str.lower().str.contains('grand total|grandtotal', regex=True)]
            if len(df) < initial_len:
                print(f"🧹 Đã lọc bỏ {initial_len - len(df)} dòng chứa từ khóa Grand Total.", flush=True)
        except Exception as e:
            print(f"⚠️ Cảnh báo lỗi khi chuẩn hóa số hoặc lọc Grand Total: {e}", flush=True)

    try:
        os.remove(local_path)
        print("🧹 Đã xóa file tạm.", flush=True)
    except Exception:
        pass

    return df

def run_job():
    print(f"\n🚀 BẮT ĐẦU CHẠY AUTO DOWNLOAD GTC TTS CA 1 LÚC: {time.strftime('%H:%M:%S')}", flush=True)
    print("🔌 Đang kết nối tới Chrome debug (127.0.0.1:9222)...", flush=True)

    with sync_playwright() as p:
        is_cdp = True
        browser = None
        context = None
        ghn_page = None
        
        # Check for cookies (for GitHub Actions headless mode)
        cookie_file = os.path.join(SCRIPT_DIR, "cookies_ghn.json")
        cookie_env = os.environ.get("GHN_COOKIES")
        
        has_cookies = cookie_env is not None or os.path.exists(cookie_file)
        
        if has_cookies:
            print("🍪 Đang chạy ở chế độ Headless với Cookies (GitHub Actions mode)...", flush=True)
            is_cdp = False
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            
            # Load cookies
            try:
                if cookie_env:
                    import json
                    storage_state = json.loads(cookie_env)
                    context.add_cookies(storage_state.get("cookies", []))
                else:
                    import json
                    with open(cookie_file, 'r', encoding='utf-8') as f:
                        storage_state = json.load(f)
                    context.add_cookies(storage_state.get("cookies", []))
                print("✅ Đã load cookies thành công.", flush=True)
            except Exception as e:
                print(f"❌ Lỗi khi load cookies: {e}", flush=True)
                
            ghn_page = context.new_page()
            ghn_page.goto(DASHBOARD_URL)
        else:
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
                    print("ℹ️ Không tìm thấy tab baocao.ghn.vn, tự động mở một tab mới trên Chrome debug...", flush=True)
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
        
        # Thiết lập thời gian chờ mặc định là 3 phút
        context.set_default_timeout(180000)
        ghn_page.set_default_timeout(180000)

        # Chờ đăng nhập nếu cần
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
                if is_cdp: browser.close()
                else: context.close()
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

            # Kiểm tra tên trang hiện tại
            page_name = ""
            try:
                pn = report_frame.locator("span.pageName, .pageName").first
                pn.wait_for(state="attached", timeout=10000)
                page_name = pn.text_content(timeout=3000).strip()
                print(f"📊 Trang hiện tại: '{page_name}'", flush=True)
            except Exception as e:
                print(f"⚠️ Không đọc được tên trang: {e}", flush=True)

            # Chuyển trang nếu đang ở MỤC LỤC
            if "MỤC LỤC" in page_name or not page_name:
                print("📑 Chuyển sang trang báo cáo (Click Next Page)...", flush=True)
                try:
                    next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                    next_btn.wait_for(state="attached", timeout=15000)
                    next_btn.hover()
                    next_btn.click(force=True)
                    print("✅ Đã click Next Page button.", flush=True)
                except Exception as e:
                    print(f"❌ Lỗi click Next Page button: {e}. Thử fallback click link...", flush=True)
                    try:
                        link = report_frame.locator('a:has-text("1. [VÙNG]")').first
                        link.click(force=True)
                    except Exception as ex:
                        print(f"❌ Click fallback link cũng lỗi: {ex}", flush=True)

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
                    if found: break

                if not found:
                    print("❌ Không chuyển được trang.", flush=True)
                    if is_cdp: browser.close()
                    else: context.close()
                    return

                print("⏳ Chờ trang ổn định (5 giây)...", flush=True)
                ghn_page.wait_for_timeout(5000)
            else:
                print("ℹ️ Đã ở sẵn trên trang báo cáo.", flush=True)
                ghn_page.wait_for_timeout(2000)

            # ══════════════════════════════════════════
            # BÁO CÁO B (TTS & CA 1 ONLY) → rawGTCTTS
            # ══════════════════════════════════════════
            df = export_table(
                ghn_page, report_frame,
                label="B. GIAO TRONG NGÀY (TTS CA 1 ONLY)",
                section_keywords=[
                    "B. BÁO CÁO GIAO TRONG NGÀY",
                    "BÁO CÁO GIAO TRONG NGÀY",
                    "B. GIAO TRONG NGÀY",
                    "GIAO TRONG NGÀY",
                ],
                filter_btn_index_fallback=1,
                table_index_fallback=1,
                expand_details=False,
                filter_tts=True,
                filter_ca1=True
            )
            if df is not None:
                upload_to_sheet_tab(df, GOOGLE_SHEET_KEY, TAB_BAO_CAO)
            else:
                print("❌ Không xuất được báo cáo B (TTS Ca 1).", flush=True)

            print(f"\n🏁 HOÀN THÀNH TOÀN BỘ LÚC: {time.strftime('%H:%M:%S')}", flush=True)

        except Exception as e:
            print(f"❌ Lỗi tổng quát: {e}", flush=True)
        finally:
            if is_cdp:
                if browser: browser.close()
            else:
                if context: context.close()

if __name__ == '__main__':
    run_job()
