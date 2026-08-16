import os
import sys
import time
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

TAB_BAO_CAO_1 = "XU_LY_THANH_CONG"
TAB_BAO_CAO_2 = "rawTTS"  # Ghi dữ liệu báo cáo B TTS (Giao trong ngày TTS) vào tab rawTTS
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
        existing_rows = ws.row_count
        
        print(f"📤 Ghi dữ liệu vào tab '{tab_name}'...")
        ws.batch_clear(["A:I"])
        
        # Clean dataframe columns before uploading to avoid locale parsing errors on Google Sheets
        if tab_name.lower() in ["raw", "rawtts"]:
            print("🧹 Đang làm sạch dữ liệu (Volume, %, Leadtime) trước khi tải lên...")
            df = df.copy()
            if 'Volume' in df.columns:
                vol_series = df['Volume'].astype(str).str.replace(',', '', regex=False)
                df['Volume'] = pd.to_numeric(vol_series, errors='coerce').fillna(0).astype(int)
            
            def clean_pct(val):
                if val is None or pd.isna(val) or val == "":
                    return 0.0
                if isinstance(val, (int, float)):
                    if val <= 1.0:
                        return float(val)
                    return float(val) / 100.0
                val_str = str(val).strip()
                has_pct = '%' in val_str
                val_str = val_str.replace('%', '')
                if ',' in val_str and '.' not in val_str:
                    val_str = val_str.replace(',', '.')
                try:
                    num = float(val_str)
                    if has_pct or num > 1.0:
                        return num / 100.0
                    return num
                except ValueError:
                    return 0.0

            def clean_float(val):
                if not val or pd.isna(val):
                    return 0.0
                val_str = str(val).strip()
                if ',' in val_str and '.' not in val_str:
                    val_str = val_str.replace(',', '.')
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0

            for col in ['% Gán', '% GTC', '% Chuyển trả']:
                if col in df.columns:
                    df[col] = df[col].apply(clean_pct)
                    
            if 'Leadtime' in df.columns:
                df['Leadtime'] = df['Leadtime'].apply(clean_float)

        df = df.fillna("")
        data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
        ws.update(data_to_upload)
        
        num_rows = len(df) + 1  # includes header row
        
        if tab_name.lower() in ["raw", "rawtts"]:
            headers = [
                'Sản Lượng Giao Thành Công', 'Sản Lượng Chuyển Trả', 'Sản Lượng Gán', 
                'Sản Lượng Trả', 'Sản Lượng Tồn', 'Sản Lượng Chưa Gán', 
                '% Chưa Gán', '%Tồn', 'Hàng Mới Về Trong Ngày', 'Tỉnh', 'Vùng', 'AM'
            ]
            ws.update(range_name="J1:U1", values=[headers])
            
            if num_rows >= 2:
                print(f"⚙️ Đang ghi công thức cho J2:U{num_rows}...")
                formulas_matrix = []
                for r in range(2, num_rows + 1):
                    row_formulas = [
                        f"=E{r}*G{r}",                                                      # J: Sản Lượng Giao Thành Công
                        f"=E{r}*H{r}",                                                      # K: Sản Lượng Chuyển Trả
                        f"=E{r}*F{r}",                                                      # L: Sản Lượng Gán
                        f"=E{r}*H{r}",                                                      # M: Sản Lượng Trả
                        f"=E{r}-J{r}-K{r}",                                                 # N: Sản Lượng Tồn
                        f"=E{r}-L{r}",                                                      # O: Sản Lượng Chưa Gán
                        f"=1-F{r}",                                                         # P: % Chưa Gán
                        f"=N{r}/E{r}",                                                      # Q: %Tồn
                        f'=IF(OR(C{r}="Hàng Mới Ca 1"; C{r}="Hàng Mới Ca 2"); E{r}; 0)',    # R: Hàng Mới Về Trong Ngày
                        f"=IFERROR(VLOOKUP(B{r}; 'Cơ cấu'!C:F; 2; FALSE); \"Chưa phân loại\")", # S: Tỉnh
                        '="NTB"',                                                           # T: Vùng
                        f"=IFERROR(VLOOKUP(B{r}; 'Cơ cấu'!C:F; 4; FALSE); \"Chưa phân loại\")"  # U: AM
                    ]
                    formulas_matrix.append(row_formulas)
                ws.update(range_name=f"J2:U{num_rows}", values=formulas_matrix, value_input_option="USER_ENTERED")
            
            if existing_rows > num_rows:
                print(f"🧹 Xóa công thức dư thừa từ hàng {num_rows + 1} đến {existing_rows}...")
                ws.batch_clear([f"J{num_rows + 1}:U{existing_rows}"])

        print(f"🎉 Tab '{tab_name}': {len(df)} dòng, {len(df.columns)} cột — cập nhật thành công!")
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
            tmp = pd.read_csv(filepath, encoding=enc, sep=sep).fillna("")
            if len(tmp.columns) > 1:
                print(f"✅ Đọc CSV: encoding={enc}, sep='{sep}', {len(tmp)} dòng, {len(tmp.columns)} cột.")
                return tmp
        except Exception:
            continue
    return None

def find_section_anchor(report_frame, ghn_page, keywords):
    """
    Tìm element chứa heading của 1 section theo danh sách keyword.
    Trả về element nếu tìm thấy, None nếu không.
    """
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
    """
    Trong số tất cả element khớp selector, lọc ra những cái có
    bounding box nằm DƯỚI anchor (tức y > anchor_bottom).
    Trả về list element đã sort theo y tăng dần.
    """
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
            if box and box['y'] >= anchor_bottom - 10:  # -10 để có margin nhỏ
                below.append((box['y'], el))
        except Exception:
            continue

    below.sort(key=lambda x: x[0])
    return [el for _, el in below]

def export_table(ghn_page, report_frame, label, section_keywords, filter_btn_index_fallback=0, table_index_fallback=0, expand_details=False, filter_tts=False):
    """
    Xuất bảng từ Looker Studio.
    Dùng section_keywords để xác định vị trí section đúng,
    sau đó tìm button Chi tiết và table-body nằm DƯỚI heading đó.
    """
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
    chi_tiet_btn.click(force=True)  # click bằng force để tránh bị đè
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
            print("   🖱️ Click outside to close dropdown.", flush=True)
    except Exception:
        pass
    ghn_page.keyboard.press("Escape")
    
    # Đợi bảng tải lại dữ liệu sau khi đổi bộ lọc (8 giây)...
    print("⏳ Chờ bảng tải lại dữ liệu sau khi đổi bộ lọc (8 giây)...", flush=True)
    ghn_page.wait_for_timeout(8000)

    # ---- Bước 3b: Click filter Loại khách hàng → chọn duy nhất TTS (nếu filter_tts=True) ----
    if filter_tts:
        print("🖱️ Click 'Loại khách hàng'...", flush=True)
        lk_below = get_elements_below_anchor(
            report_frame, anchor,
            'button:has-text("Loại khách hàng")',
            ghn_page
        )
        if lk_below:
            lk_btn = lk_below[0]
            print("✅ Tìm thấy button 'Loại khách hàng' dưới heading.", flush=True)
        else:
            print("⚠️ Không tìm thấy button 'Loại khách hàng' theo vị trí, dùng fallback first.", flush=True)
            lk_btn = report_frame.locator('button:has-text("Loại khách hàng")').first

        lk_btn.scroll_into_view_if_needed()
        ghn_page.wait_for_timeout(500)
        lk_btn.click(force=True)
        ghn_page.wait_for_timeout(2500)

        print("👉 Xử lý chọn chỉ duy nhất 'TTS' trong dropdown...", flush=True)
        tts_opt = None
        all_opts = report_frame.locator('div.item, div.item-single, div.row, div[role="option"]').all()
        for opt in all_opts:
            try:
                if opt.is_visible():
                    text = opt.inner_text().strip()
                    first_line = text.split('\n')[0].strip()
                    if first_line == "TTS":
                        tts_opt = opt
                        break
            except Exception:
                continue

        if tts_opt:
            print("   🖱️ Hover và click chọn 'Chỉ' (Only) cho 'TTS'...", flush=True)
            tts_opt.hover()
            ghn_page.wait_for_timeout(1000)
            
            # Tìm nút 'Chỉ' (class="only") hoặc bằng text
            only_btn = tts_opt.locator('span.only, .only').first
            if only_btn.count() == 0:
                only_btn = tts_opt.locator('text="Chỉ"').first
            if only_btn.count() == 0:
                only_btn = tts_opt.locator('text="Only"').first
                
            if only_btn.count() > 0:
                only_btn.click(force=True)
                print("   ✅ Đã click button 'Chỉ' (Only).", flush=True)
            else:
                print("   ⚠️ Không tìm thấy button 'Chỉ', click trực tiếp...", flush=True)
                tts_opt.evaluate("el => el.click()")
            ghn_page.wait_for_timeout(1500)
        else:
            print("   ❌ Không tìm thấy option 'TTS'!", flush=True)

        try:
            if anchor:
                anchor.click(force=True, timeout=3000)
        except Exception:
            pass
        ghn_page.keyboard.press("Escape")
        
        print("⏳ Chờ bảng tải lại dữ liệu sau khi lọc 'TTS' (8 giây)...", flush=True)
        ghn_page.wait_for_timeout(8000)

    # ---- Bước 4: Click [+] mở rộng cột Chi tiết (nếu expand_details=True) ----
    if expand_details:
        print("➕ Bấm [+] mở rộng bảng...", flush=True)
        
        hc = None
        # Thử lấy phần tử header cell và click, hồi phục nếu frame reload
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
                # Cập nhật lại anchor
                anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

        print("⏳ Chờ bảng tải lại dữ liệu sau khi bấm [+] (8 giây)...", flush=True)
        ghn_page.wait_for_timeout(8000)
    else:
        print("⏭️ Bỏ qua bước bấm [+] (Giữ nguyên 8 cột để xuất báo cáo B).", flush=True)

    # ---- Bước 5: Right-click table header cell → Export ----
    print("🖱️ Click chuột phải vào header cell 'Chi tiết' của bảng...", flush=True)
    
    hc_target = None
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
                hc_target = header_cells_below[0] if header_cells_below else None
            else:
                hc_target = None

            if not hc_target:
                all_hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').all()
                if all_hc:
                    hc_target = all_hc[filter_btn_index_fallback] if len(all_hc) > filter_btn_index_fallback else all_hc[0]
                else:
                    hc_target = report_frame.locator('div.header-cell:has-text("Chi tiết")').first

            hc_target.scroll_into_view_if_needed(timeout=5000)
            ghn_page.wait_for_timeout(500)
            hc_target.dispatch_event("contextmenu")
            ghn_page.wait_for_timeout(2000)
            break
        except Exception as e:
            print(f"⚠️ Thử {attempt+1}: Lỗi chuột phải header cell ({e}). Đang lấy lại frame...", flush=True)
            ghn_page.wait_for_timeout(2000)
            report_frame = get_looker_frame_helper(ghn_page)
            anchor = find_section_anchor(report_frame, ghn_page, section_keywords)

    print("👉 Chọn 'Export chart...'...", flush=True)
    export_option = report_frame.locator(
        '//*[contains(text(), "Export chart") or contains(text(), "Xuất biểu đồ")]'
    ).first
    export_option.wait_for(state="visible", timeout=10000)
    export_option.hover()
    ghn_page.wait_for_timeout(1500)

    print("👉 Chọn 'Export data'...", flush=True)
    export_data = report_frame.locator(
        '//*[contains(text(), "Export data") or contains(text(), "Xuất dữ liệu")]'
    ).first
    export_data.wait_for(state="visible", timeout=5000)
    export_data.click(force=True)
    ghn_page.wait_for_timeout(3000)

    # ---- Bước 6: Tick Keep value formatting ----
    print("👉 Tích 'Keep value formatting'...", flush=True)
    try:
        keep_cb = report_frame.locator(
            'mat-checkbox:has-text("Keep value formatting"), mat-checkbox:has-text("Giữ định dạng giá trị")'
        ).first
        keep_cb.wait_for(state="visible", timeout=5000)
        if not keep_cb.locator('input[type="checkbox"]').is_checked():
            keep_cb.click()
            print("✅ Đã tích 'Keep value formatting'.", flush=True)
        else:
            print("ℹ️ Checkbox đã tích sẵn.", flush=True)
    except Exception as e:
        print(f"⚠️ Không tick được checkbox: {e}", flush=True)

    # ---- Bước 7: Bấm Export → tải file ----
    print("👉 Bấm Export để tải file...", flush=True)
    confirm_btn = report_frame.locator('button:has-text("Export"), button:has-text("Xuất")').first

    with ghn_page.expect_download() as dl_info:
        confirm_btn.click()

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
            df = pd.read_excel(local_path).fillna("")
        except Exception as ex:
            print(f"❌ Lỗi đọc Excel: {ex}", flush=True)

    try:
        os.remove(local_path)
        print("🧹 Đã xóa file tạm.", flush=True)
    except Exception:
        pass

    return df


def run_job():
    print(f"\n🚀 BẮT ĐẦU CHẠY AUTO DOWNLOAD TTS LÚC: {time.strftime('%H:%M:%S')}", flush=True)
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
        
        # Thiết lập thời gian chờ mặc định là 3 phút để không bị timeout khi chờ đăng nhập
        context.set_default_timeout(180000)
        ghn_page.set_default_timeout(180000)

        # Chờ đăng nhập nếu cần thiết
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

            # Kiểm tra và chuyển trang nếu đang ở MỤC LỤC
            page_name = ""
            try:
                pn = report_frame.locator("span.pageName, .pageName").first
                pn.wait_for(state="attached", timeout=10000)
                page_name = pn.text_content(timeout=3000).strip()
                print(f"📊 Trang hiện tại: '{page_name}'", flush=True)
            except Exception as e:
                print(f"⚠️ Không đọc được tên trang: {e}", flush=True)

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
            # BÁO CÁO B (TTS ONLY) → rawtts
            # ══════════════════════════════════════════
            df2 = export_table(
                ghn_page, report_frame,
                label="B. GIAO TRONG NGÀY (TTS ONLY)",
                section_keywords=[
                    "B. BÁO CÁO GIAO TRONG NGÀY",
                    "BÁO CÁO GIAO TRONG NGÀY",
                    "B. GIAO TRONG NGÀY",
                    "GIAO TRONG NGÀY",
                ],
                filter_btn_index_fallback=1,
                table_index_fallback=1,
                expand_details=True,  # Mở rộng [+]
                filter_tts=True       # Chỉ chọn loại khách hàng TTS
            )
            if df2 is not None:
                upload_to_sheet_tab(df2, GOOGLE_SHEET_KEY, TAB_BAO_CAO_2)
            else:
                print("❌ Không xuất được báo cáo B (TTS).", flush=True)

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


if __name__ == '__main__':
    run_job()
