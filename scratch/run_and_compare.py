import os
import sys
import time
import re
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

# Force stdout encoding to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = r"C:\Users\lap4all\Downloads"
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"
GOOGLE_SHEET_KEY = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"

def get_credentials_path():
    candidates = [
        os.path.join(SCRIPT_DIR, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "credentials.json")
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def run_compare():
    print("🌐 Khởi chạy Edge...")
    with sync_playwright() as p:
        try:
            profile_dir = os.path.join(SCRIPT_DIR, "playwright_profile")
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                channel="msedge",
                headless=False,
                args=["--start-maximized"]
            )
            ghn_page = context.new_page()
            ghn_page.goto(DASHBOARD_URL)
            ghn_page.wait_for_selector("iframe", timeout=60000)
            
            def get_looker_frame():
                for f in ghn_page.frames:
                    if f != ghn_page.main_frame and ("lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url):
                        return f
                return ghn_page.main_frame

            report_frame = get_looker_frame()
            
            # Navigate page
            page_name = ""
            try:
                pn = report_frame.locator("span.pageName, .pageName").first
                pn.wait_for(state="attached", timeout=10000)
                page_name = pn.text_content(timeout=3000).strip()
            except:
                pass
                
            if "MỤC LỤC" in page_name or not page_name:
                next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                next_btn.click(force=True)
                ghn_page.wait_for_timeout(5000)
                report_frame = get_looker_frame()
            
            # Change filter 'Bưu Cục - Tỉnh'
            anchor = report_frame.locator('text="GIAO TRONG NGÀY"').first
            if anchor.count() > 0:
                anchor.scroll_into_view_if_needed()
                ghn_page.wait_for_timeout(1000)
                
            btn_details = report_frame.locator('button:has-text("Chi tiết"), button:has-text("Details")').first
            btn_details.click()
            ghn_page.wait_for_timeout(1000)
            
            opt = report_frame.locator('mat-option:has-text("Bưu Cục - Tỉnh")').first
            opt.click()
            ghn_page.wait_for_timeout(1000)
            ghn_page.keyboard.press("Escape")
            ghn_page.wait_for_timeout(8000)
            
            # Open export menu
            ghn_page.keyboard.press("Escape")
            ghn_page.wait_for_timeout(500)
            
            hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').first
            hc.click(button="right")
            ghn_page.wait_for_timeout(1500)
            
            export_menu = report_frame.locator('[role="menuitem"]:has-text("Xuất"), .mat-menu-item:has-text("Xuất"), .goog-menuitem:has-text("Xuất")').first
            export_menu.hover()
            ghn_page.wait_for_timeout(1000)
            
            sub_el = report_frame.locator('[role="menuitem"]:has-text("Xuất dữ liệu"), .mat-menu-item:has-text("Xuất dữ liệu")').first
            sub_el.click(force=True)
            ghn_page.wait_for_timeout(3000)
            
            # Format select Excel
            format_select = report_frame.locator('mat-select').first
            format_select.click()
            ghn_page.wait_for_timeout(1000)
            excel_opt = report_frame.locator('mat-option:has-text("Excel")').first
            excel_opt.click()
            ghn_page.wait_for_timeout(1000)
            
            # Keep value formatting
            keep_cb = report_frame.locator('mat-checkbox:has-text("Keep value formatting"), mat-checkbox:has-text("Giữ định dạng giá trị")').first
            is_checked = False
            try:
                inner_input = keep_cb.locator('input')
                if inner_input.count() > 0:
                    is_checked = inner_input.is_checked()
                else:
                    aria_checked = keep_cb.get_attribute("aria-checked")
                    is_checked = (aria_checked == "true")
            except:
                pass
            if not is_checked:
                keep_cb.click()
                ghn_page.wait_for_timeout(1000)
                
            # Click export to download
            confirm_btn = report_frame.locator('mat-dialog-container button:has-text("Export"), mat-dialog-container button:has-text("Xuất"), button:has-text("Export"), button:has-text("Xuất")').first
            with ghn_page.expect_download() as dl_info:
                confirm_btn.click()
                
            download = dl_info.value
            fname = download.suggested_filename
            local_path = os.path.join(SCRIPT_DIR, fname)
            download.save_as(local_path)
            context.close()
            
            # Read and print details
            df = pd.read_excel(local_path).fillna("")
            print("\nColumns in Excel file:")
            print(list(df.columns))
            
            rename_map = {
                'Cấp quản lý': 'Cấp Quản Lý',
                'Loại hình': 'Loại Hàng',
                'Sản lượng': 'Volume',
                '% LTC': '% GTC',
                '% Đóng kiện': '% Chuyển trả',
                '% LC': 'Leadtime'
            }
            df.rename(columns=rename_map, inplace=True)
            
            # Print the row for Cam Linh on 2026-07-04 in df
            for idx, r in df.iterrows():
                if "Cam Linh" in str(r.iloc[1]) and "2026-07-04" in str(r.iloc[3]):
                    print("\n[COMPARE IN DATAFRAME] Cam Linh row:")
                    print("Values:", list(r))
                    print("Types :", [type(x) for x in list(r)])
                    
            # Upload to test sheet
            json_path = get_credentials_path()
            creds = Credentials.from_service_account_file(json_path, scopes=scopes)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(GOOGLE_SHEET_KEY)
            
            try:
                ws = sh.worksheet("test_gtc_upload")
                sh.del_worksheet(ws)
            except:
                pass
            ws = sh.add_worksheet(title="test_gtc_upload", rows=2000, cols=30)
            
            data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
            ws.update(data_to_upload, value_input_option='USER_ENTERED')
            print("\nSuccessfully uploaded to test_gtc_upload!")
            
            # Read back from test sheet
            unformatted = ws.row_values(2, value_render_option='UNFORMATTED_VALUE')
            print("\nUnformatted row 2 from test_gtc_upload:", unformatted)
            
        except Exception as e:
            print("❌ Lỗi:", e)

if __name__ == "__main__":
    run_compare()
