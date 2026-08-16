import os
import io
import sys
import time
from playwright.sync_api import sync_playwright
import pandas as pd

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"

def test_download():
    print("🌐 Khởi chạy Edge (Persistent)...")
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
            print(f"👉 Mở trang báo cáo: {DASHBOARD_URL}")
            ghn_page.goto(DASHBOARD_URL)
            
            # Wait for login
            ghn_page.wait_for_selector("iframe", timeout=120000)
            print("✅ Đăng nhập thành công!")
            
            def get_looker_frame():
                for f in ghn_page.frames:
                    if f != ghn_page.main_frame and (
                        "lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url
                    ):
                        return f
                return ghn_page.main_frame

            report_frame = get_looker_frame()
            
            # Click Next Page if at MỤC LỤC
            page_name = ""
            try:
                pn = report_frame.locator("span.pageName, .pageName").first
                pn.wait_for(state="attached", timeout=10000)
                page_name = pn.text_content(timeout=3000).strip()
            except Exception:
                pass
                
            if "MỤC LỤC" in page_name or not page_name:
                print("📑 Chuyển sang trang báo cáo (Click Next Page)...")
                next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                next_btn.click(force=True)
                ghn_page.wait_for_timeout(5000)
                report_frame = get_looker_frame()
            
            # Find the header cell 'Chi tiết' below anchor
            anchor = report_frame.locator('text="GIAO TRONG NGÀY"').first
            if anchor.count() > 0:
                anchor.scroll_into_view_if_needed()
                ghn_page.wait_for_timeout(1000)
            
            # Right click the header cell to open menu
            hc = report_frame.locator('div.header-cell:has-text("Chi tiết")').first
            hc.scroll_into_view_if_needed()
            hc.click(button="right")
            ghn_page.wait_for_timeout(1500)
            
            # Find menu item for Export
            export_menu = None
            for kw in ["Export data", "Xuất dữ liệu"]:
                loc = report_frame.locator(f'[role="menuitem"]:has-text("{kw}"), .mat-menu-item:has-text("{kw}"), .goog-menuitem:has-text("{kw}")').first
                if loc.count() > 0 and loc.is_visible():
                    export_menu = loc
                    break
            
            if not export_menu:
                # If hover submenu is needed
                loc_parent = report_frame.locator('[role="menuitem"]:has-text("Xuất"), .mat-menu-item:has-text("Xuất"), .goog-menuitem:has-text("Xuất")').first
                if loc_parent.count() > 0 and loc_parent.is_visible():
                    loc_parent.hover()
                    ghn_page.wait_for_timeout(1000)
                    for sub_kw in ["Export data", "Xuất dữ liệu"]:
                        loc = report_frame.locator(f'[role="menuitem"]:has-text("{sub_kw}"), .mat-menu-item:has-text("{sub_kw}"), .goog-menuitem:has-text("{sub_kw}")').first
                        if loc.count() > 0 and loc.is_visible():
                            export_menu = loc
                            break
            
            if not export_menu:
                print("❌ Không tìm thấy menu Export.")
                context.close()
                return
                
            print(f"✅ Click menu Export: '{export_menu.inner_text().strip()}'")
            export_menu.click(force=True)
            ghn_page.wait_for_timeout(3000)
            
            # Open format dropdown
            format_select = report_frame.locator('mat-select').first
            format_select.click()
            ghn_page.wait_for_timeout(1000)
            
            # Click option Excel
            options = report_frame.locator('mat-option').all()
            if not options:
                options = ghn_page.locator('mat-option').all()
            
            excel_opt = None
            for opt in options:
                text = opt.inner_text().strip().lower()
                if "excel" in text or "xlsx" in text:
                    excel_opt = opt
                    break
            
            if excel_opt:
                excel_opt.click()
                print("✅ Đã chuyển định dạng sang Excel (xlsx).")
                ghn_page.wait_for_timeout(1000)
            
            # Tick keep value formatting if not ticked
            keep_cb = report_frame.locator(
                'mat-checkbox:has-text("Keep value formatting"), mat-checkbox:has-text("Giữ định dạng giá trị")'
            ).first
            is_checked = False
            try:
                inner_input = keep_cb.locator('input')
                if inner_input.count() > 0:
                    is_checked = inner_input.is_checked()
                else:
                    aria_checked = keep_cb.get_attribute("aria-checked")
                    is_checked = (aria_checked == "true")
            except Exception:
                pass
                
            if not is_checked:
                keep_cb.click()
                print("✅ Đã tích Keep value formatting.")
                ghn_page.wait_for_timeout(1000)
                
            # Click Export to download
            confirm_btn = report_frame.locator('mat-dialog-container button:has-text("Export"), mat-dialog-container button:has-text("Xuất"), button:has-text("Export"), button:has-text("Xuất")').first
            with ghn_page.expect_download() as dl_info:
                confirm_btn.click()
                
            download = dl_info.value
            fname = download.suggested_filename
            print(f"📥 Đã tải file: {fname}")
            local_path = os.path.join(r"C:\Users\lap4all\Downloads", fname)
            download.save_as(local_path)
            
            context.close()
            
            # Read and print Cam Linh rows
            df = pd.read_excel(local_path).fillna("")
            print("\n--- IN-MEMORY DATA IN DOWNLOADED EXCEL ---")
            print(df.head(5))
            
            print("\nRows for 'Cam Linh':")
            for idx, r in df.iterrows():
                if "Cam Linh" in str(r.iloc[1]):
                    print(list(r))
                    
            # os.remove(local_path)
            print("\n🧹 Giữ lại file Excel tại Downloads để phân tích.")
            
        except Exception as e:
            print(f"❌ Lỗi: {e}")

if __name__ == "__main__":
    test_download()
