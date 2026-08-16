import os
import io
import sys
from playwright.sync_api import sync_playwright

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"

def inspect_dialog():
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
                print(f"📊 Trang hiện tại: '{page_name}'")
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
            
            # Now the export dialog is open!
            # Take screenshot of dialog
            ghn_page.screenshot(path="temp_export_dialog.png")
            print("📸 Đã chụp ảnh dialog và lưu tại: temp_export_dialog.png")
            
            # Print all mat-selects
            selects = report_frame.locator('mat-select').all()
            print(f"Found {len(selects)} mat-select elements:")
            for s in selects:
                print(f"  - mat-select: id={s.get_attribute('id')}, value/text='{s.inner_text().strip()}'")
                
            # Print all options/text inside dialog
            dialog = report_frame.locator('mat-dialog-container').first
            if dialog.count() > 0:
                print("\nText in dialog:")
                print(dialog.inner_text())
            else:
                print("❌ Không tìm thấy mat-dialog-container.")
                
            # Close dialog
            ghn_page.keyboard.press("Escape")
            context.close()
            
        except Exception as e:
            print(f"❌ Lỗi: {e}")

if __name__ == "__main__":
    inspect_dialog()
