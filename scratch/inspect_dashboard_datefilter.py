import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"

def inspect_dashboard():
    with sync_playwright() as p:
        profile_dir = os.path.join(SCRIPT_DIR, "playwright_profile")
        os.makedirs(profile_dir, exist_ok=True)
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            channel="msedge",
            headless=False,
            args=["--start-maximized"]
        )
        ghn_page = context.new_page()
        print(f"👉 Opening: {DASHBOARD_URL}")
        ghn_page.goto(DASHBOARD_URL)
        
        time.sleep(5)
        
        def get_looker_frame():
            for f in ghn_page.frames:
                if f != ghn_page.main_frame and ("lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url):
                    return f
            return ghn_page.main_frame

        report_frame = get_looker_frame()
        print("Frame URL:", report_frame.url[:100] if report_frame.url else "N/A")
        
        # Check current page name
        try:
            pn = report_frame.locator("span.pageName, .pageName").first
            page_name = pn.text_content(timeout=5000).strip()
            print("Page Name:", page_name)
            if "MỤC LỤC" in page_name or not page_name:
                next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                next_btn.click(force=True)
                time.sleep(5)
        except Exception as e:
            print("Error checking page name:", e)

        # Re-get frame
        report_frame = get_looker_frame()

        # Check date range controls on the frame / page
        date_controls = report_frame.locator('.date-range-picker, .date-range, input[type="text"], .control-element, .lego-control').all()
        print(f"Found {len(date_controls)} potential control elements.")
        for i, c in enumerate(date_controls):
            try:
                txt = c.inner_text().strip()
                if txt:
                    print(f"Control #{i}: {repr(txt)}")
            except Exception:
                pass
                
        # Take screenshot of page
        screenshot_path = os.path.join(SCRIPT_DIR, "dashboard_screenshot.png")
        ghn_page.screenshot(path=screenshot_path, full_page=True)
        print("Screenshot saved to:", screenshot_path)

        context.close()

if __name__ == '__main__':
    inspect_dashboard()
