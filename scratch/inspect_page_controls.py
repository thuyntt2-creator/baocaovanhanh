import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = r"c:\Users\lap4all\Documents\Auto report"
DASHBOARD_URL = "https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5"

def run():
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            context = browser.contexts[0]
            ghn_page = None
            for pg in context.pages:
                if "baocao.ghn.vn" in pg.url:
                    ghn_page = pg
                    break
            if not ghn_page:
                ghn_page = context.new_page()
                ghn_page.goto(DASHBOARD_URL)
        except Exception:
            profile_dir = os.path.join(SCRIPT_DIR, "playwright_profile")
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                channel="msedge",
                headless=False,
                args=["--start-maximized"]
            )
            ghn_page = context.new_page()
            ghn_page.goto(DASHBOARD_URL)

        ghn_page.wait_for_timeout(3000)
        
        def get_looker_frame():
            for f in ghn_page.frames:
                if f != ghn_page.main_frame and ("lookerstudio" in f.url or "datastudio" in f.url or "google" in f.url):
                    return f
            return ghn_page.main_frame

        report_frame = get_looker_frame()
        print("Frame:", report_frame.url[:80])

        # Navigate to report page if on MỤC LỤC
        try:
            pn = report_frame.locator("span.pageName, .pageName").first
            pname = pn.text_content(timeout=3000).strip()
            print("Current Page:", pname)
            if "MỤC LỤC" in pname or not pname:
                next_btn = report_frame.locator("span.nextBtn, .nextBtn").first
                next_btn.click(force=True)
                ghn_page.wait_for_timeout(5000)
        except Exception as e:
            print("Page check err:", e)

        report_frame = get_looker_frame()

        # Find all controls / filters on page
        controls = report_frame.locator('.control-element, .lego-control, div[role="button"], button').all()
        print(f"Found {len(controls)} potential controls:")
        for idx, c in enumerate(controls):
            try:
                if c.is_visible():
                    txt = c.inner_text().strip().replace('\n', ' ')
                    if txt and len(txt) < 100:
                        print(f"  #{idx}: {repr(txt)}")
            except Exception:
                pass

        # Save screenshot
        ss_path = os.path.join(SCRIPT_DIR, "scratch", "page_controls.png")
        ghn_page.screenshot(path=ss_path, full_page=True)
        print("Screenshot saved to:", ss_path)

if __name__ == '__main__':
    run()
