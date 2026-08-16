"""
Export cookies từ session trình duyệt hiện tại để dùng trên GitHub Actions.
Chạy script này trên máy local khi đã đăng nhập baocao.ghn.vn.
Output: cookies_ghn.json → copy nội dung vào GitHub Secret GHN_COOKIES
"""
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "cookies_ghn.json")

def main():
    print("🔌 Kết nối Chrome debug (127.0.0.1:9222)...")
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        except Exception:
            print("⚠️ Không kết nối được Chrome debug. Thử persistent profile...")
            profile_dir = os.path.join(SCRIPT_DIR, "playwright_profile")
            if not os.path.exists(profile_dir):
                print("❌ Không tìm thấy profile. Hãy chạy download script trước để đăng nhập.")
                return
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                channel="msedge",
                headless=False,
                args=["--start-maximized"]
            )
            page = context.new_page()
            page.goto("https://baocao.ghn.vn/dashboards/63bd175cd4435a369fade8f5")
            page.wait_for_timeout(5000)
            
            storage = context.storage_state()
            with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
                json.dump(storage, f, ensure_ascii=False, indent=2)
            
            print(f"✅ Đã export cookies → {OUTPUT_FILE}")
            print(f"📋 Copy nội dung file này vào GitHub Secret: GHN_COOKIES")
            context.close()
            return

        if len(browser.contexts) == 0:
            print("❌ Không tìm thấy context.")
            return
        
        context = browser.contexts[0]
        storage = context.storage_state()
        
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(storage, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Đã export cookies → {OUTPUT_FILE}")
        print(f"📋 Copy nội dung file này vào GitHub Secret: GHN_COOKIES")
        browser.close()

if __name__ == "__main__":
    main()
