import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})
        
        print("Navigating to https://www.quyhoachbuucuc.info/web/index.html ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        # Fill credentials
        inputs = await page.query_selector_all("input")
        print("Found inputs:", len(inputs))
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            
            # Look for login button specifically
            login_btn = await page.query_selector("button:has-text('Đăng nhập'), input[type='submit'], .btn-login, button.btn")
            if login_btn:
                print("Clicking login button...")
                await login_btn.click()
            else:
                print("Pressing Enter on password input...")
                await inputs[1].press("Enter")
                
        # Wait for navigation or data loading
        await page.wait_for_timeout(6000)
        
        print("Current URL:", page.url)
        await page.screenshot(path=os.path.join(out_dir, "03_after_click_login.png"))
        print("Saved 03_after_click_login.png")
        
        # Check window properties again
        globals_list = await page.evaluate("""
        () => Object.keys(window).filter(k => !k.startsWith('webkit') && !k.startsWith('on'))
        """)
        print("Window globals after 6s:", globals_list)
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
