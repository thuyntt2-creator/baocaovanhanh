import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("Playwright not installed, trying selenium or requests...")
        return

    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
    os.makedirs(out_dir, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False) # run browser so we can see / capture
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})
        
        print("Navigating to https://www.quyhoachbuucuc.info/web/index.html ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        # Take screenshot of login page
        await page.screenshot(path=os.path.join(out_dir, "01_login_page.png"))
        print("Captured 01_login_page.png")
        
        # Look for login form inputs
        # User - Pass: ntb / taghjnxorjvq
        # Let's inspect page content
        content = await page.content()
        print("Page title:", await page.title())
        
        # Fill login inputs if present
        inputs = await page.query_selector_all("input")
        print(f"Found {len(inputs)} input fields.")
        
        # Try filling username and password
        user_input = await page.query_selector("input[type='text'], input[name='user'], input[name='username'], input[placeholder*='user'], input[placeholder*='Tài khoản']")
        pass_input = await page.query_selector("input[type='password'], input[name='pass'], input[name='password']")
        
        if not user_input or not pass_input:
            # Maybe inputs are just the 1st and 2nd input
            if len(inputs) >= 2:
                user_input = inputs[0]
                pass_input = inputs[1]
                
        if user_input and pass_input:
            await user_input.fill("ntb")
            await pass_input.fill("taghjnxorjvq")
            print("Filled credentials!")
            
            # Click submit/login button
            login_btn = await page.query_selector("button, input[type='submit'], .btn-login, #login, button:has-text('Đăng nhập')")
            if login_btn:
                await login_btn.click()
            else:
                await pass_input.press("Enter")
                
            await page.wait_for_timeout(3000)
            await page.screenshot(path=os.path.join(out_dir, "02_after_login.png"))
            print("Captured 02_after_login.png")
        else:
            print("Could not locate username/password inputs automatically.")
            
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
