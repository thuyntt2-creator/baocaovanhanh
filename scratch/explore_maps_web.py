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
        
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        # Log in
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(3000)
        
        # Dump text/links/buttons/selects on the page
        print("Page URL after login:", page.url)
        content = await page.content()
        
        # Print headings, buttons, selects
        buttons = await page.query_selector_all("button, a, select, option, div[role='button']")
        print(f"Found {len(buttons)} interactive elements.")
        
        for idx, btn in enumerate(buttons[:30]):
            txt = await btn.text_content()
            tag = await btn.evaluate("el => el.tagName")
            id_val = await btn.get_attribute("id")
            cls_val = await btn.get_attribute("class")
            print(f"  {idx:2d}. <{tag}> id='{id_val}' class='{cls_val}' text='{txt.strip()}'")
            
        # Capture current screen after login
        await page.screenshot(path=os.path.join(out_dir, "dashboard.png"), full_page=True)
        print("Captured dashboard.png")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
