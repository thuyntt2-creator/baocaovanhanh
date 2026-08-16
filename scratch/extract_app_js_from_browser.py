import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})
        
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(6000)
        
        # Fetch app.js from authenticated browser session
        app_js = await page.evaluate("""
        async () => {
            let resp = await fetch("/web/app.js");
            return await resp.text();
        }
        """)
        
        print("Full app.js size:", len(app_js))
        
        # Save to local file
        out_path = r"c:\Users\lap4all\Documents\Auto report\scratch\app_js_full.js"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(app_js)
            
        print(f"Saved full app.js to: {out_path}")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
