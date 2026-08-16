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
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(6000)
        
        # Check DATA contents in detail
        info = await page.evaluate("""
        () => {
            let res = {};
            res.data_keys = window.DATA ? Object.keys(DATA) : [];
            res.region_options = Array.from(document.querySelectorAll('#regionfilter option')).map(o => o.value);
            res.tinh_options = Array.from(document.querySelectorAll('select option')).map(o => o.innerText);
            return res;
        }
        """)
        print("Data info after 6s:", info)
        
        # Click Re-zone tab
        rezone_tab = await page.query_selector("button:has-text('Re-zone')")
        if rezone_tab:
            await rezone_tab.click()
            await page.wait_for_timeout(2000)
            
        # Print table text on Re-zone tab
        rz_text = await page.evaluate("""
        () => {
            let el = document.getElementById('rezone') || document.body;
            return el.innerText.slice(0, 1500);
        }
        """)
        print("\nRezone text content:")
        print(rz_text)
        
        await page.screenshot(path=os.path.join(out_dir, "05_rezone_content.png"))
        print("\nSaved 05_rezone_content.png")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
