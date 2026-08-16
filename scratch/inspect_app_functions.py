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
            
        await page.wait_for_timeout(4000)
        
        # Inspect all custom functions and objects in window
        fns_info = await page.evaluate("""
        () => {
            let res = {};
            for (let k of Object.keys(window)) {
                if (typeof window[k] === 'function' && !k.startsWith('on') && !k.startsWith('webkit')) {
                    res[k] = window[k].toString().slice(0, 300);
                }
            }
            return res;
        }
        """)
        
        print("Custom functions in window:")
        for fn, src in fns_info.items():
            print(f"=== {fn} ===")
            print(src.replace('\n', ' '))
            print()
            
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
