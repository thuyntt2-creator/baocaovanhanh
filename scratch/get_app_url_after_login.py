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
            
        await page.wait_for_timeout(4000)
        
        print("Page URL after login:", page.url)
        
        # Print all scripts loaded on the current page after login
        script_sources = await page.evaluate("""
        () => Array.from(document.querySelectorAll('script')).map(s => ({src: s.src, inner: s.innerText.slice(0, 200)}))
        """)
        
        print("\nScripts loaded after login:")
        for s in script_sources:
            if s['src']:
                print("  SRC:", s['src'])
            elif s['inner'].strip():
                print("  INLINE:", s['inner'].strip().replace('\n', ' ')[:100])
                
        # Also print global variables that might hold the map or state
        globals_info = await page.evaluate("""
        () => {
            let res = {};
            for (let k of Object.keys(window)) {
                if (k.includes('map') || k.includes('Map') || k.includes('deck') || k.includes('state') || k.includes('data')) {
                    res[k] = typeof window[k];
                }
            }
            return res;
        }
        """)
        print("\nGlobals matching map/state/deck:", globals_info)
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
