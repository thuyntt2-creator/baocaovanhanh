import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1920, "height": 1080})
        page = await context.new_page()
        
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(4000)
        
        # Search for map object in global scope or map element
        js_find_map = """
        () => {
            let matches = [];
            for (let k in window) {
                try {
                    let obj = window[k];
                    if (obj && typeof obj === 'object' && (obj.flyTo || obj.panTo || obj.setView || obj.setCenter)) {
                        matches.push({ key: k, type: obj.constructor ? obj.constructor.name : typeof obj });
                    }
                } catch(e) {}
            }
            // Also check document.getElementById('map')
            let mapEl = document.getElementById('map');
            if (mapEl) {
                for (let k in mapEl) {
                    if (k.startsWith('_') || k.includes('map') || k.includes('Map')) {
                        matches.push({ key: 'mapEl.' + k, type: typeof mapEl[k] });
                    }
                }
            }
            return matches;
        }
        """
        matches = await page.evaluate(js_find_map)
        print("Found map instances:", matches)
        
        # Also print scripts loaded on page
        scripts = await page.evaluate("""
        () => Array.from(document.querySelectorAll('script')).map(s => s.src || s.innerText.slice(0, 100))
        """)
        print("\nScripts on page:")
        for s in scripts:
            if s.strip():
                print("  -", s[:120])
                
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
