import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
    os.makedirs(out_dir, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1920, "height": 1080})
        page = await context.new_page()
        
        print("Navigating to https://www.quyhoachbuucuc.info/web/index.html ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(4000)
        
        # Take full page screenshot of the map interface
        await page.screenshot(path=os.path.join(out_dir, "01_map_dashboard_full.png"))
        print("Captured 01_map_dashboard_full.png")
        
        # Inspect JS objects and map instance name
        js_inspect = """
        () => {
            let info = {};
            info.keys = Object.keys(window).filter(k => !k.startswith || !k.startswith('webkit'));
            info.has_map = 'map' in window;
            info.map_type = window.map ? (window.map.constructor ? window.map.constructor.name : typeof window.map) : 'none';
            info.selects = Array.from(document.querySelectorAll('select')).map(s => ({id: s.id, options: Array.from(s.options).map(o => o.text)}));
            info.buttons = Array.from(document.querySelectorAll('button')).map(b => b.innerText);
            return info;
        }
        """
        res = await page.evaluate(js_inspect)
        print("JS inspection result:")
        print("  map_type:", res.get("map_type"))
        print("  selects:", res.get("selects"))
        print("  buttons:", res.get("buttons")[:15])
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
