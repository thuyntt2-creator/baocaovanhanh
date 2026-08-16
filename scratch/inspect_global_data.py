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
        
        # Print keys of window.DATA
        data_keys = await page.evaluate("""
        () => {
            if (!window.DATA) return 'NO DATA';
            let info = {};
            for (let k in DATA) {
                if (Array.isArray(DATA[k])) {
                    info[k] = `Array(${DATA[k].length})`;
                } else if (typeof DATA[k] === 'object') {
                    info[k] = `Object(${Object.keys(DATA[k]).length} keys)`;
                } else {
                    info[k] = DATA[k];
                }
            }
            return info;
        }
        """)
        print("DATA keys:", data_keys)
        
        # Sample hubs in DATA.hubs
        hubs_sample = await page.evaluate("""
        () => {
            if (!window.DATA || !DATA.hubs) return [];
            return DATA.hubs.slice(0, 5).map(h => ({
                name: h.name, code: h.code, region: h.region, _region: h._region, province: h.province, province_name: h.province_name, lat: h.lat, lng: h.lng
            }));
        }
        """)
        print("\nHubs sample:", hubs_sample)
        
        # Sample regions in DATA
        regions_sample = await page.evaluate("""
        () => {
            if (!window.DATA || !DATA.hubs) return [];
            let regions = new Set(DATA.hubs.map(h => h.region || h._region || h.tinh));
            return Array.from(regions);
        }
        """)
        print("\nUnique regions in hubs:", regions_sample[:20])
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
