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
        
        # Filter region to NTB
        await page.evaluate("() => { if(window.setRegionFilter) setRegionFilter('NTB'); }")
        await page.wait_for_timeout(2000)
        
        # Dump available hubs and new wards in DATA
        data_summary = await page.evaluate("""
        () => {
            let res = { hubs: [], new_wards: [] };
            if (window.DATA) {
                if (DATA.hubs) {
                    res.hubs = DATA.hubs.filter(h => h.type === 'express' && (h.region === 'NTB' || (h._region && h._region.includes('NTB'))))
                                       .map(h => ({ name: h.name, code: h.code, lat: h.lat, lng: h.lng, region: h.region }));
                }
                if (DATA.rez && DATA.rez.new_wards) {
                    res.new_wards = DATA.rez.new_wards.slice(0, 30).map(w => ({ name: w.name, code: w.code, region: w.region }));
                }
            }
            return res;
        }
        """)
        
        print(f"Found {len(data_summary['hubs'])} Express hubs in NTB:")
        for h in data_summary['hubs'][:20]:
            print(f"  BC: {h['name']} ({h['code']}) at [{h['lng']}, {h['lat']}]")
            
        print(f"\nFound {len(data_summary['new_wards'])} new wards in DATA.rez:")
        for w in data_summary['new_wards'][:10]:
            print(f"  Ward: {w['name']} ({w['code']})")
            
        # Capture general map after region filter NTB
        await page.screenshot(path=os.path.join(out_dir, "00_ntb_region_map.png"))
        print("\nCaptured 00_ntb_region_map.png")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
