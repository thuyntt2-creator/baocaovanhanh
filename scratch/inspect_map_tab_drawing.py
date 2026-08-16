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
        
        print("Navigating to quyhoachbuucuc.info ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(5000)
        
        # Ensure we are on the 'Bản đồ' tab
        await page.evaluate("""
        () => {
            if (window.switchTab) switchTab('map');
            if (window.setRegionFilter) setRegionFilter('NTB');
            ui.colormode = 'Theo lãnh thổ BC';
            if (window.draw) draw();
        }
        """)
        await page.wait_for_timeout(3000)
        
        # Take screenshot of the map tab
        await page.screenshot(path=os.path.join(out_dir, "06_map_tab_drawn.png"))
        print("Captured 06_map_tab_drawn.png")
        
        # Let's inspect how to select/click a hub or ward on the MAP tab to show popup
        res = await page.evaluate("""
        () => {
            let info = {};
            if (window.DATA && DATA.hubs) {
                let ntbHubs = DATA.hubs.filter(h => (h.region === 'NTB' || h._region === 'NTB') && h.assigned);
                info.ntbHubs = ntbHubs.map(h => ({ name: h.name, code: h.code, lat: h.lat, lng: h.lng }));
            }
            return info;
        }
        """)
        
        print(f"Found {len(res.get('ntbHubs', []))} assigned Express hubs in NTB.")
        for h in res.get('ntbHubs', [])[:10]:
            print(f"  Hub: {h['name']} ({h['code']}) at [{h['lng']}, {h['lat']}]")
            
        # Try calling clickHub for (KHO) Nha Trang
        print("\nTesting clickHub on (KHO) Nha Trang...")
        await page.evaluate("""
        () => {
            if (window.DATA && DATA.hubs) {
                let nt = DATA.hubs.find(h => h.name.includes('Nha Trang') && h.name.includes('(KHO)'));
                if (nt) {
                    if (window.clickHub) clickHub(nt);
                    if (window.map) map.flyTo({ center: [nt.lng, nt.lat], zoom: 13.5 });
                }
            }
        }
        """)
        await page.wait_for_timeout(3500)
        
        await page.screenshot(path=os.path.join(out_dir, "07_nha_trang_map_clicked.png"))
        print("Captured 07_nha_trang_map_clicked.png")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
