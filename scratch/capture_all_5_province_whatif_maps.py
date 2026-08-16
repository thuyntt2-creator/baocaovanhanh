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
        
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(5000)
        
        # Set region to NTB
        await page.evaluate("() => { if(window.setRegionFilter) setRegionFilter('NTB'); }")
        await page.wait_for_timeout(2000)
        
        # Check rezone new_wards in DATA.rez
        res = await page.evaluate("""
        () => {
            let wards = [];
            if (window.DATA && DATA.rez && DATA.rez.new_wards) {
                wards = DATA.rez.new_wards.map(w => ({ name: w.name, code: w.code, region: w.region, province: w.province, bc: w.bc_name || w.bc }));
            }
            let hubs = [];
            if (window.DATA && DATA.hubs) {
                hubs = DATA.hubs.map(h => ({ name: h.name, code: h.code, region: h.region || h._region, province: h.province_name }));
            }
            return { wards: wards.slice(0, 40), hubs: hubs.filter(h => h.region === 'NTB').slice(0, 40) };
        }
        """)
        
        print(f"Found {len(res['wards'])} wards and {len(res['hubs'])} hubs for NTB.")
        for w in res['wards'][:15]:
            print(f"  Ward: {w['name']} ({w['code']}) - Province: {w.get('province')}")
            
        for h in res['hubs'][:15]:
            print(f"  Hub: {h['name']} ({h['code']}) - Province: {h.get('province')}")
            
        # Switch tab to Re-zone to see the list of new wards and click them
        await page.evaluate("() => switchTab('rezone')")
        await page.wait_for_timeout(2000)
        
        await page.screenshot(path=os.path.join(out_dir, "04_rezone_tab.png"), full_page=True)
        print("Captured 04_rezone_tab.png")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
