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
        
        # Test enabling planMode and rendering What-If card for (KHO) Nha Trang
        print("Testing Plan Mode (What-If)...")
        await page.evaluate("""
        () => {
            if (window.setRegionFilter) setRegionFilter('NTB');
            ui.colormode = 'Theo lãnh thổ BC';
            ui.planMode = true;
            
            // Find (KHO) Nha Trang and covered wards
            let nt = DATA.hubs.find(h => h.name.includes('Nha Trang') && h.name.includes('(KHO)'));
            if (nt) {
                plan.target = nt.code;
                // Select 2 wards
                let wards = DATA.wards.filter(w => w.delivery_hub === nt.code || w.pick_hub === nt.code).slice(0, 3);
                for (let w of wards) {
                    plan.selected.set(w.ward_code, true);
                }
                renderPlan();
                if (window.map) map.flyTo({ center: [nt.lng, nt.lat], zoom: 13 });
                draw();
            }
        }
        """)
        await page.wait_for_timeout(3500)
        
        # Take screenshot of map with floating What-If card!
        shot_path = os.path.join(out_dir, "08_whatif_card_on_map.png")
        await page.screenshot(path=shot_path)
        print("Saved 08_whatif_card_on_map.png!")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
