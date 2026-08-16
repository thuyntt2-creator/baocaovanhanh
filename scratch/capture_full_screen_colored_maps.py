import asyncio
import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
    os.makedirs(out_dir, exist_ok=True)

    targets = [
        ("map_whatif_nha_trang", "Nha Trang", "Nha Trang", 13.2, "Bản đồ Quy hoạch Phường Nha Trang"),
        ("map_whatif_tay_nha_trang", "Tây Nha Trang", "Tây Nha Trang", 13.0, "Bản đồ Quy hoạch Phường Tây Nha Trang"),
        ("map_whatif_cam_linh", "Cam Linh", "Cam Linh", 12.8, "Bản đồ Quy hoạch Cam Linh"),
        ("map_whatif_ninh_hoa", "Ninh Hòa", "Ninh Hòa", 13.0, "Bản đồ Quy hoạch Ninh Hòa"),
        ("map_whatif_don_duong", "Nghĩa Đức", "Đơn Dương", 12.2, "Bản đồ Quy hoạch Đơn Dương"),
        ("map_whatif_da_lat", "Xuân Hương", "Xuân Hương", 13.0, "Bản đồ Quy hoạch Đà Lạt"),
        ("map_whatif_bao_loc", "Lao", "Lộc", 13.0, "Bản đồ Quy hoạch Bảo Lộc"),
        ("map_whatif_phan_rang", "Phan Rang", "Phan Rang", 13.0, "Bản đồ Quy hoạch Phan Rang"),
        ("map_whatif_phan_thiet", "Hàm Thắng", "Thuận", 13.0, "Bản đồ Quy hoạch Phan Thiết"),
        ("map_whatif_nam_thanh", "Đồng Kho", "Thành", 12.0, "Bản đồ Quy hoạch Nam Thành"),
        ("map_whatif_gia_nghia", "Gia Nghĩa", "Gia Nghĩa", 12.5, "Bản đồ Quy hoạch Gia Nghĩa")
    ]

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
            
        await page.wait_for_timeout(6000)
        
        # Collapse left sidebar controls so map takes 100% full screen
        await page.evaluate("""
        () => {
            let ctrl = document.getElementById("controls");
            if (ctrl) ctrl.classList.add("collapsed");
            if (window.setRegionFilter) setRegionFilter('NTB');
            ui.colormode = 'Theo lãnh thổ BC';
            if (window.draw) draw();
        }
        """)
        await page.wait_for_timeout(3000)
        
        for fname, hub_term, ward_term, zoom_lvl, desc in targets:
            print(f"Generating full-screen colored map for: {desc}...")
            
            js_code = f"""
            () => {{
                let ctrl = document.getElementById("controls");
                if (ctrl) ctrl.classList.add("collapsed");
                
                if (window.setRegionFilter) setRegionFilter('NTB');
                ui.colormode = 'Theo lãnh thổ BC';
                ui.planMode = true;
                
                let hterm = {json.dumps(hub_term.lower())};
                let wterm = {json.dumps(ward_term.lower())};
                
                let targetHub = DATA.hubs.find(h => h.name.toLowerCase().includes(hterm) && h.assigned);
                if (!targetHub) targetHub = DATA.hubs.find(h => h.name.toLowerCase().includes(hterm));
                
                if (targetHub) {{
                    plan.target = targetHub.code;
                    let wards = DATA.wards.filter(w => w.name.toLowerCase().includes(wterm) || w.delivery_hub === targetHub.code).slice(0, 3);
                    for (let w of wards) {{
                        plan.selected.set(w.ward_code, true);
                    }}
                    renderPlan();
                    if (window.map) map.flyTo({{ center: [targetHub.lng, targetHub.lat], zoom: {zoom_lvl} }});
                    draw();
                }}
            }}
            """
            await page.evaluate(js_code)
            await page.wait_for_timeout(3500)
            
            shot_path = os.path.join(out_dir, f"{fname}.png")
            await page.screenshot(path=shot_path)
            print(f"  ✓ Saved full-screen colored map: {fname}.png")
            
        await browser.close()
        print("\nAll full-screen colored maps captured successfully!")

if __name__ == "__main__":
    asyncio.run(main())
