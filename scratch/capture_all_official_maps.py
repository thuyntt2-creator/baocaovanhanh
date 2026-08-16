import asyncio
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    
    out_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
    os.makedirs(out_dir, exist_ok=True)

    # Locations to zoom / center on MapLibre map: (lng, lat, zoom)
    locations = [
        ("01_nha_trang_khanh_hoa", 109.19, 12.24, 12.5, "Bản đồ Quy hoạch Cụm Nha Trang - Khánh Hòa"),
        ("02_cam_ranh_khanh_hoa", 109.14, 11.91, 12.0, "Bản đồ Quy hoạch Cụm Cam Ranh & Nam Cam Ranh"),
        ("03_van_ninh_khanh_hoa", 109.22, 12.69, 11.8, "Bản đồ Quy hoạch Vạn Ninh - Tu Bông"),
        ("04_don_duong_lam_dong", 108.55, 11.82, 12.0, "Bản đồ Phân vùng Bưu cục Đơn Dương - Lạc Xuân"),
        ("05_da_lat_lam_dong", 108.44, 11.94, 12.8, "Bản đồ Quy hoạch TP. Đà Lạt - Phường Lâm Viên & Xuân Hương"),
        ("06_bao_loc_lam_dong", 107.80, 11.54, 12.5, "Bản đồ Quy hoạch TP. Bảo Lộc - BC B'Lao Mới"),
        ("07_phan_rang_ninh_thuan", 109.00, 11.56, 12.8, "Bản đồ Mở mới Bưu cục Đông Hải - Phan Rang"),
        ("08_phan_thiet_binh_thuan", 108.10, 10.93, 12.8, "Bản đồ Quy hoạch TP. Phan Thiết - Phường Bình Thuận"),
        ("09_tanh_linh_binh_thuan", 107.65, 11.12, 11.8, "Bản đồ Mở mới Bưu cục Nam Thành - Bình Thuận"),
        ("10_gia_nghia_dak_nong", 107.69, 12.00, 12.2, "Bản đồ Quy hoạch TP. Gia Nghĩa - Đắc Nông")
    ]

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1920, "height": 1080})
        page = await context.new_page()
        
        print("Navigating to quyhoachbuucuc.info ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        # Log in
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(4000)
        
        # Switch colormode to 'Theo lãnh thổ BC' or 'Theo vùng' if selector available
        try:
            await page.select_option("#colormode", value="Theo lãnh thổ BC")
        except Exception as e:
            print("Color mode select note:", e)
            
        # Hide any control overlays or panels if needed for clean map view
        await page.wait_for_timeout(2000)
        
        # Iterate over location target coordinates and flyTo in MapLibre map instance
        for fname, lng, lat, zoom, label in locations:
            print(f"Zooming map to {label} ({lng}, {lat}, z={zoom})...")
            
            # Execute JS to flyTo MapLibre map instance if available
            fly_script = f"""
            if (window.map) {{
                window.map.flyTo({{ center: [{lng}, {lat}], zoom: {zoom}, duration: 1000 }});
            }}
            """
            await page.evaluate(fly_script)
            await page.wait_for_timeout(3500) # wait for map tiles to load completely
            
            shot_path = os.path.join(out_dir, f"{fname}.png")
            await page.screenshot(path=shot_path)
            print(f"  ✓ Saved screenshot: {fname}.png")
            
        await browser.close()
        print("\nAll official maps captured successfully!")

if __name__ == "__main__":
    asyncio.run(main())
