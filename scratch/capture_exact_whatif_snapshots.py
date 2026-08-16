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
        
        print("Navigating to quyhoachbuucuc.info ...")
        await page.goto("https://www.quyhoachbuucuc.info/web/index.html", wait_until="networkidle")
        
        inputs = await page.query_selector_all("input")
        if len(inputs) >= 2:
            await inputs[0].fill("ntb")
            await inputs[1].fill("taghjnxorjvq")
            await inputs[1].press("Enter")
            
        await page.wait_for_timeout(6000)
        
        # Switch tab to Re-zone
        rezone_tab = await page.query_selector("button:has-text('Re-zone')")
        if rezone_tab:
            await rezone_tab.click()
            await page.wait_for_timeout(2000)
            
        # Click expand details button
        expand_btn = await page.query_selector(".tbl-body-btn, button:has-text('Xem chi tiết'), button:has-text('36 dòng')")
        if expand_btn:
            print("Expanding 36 rezone rows...")
            await expand_btn.click()
            await page.wait_for_timeout(1000)
            
        # Get all rows in the rezone table
        rows = await page.query_selector_all("tr.region-row, .tbl-body tr, table tr[data-code]")
        print(f"Found {len(rows)} rezone table rows.")
        
        # Capture screenshots of key rows
        for idx, row in enumerate(rows[:20]):
            try:
                row_text = await row.text_content()
                clean_text = row_text.strip().replace('\n', ' ')
                print(f"\nRow {idx+1}: {clean_text[:80]}")
                
                # Click the row to open what-if panel / detail view
                await row.click()
                await page.wait_for_timeout(2000)
                
                # Take screenshot
                safe_name = "".join(c if c.isalnum() else "_" for c in clean_text[:30])
                shot_path = os.path.join(out_dir, f"rezone_row_{idx+1:02d}_{safe_name}.png")
                await page.screenshot(path=shot_path)
                print(f"  ✓ Saved screenshot: rezone_row_{idx+1:02d}_{safe_name}.png")
            except Exception as e:
                print(f"  Error clicking row {idx+1}:", e)
                
        await browser.close()
        print("\nFinished capturing rezone what-if snapshots!")

if __name__ == "__main__":
    asyncio.run(main())
