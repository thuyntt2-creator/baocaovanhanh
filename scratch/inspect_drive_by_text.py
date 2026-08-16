import sys
import io
import os
import asyncio
from playwright.async_api import async_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        context = browser.contexts[0]
        
        target_page = None
        for pg in context.pages:
            if "drive.google.com" in pg.url and "19FQ9tbV_OrNH7qT7sfQklBwit" in pg.url:
                target_page = pg
                break
                
        if not target_page:
            print("Google Drive page not found among open tabs. Opening new tab...")
            target_page = await context.new_page()
            await target_page.goto("https://drive.google.com/drive/u/0/folders/19FQ9tbV_OrNH7qT7sfQklBwit_Q1CPm2")
            await target_page.wait_for_timeout(8000)
            
        print(f"Connected to Drive tab: URL='{target_page.url}'")
        
        # Locate all elements that contain "Rớt LC"
        elements = await target_page.locator('div:has-text("Rớt LC")').all()
        print(f"Found {len(elements)} elements with 'Rớt LC'.")
        
        # Let's search for span or div containing the name
        spans = await target_page.locator('span').all()
        print(f"Total span elements: {len(spans)}")
        matches = []
        for s in spans:
            try:
                t = await s.inner_text()
                if "Rớt LC" in t:
                    matches.append(t)
            except:
                pass
        print(f"Spans containing 'Rớt LC' ({len(matches)}):")
        for m in set(matches):
            print(f"- '{m}'")
            
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
