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
            print("Google Drive page not found among open tabs.")
            await browser.close()
            return
            
        print(f"Connected to open Drive tab: Title='{await target_page.title()}'")
        
        # List all rows on the page
        rows = await target_page.locator("div[role='row']").all()
        print(f"Found {len(rows)} row elements in open page.")
        for idx, row in enumerate(rows):
            try:
                text = await row.inner_text()
                text = text.replace("\n", " ")
                print(f"Row {idx}: {text}")
            except Exception as e:
                print(f"Row {idx} error: {e}")
                
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
