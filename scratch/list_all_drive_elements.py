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
        page = await context.new_page()
        
        url = "https://drive.google.com/drive/u/0/folders/19FQ9tbV_OrNH7qT7sfQklBwit_Q1CPm2"
        print(f"Navigating to: {url}")
        await page.goto(url)
        await page.wait_for_timeout(5000)
        
        rows = await page.locator("div[role='row']").all()
        print(f"Found {len(rows)} rows.")
        for idx, row in enumerate(rows):
            text = await row.inner_text()
            text = text.replace("\n", " ")
            print(f"Row {idx}: {text}")
            
        await page.close()
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
