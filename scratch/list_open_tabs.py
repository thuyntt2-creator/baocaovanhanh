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
        print(f"Open tabs count: {len(context.pages)}")
        for idx, page in enumerate(context.pages):
            try:
                print(f"Tab {idx}: Title='{await page.title()}' URL='{page.url}'")
            except Exception as e:
                print(f"Tab {idx}: error: {e}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
