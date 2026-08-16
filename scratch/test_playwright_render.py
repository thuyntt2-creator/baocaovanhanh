from playwright.sync_api import sync_playwright
import os

html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: sans-serif; background: #0f172a; color: white; padding: 20px; }
  h1 { color: #38bdf8; }
</style>
</head>
<body>
  <h1>Test Render</h1>
</body>
</html>
"""

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.set_content(html)
    page.screenshot(path="scratch/test_render.png")
    browser.close()

print("Screenshot created! Size:", os.path.getsize("scratch/test_render.png"))
