import requests
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "https://www.quyhoachbuucuc.info/web/index.html"
resp = requests.get(url)
html = resp.text

print("index.html size:", len(html))
print("\nScript tags in index.html:")
for src in re.findall(r'src=["\']([^"\']+)["\']', html):
    print(" ", src)

print("\nInline JS snippet in index.html (first 500 chars):")
scripts_inline = re.findall(r'<script[^>]*>(.*?)</script>', html, re.DOTALL)
for i, s in enumerate(scripts_inline):
    if not s.strip().startswith('var msg'):
        print(f"--- Script {i} ---")
        print(s[:500])
