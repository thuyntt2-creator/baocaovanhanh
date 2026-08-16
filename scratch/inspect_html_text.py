import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

html_path = r"C:\Users\lap4all\Downloads\Huong_dan_chi_tiet_lap_AOP_Hang_Nang.html"
with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
    html = f.read()

# remove tags
text = re.sub('<[^<]+?>', '\n', html)
lines = [l.strip() for l in text.split('\n') if l.strip()]

print(f"Total clean lines: {len(lines)}")
for idx, line in enumerate(lines):
    if any(term in line.lower() for term in ["setup", "mở mới", "mở", "di dời", "dời", "đức linh"]):
        print(f"L{idx}: {line}")
