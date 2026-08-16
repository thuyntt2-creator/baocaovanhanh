import requests
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "https://www.quyhoachbuucuc.info/web/app.js?v=1785502319"
code = requests.get(url).text

print("App.js size:", len(code))

# Search for "what-if", "gán cho BC MỚI", "Định biên", "BC mới gánh cụm này", "Lưu đề xuất"
matches = []
for line in code.split(';'):
    if any(k in line for k in ['what-if', 'gán cho', 'Định biên', 'BC MỚI', 'Lưu đề xuất', 'prop-', 'buildProposalHTML', 'reviewProposal', 'clickHub', 'clickWard']):
        matches.append(line.strip())

print(f"Found {len(matches)} matching lines:")
for m in matches[:15]:
    print("  -", m[:120])
