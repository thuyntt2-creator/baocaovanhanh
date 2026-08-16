import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r"c:\Users\lap4all\Documents\Auto report\scratch\app_js_full.js", "r", encoding="utf-8") as f:
    code = f.read()

print("Full app.js size:", len(code))

# Search for "Quy hoạch", "what-if", "lưu đề xuất", "reassign", "prop-", "clickWard", "clickHub"
keywords = ["what-if", "gán cho", "Định biên", "Lưu đề xuất", "showDetail", "clickHub", "clickWard", "clickNewWard", "togglePlan", "buildProposalHTML"]

for kw in keywords:
    matches = [m.start() for m in re.finditer(re.escape(kw), code, re.IGNORECASE)]
    print(f"Keyword '{kw}': {len(matches)} occurrences")
    for pos in matches[:3]:
        snippet = code[max(0, pos-100):min(len(code), pos+200)].replace('\n', ' ')
        print(f"  [{pos}] -> {snippet}")
    print()
