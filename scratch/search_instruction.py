import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\lap4all\Documents\Auto report\scratch\instruction_readable.txt"

with open(path, 'r', encoding='utf-8') as f:
    text = f.read()

# Let's find some keywords: 'tách', 'tỷ lệ', 'tỷ trọng', 'tháng 9', 'T9', 'T10', 'Forecast'
# We will print matching blocks of text.

keywords = ["tách", "tỷ lệ", "tỷ trọng", "tháng 9", "Forecast T9", "Forecast T8"]
for kw in keywords:
    print(f"\n=== Searching for: {kw} ===")
    matches = [m.start() for m in re.finditer(re.escape(kw), text, re.IGNORECASE)]
    print(f"Found {len(matches)} occurrences.")
    for idx, pos in enumerate(matches[:5]): # print first 5 matches
        start = max(0, pos - 150)
        end = min(len(text), pos + 250)
        print(f"Match {idx+1} at index {pos}:\n{text[start:end]}\n{'-'*40}")
