import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    with open(r"c:\Users\lap4all\Documents\Auto report\snapshot_push.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    print("snapshot_push.json contents:")
    print(json.dumps(data, ensure_ascii=False, indent=2))
except Exception as e:
    print("Error:", e)
