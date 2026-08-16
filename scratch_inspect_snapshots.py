import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r"c:\Users\lap4all\Documents\Auto report\snapshot_aging.json", "r", encoding="utf-8") as f:
    state = json.load(f)

print("Daily snapshots:")
for k, v in sorted(state.get("daily_snapshots", {}).items()):
    print(f"Date: {k}, grandTotal: {v.get('grandTotal')}")

print("\nHistory length:", len(state.get("history", [])))
for i, item in enumerate(state.get("history", [])):
    print(f"History index {i}: Time: {item.get('time')}, sum total: {sum(item.get('totals', {}).values())}")
