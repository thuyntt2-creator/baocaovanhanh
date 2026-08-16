import sys, json
sys.stdout.reconfigure(encoding='utf-8')
with open('snapshot_aging.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=== snapshot_aging.json history ===")
for idx, item in enumerate(data.get('history', [])):
    time_val = item.get("timestamp", item.get("time"))
    totals = item.get("totals", {})
    total_val = sum(totals.values()) if isinstance(totals, dict) else 0
    print(f"Item {idx}: time = {time_val}, grandTotal = {total_val}")
    print("Totals:", totals)
