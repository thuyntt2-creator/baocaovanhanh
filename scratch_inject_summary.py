import json

file_path = r"c:\Users\lap4all\Documents\Auto report\snapshot_push.json"
with open(file_path, "r", encoding="utf-8") as f:
    data = json.load(f)

if len(data.get("history", [])) > 0:
    data["history"][0]["summary"] = {
        "grand_unassigned": 510,
        "grand_assigned": 950,
        "grand_processed": 180,
        "grand_active": 1460
    }
    
with open(file_path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("Moked summary injected successfully into the first history entry.")
