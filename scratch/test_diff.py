import json
import os

SNAPSHOT_FILE = r"c:\Users\lap4all\Documents\Auto report\snapshot_push.json"
if os.path.exists(SNAPSHOT_FILE):
    with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
        state = json.load(f)
    print("History length:", len(state["history"]))
    for i, item in enumerate(state["history"]):
        print(f"Index {i}: time={item.get('time')}, summary={item.get('summary')}")
    
    first_snap = state["history"][0]
    first_summary = first_snap.get("summary", {})
    print("First summary grand_unassigned:", first_summary.get("grand_unassigned"))
    
    cur_unassigned = 500
    diff_unassigned = cur_unassigned - first_summary.get("grand_unassigned", cur_unassigned)
    print("Computed diff_unassigned (for cur=500):", diff_unassigned)
else:
    print("File not found")
