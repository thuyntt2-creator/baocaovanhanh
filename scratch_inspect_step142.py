import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

log_path = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\78d41a18-f126-407d-afc0-e289534b2d9a\.system_generated\logs\transcript.jsonl"
with open(log_path, 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
        if idx == 142:
            data = json.loads(line)
            print("Step 142:")
            print(json.dumps(data, ensure_ascii=False, indent=2))
