import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

log_path = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\78d41a18-f126-407d-afc0-e289534b2d9a\.system_generated\logs\transcript.jsonl"
with open(log_path, 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
        data = json.loads(line)
        if data.get("step_index") == 22:
            content = data.get("content")
            # The content has line numbers prefixed like "1: import sys, io\n2: import os\n..."
            lines = content.split('\n')
            original_lines = []
            for l in lines:
                if ":" in l:
                    parts = l.split(":", 1)
                    if parts[0].strip().isdigit():
                        original_lines.append(parts[1][1:]) # skip the space after colon
            
            with open("update_aging_assignments_original.py", "w", encoding="utf-8") as out:
                out.write("\n".join(original_lines))
            print("Successfully extracted original update_aging_assignments.py")
            break
