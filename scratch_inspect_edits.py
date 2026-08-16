import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

log_path = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\78d41a18-f126-407d-afc0-e289534b2d9a\.system_generated\logs\transcript.jsonl"
with open(log_path, 'r', encoding='utf-8') as f:
    for idx, line in enumerate(f):
        data = json.loads(line)
        for call in data.get("tool_calls", []):
            args = call.get("args", {})
            target = args.get("TargetFile") or args.get("Target")
            if target and "update_aging_assignments.py" in target:
                print(f"Step {idx}: tool: {call.get('name')}")
                if "TargetContent" in args:
                    print("--- TARGET CONTENT ---")
                    print(args.get("TargetContent"))
                    print("--- REPLACEMENT CONTENT ---")
                    print(args.get("ReplacementContent"))
                elif "ReplacementChunks" in args:
                    chunks = args.get("ReplacementChunks")
                    if isinstance(chunks, str):
                        try:
                            chunks = json.loads(chunks)
                        except:
                            pass
                    if isinstance(chunks, list):
                        for chunk in chunks:
                            print("--- TARGET CONTENT (CHUNK) ---")
                            print(chunk.get("TargetContent"))
                            print("--- REPLACEMENT CONTENT (CHUNK) ---")
                            print(chunk.get("ReplacementContent"))
                elif "CodeContent" in args:
                    print("--- WRITE FILE CONTENT ---")
                    print(args["CodeContent"][:1000])
                print("=" * 60)
