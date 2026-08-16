import os
import sys
import io
import json

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

TRANSCRIPT_PATH = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\8f8dc48d-1cb2-45ee-abb9-440faae7062c\.system_generated\logs\transcript.jsonl"

def main():
    if not os.path.exists(TRANSCRIPT_PATH):
        print(f"Transcript path not found: {TRANSCRIPT_PATH}")
        return
        
    print(f"Searching transcript at {TRANSCRIPT_PATH}...")
    keywords = ["nóng", "nong", "tăng vọt", "tang vot", "biến động", "alert", "cảnh báo"]
    
    with open(TRANSCRIPT_PATH, "r", encoding="utf-8") as f:
        for idx, line in enumerate(f):
            try:
                data = json.loads(line)
                content = str(data.get("content", ""))
                type_ = data.get("type", "")
                
                # Check for keywords
                found = [kw for kw in keywords if kw.lower() in content.lower()]
                if found:
                    print(f"\n--- Line {idx} (Type: {type_}) ---")
                    print(f"Matched keywords: {found}")
                    # Print part of content
                    if len(content) > 500:
                        print(content[:500] + "\n...(truncated)...")
                    else:
                        print(content)
            except Exception as e:
                print(f"Error parsing line {idx}: {e}")

if __name__ == "__main__":
    main()
