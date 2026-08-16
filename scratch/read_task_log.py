import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
log_path = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\76aa8370-a76b-4d53-a8b2-34d362cf3e4d\.system_generated\tasks\task-188.log"

if os.path.exists(log_path):
    with open(log_path, 'r', encoding='utf-8') as f:
        print(f.read()[-3000:]) # print last 3000 chars
else:
    print("Log not found yet.")
