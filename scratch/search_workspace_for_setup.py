import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

search_dir = r"c:\Users\lap4all\Documents\Auto report"
terms = ["setup", "di dời", "di doi", "tiện ích", "utilities"]

print("=== Searching workspace files ===")
for root, dirs, files in os.walk(search_dir):
    # skip pycache and playwrigt
    if "__pycache__" in root or "playwright" in root or ".gemini" in root or ".git" in root:
        continue
    for f in files:
        if f.endswith(('.py', '.md', '.txt', '.json')):
            fpath = os.path.join(root, f)
            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as file:
                    content = file.read().lower()
                    for t in terms:
                        if t in content:
                            print(f"Found '{t}' in {os.path.relpath(fpath, search_dir)}")
                            break
            except Exception as e:
                pass
print("Completed.")
