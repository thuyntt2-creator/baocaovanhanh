import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

search_dirs = [
    r"C:\Users\lap4all\Downloads",
    r"C:\Users\lap4all\Desktop",
    r"C:\Users\lap4all\Documents",
    os.getcwd()
]

targets = ["AOP_MAU_NTB_v5.xlsx", "AOP_NTB_V2_final.xlsx"]

print("🔍 Searching for target files...")
for s_dir in search_dirs:
    if not os.path.exists(s_dir):
        continue
    print(f"Checking in: {s_dir}...")
    for root, dirs, files in os.walk(s_dir):
        for f in files:
            if f in targets:
                print(f"✨ Found: {os.path.join(root, f)} (size: {os.path.getsize(os.path.join(root, f))} bytes)")
