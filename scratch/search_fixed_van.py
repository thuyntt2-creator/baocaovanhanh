import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
workspace_dir = r"c:\Users\lap4all\Documents\Auto report"

print("=== TÌM KIẾM FILE CHỨA 'fixed_van' HOẶC 'van' ===")

def search_dir(dpath):
    print(f"\nThư mục: {dpath}")
    if not os.path.exists(dpath):
        print("  Không tồn tại")
        return
    for root, dirs, files in os.walk(dpath):
        for f in files:
            flc = f.lower()
            if 'van' in flc or 'fixed' in flc:
                print(f"  - {os.path.join(root, f)}")

search_dir(downloads_dir)
search_dir(workspace_dir)
