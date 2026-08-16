import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
files = [f for f in os.listdir(downloads_dir) if "AOP_Hang_NTB" in f]

for f in files:
    path = os.path.join(downloads_dir, f)
    mtime = os.path.getmtime(path)
    size = os.path.getsize(path)
    print(f"File: {f} | Size: {size} | Modified: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime))}")

