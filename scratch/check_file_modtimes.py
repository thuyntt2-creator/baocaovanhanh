import os
import sys
from datetime import datetime

def check_dir(dpath):
    print(f"\nFiles in {dpath}:")
    if not os.path.exists(dpath):
        print("  Directory does not exist")
        return
    files = []
    for f in os.listdir(dpath):
        fp = os.path.join(dpath, f)
        if os.path.isfile(fp):
            mtime = os.path.getmtime(fp)
            dt = datetime.fromtimestamp(mtime)
            files.append((dt, f, os.path.getsize(fp)))
    files.sort(reverse=True)
    for dt, f, sz in files[:15]:
        print(f"  {dt.strftime('%Y-%m-%d %H:%M:%S')} - {f} ({sz} bytes)")

def main():
    check_dir(r"c:\Users\lap4all\Documents\Auto report")
    check_dir(r"c:\Users\lap4all\Desktop\Backlog_Automation")

if __name__ == "__main__":
    main()
