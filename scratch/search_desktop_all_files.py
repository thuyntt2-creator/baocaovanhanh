import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    desktop_dir = r"c:\Users\lap4all\Desktop\Backlog_Automation"
    if not os.path.exists(desktop_dir):
        print("Desktop dir does not exist")
        return
        
    for fname in os.listdir(desktop_dir):
        fpath = os.path.join(desktop_dir, fname)
        if os.path.isfile(fpath):
            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            except Exception:
                continue
            if 'rawgtc' in content.lower():
                print(f"Found 'rawgtc' in file: {fname}")
                lines = content.splitlines()
                for idx, line in enumerate(lines):
                    if 'rawgtc' in line.lower():
                        print(f"  Line {idx+1}: {line.strip()[:120]}")

if __name__ == "__main__":
    main()
