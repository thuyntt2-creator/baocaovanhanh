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
        if fname.endswith(".py"):
            fpath = os.path.join(desktop_dir, fname)
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception:
                try:
                    with open(fpath, 'r', encoding='latin-1') as f:
                        content = f.read()
                except Exception:
                    continue
                    
            if 'raw' in content.lower():
                lines = content.splitlines()
                for idx, line in enumerate(lines):
                    if 'worksheet(' in line.lower() and 'raw' in line.lower():
                        print(f"File {fname} Line {idx+1}: {line.strip()}")
                    elif '.worksheet' in line.lower() and 'raw' in line.lower():
                        print(f"File {fname} Line {idx+1}: {line.strip()}")

if __name__ == "__main__":
    main()
