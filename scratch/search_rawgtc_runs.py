import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def search_files(dpath):
    print(f"\nSearching in: {dpath}")
    if not os.path.exists(dpath):
        return
    for fname in os.listdir(dpath):
        if fname.endswith(('.bat', '.ps1')):
            fpath = os.path.join(dpath, fname)
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception:
                try:
                    with open(fpath, 'r', encoding='latin-1') as f:
                        content = f.read()
                except Exception:
                    continue
            if 'rawgtc' in content.lower():
                print(f"Found in {fname}")
                # print lines containing rawgtc
                lines = content.splitlines()
                for idx, line in enumerate(lines):
                    if 'rawgtc' in line.lower():
                        print(f"  Line {idx+1}: {line.strip()}")

def main():
    search_files(r"c:\Users\lap4all\Documents\Auto report")
    search_files(r"c:\Users\lap4all\Desktop\Backlog_Automation")

if __name__ == "__main__":
    main()
