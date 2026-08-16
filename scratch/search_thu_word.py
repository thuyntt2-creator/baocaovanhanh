import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def search_dir(dpath):
    print(f"\nSearching in: {dpath}")
    if not os.path.exists(dpath):
        return
    for fname in os.listdir(dpath):
        if fname.endswith(".py"):
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
            if 'thứ' in content.lower() or 'chủ nhật' in content.lower() or 'monday' in content.lower():
                # print lines containing thứ/chủ nhật
                lines = content.splitlines()
                for idx, line in enumerate(lines):
                    if any(w in line.lower() for w in ['thứ', 'chủ nhật', 'weekday']):
                        print(f"  {fname} Line {idx+1}: {line.strip()[:120]}")

def main():
    search_dir(r"c:\Users\lap4all\Documents\Auto report")
    search_dir(r"c:\Users\lap4all\Desktop\Backlog_Automation")

if __name__ == "__main__":
    main()
