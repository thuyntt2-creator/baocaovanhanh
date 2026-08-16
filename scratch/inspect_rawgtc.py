import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    target_file = r"c:\Users\lap4all\Desktop\Backlog_Automation\rawgtc.py"
    if not os.path.exists(target_file):
        print("File does not exist")
        return
        
    with open(target_file, 'r', encoding='utf-8') as f:
        content = f.read()
        
    lines = content.splitlines()
    for idx, line in enumerate(lines):
        if 'worksheet' in line.lower() or 'key' in line.lower() or 'update' in line.lower() or 'clear' in line.lower():
            print(f"Line {idx+1}: {line.strip()[:120]}")

if __name__ == "__main__":
    main()
