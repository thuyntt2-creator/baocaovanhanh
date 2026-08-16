import os
import sys
import io

# Fix encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    target_file = r"c:\Users\lap4all\Documents\Auto report\calculate_report_gan.py"
    if not os.path.exists(target_file):
        print("File does not exist")
        return
        
    with open(target_file, 'r', encoding='utf-8') as f:
        content = f.read()
        print("Successfully read calculate_report_gan.py")
        
    queries = ['config_gan', 'sheet', 'worksheet', 'sh.', 'open_by_key', '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4', '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk', 'raw']
    for q in queries:
        count = content.lower().count(q.lower())
        print(f"Query '{q}': found {count} times")
        if count > 0:
            lines = content.splitlines()
            printed = 0
            for idx, line in enumerate(lines):
                if q.lower() in line.lower():
                    print(f"  Line {idx+1}: {line.strip()[:120]}")
                    printed += 1
                    if printed >= 15:
                        print("  ... more matches hidden ...")
                        break

if __name__ == "__main__":
    main()
