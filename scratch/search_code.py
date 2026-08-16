import os
import sys

def main():
    target_file = r"c:\Users\lap4all\Documents\Auto report\calculate_report_gan.py"
    if not os.path.exists(target_file):
        print("File does not exist")
        return
        
    encodings = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'latin-1']
    content = None
    for enc in encodings:
        try:
            with open(target_file, 'r', encoding=enc) as f:
                content = f.read()
                print(f"Successfully read with encoding: {enc}")
                break
        except Exception:
            pass
            
    if not content:
        print("Failed to read file with any encoding")
        return
        
    queries = ['config_gan', 'sheet', 'worksheet', 'sh.', 'open_by_key', '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4', '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk']
    for q in queries:
        count = content.lower().count(q.lower())
        print(f"Query '{q}': found {count} times")
        if count > 0:
            # print lines containing the query
            lines = content.splitlines()
            for idx, line in enumerate(lines):
                if q.lower() in line.lower():
                    print(f"  Line {idx+1}: {line.strip()[:120]}")

if __name__ == "__main__":
    main()
