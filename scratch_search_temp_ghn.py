import os
import sys
import io
import openpyxl

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    excel_path = os.path.join(BASE_DIR, 'temp_ghn.xlsx')
    if not os.path.exists(excel_path):
        print("temp_ghn.xlsx not found.")
        return
        
    target_str = '21624000'
    target_num = 21624000
    
    print(f"Scanning temp_ghn.xlsx for '{target_str}'...")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        ws = wb.active
        found = False
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                if val == target_str or val == target_num:
                    found = True
                    print(f"  🎯 Found at Row {r_idx}, Col {c_idx}! Row content: {list(row)}")
        if not found:
            print("  Not found in temp_ghn.xlsx.")
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
