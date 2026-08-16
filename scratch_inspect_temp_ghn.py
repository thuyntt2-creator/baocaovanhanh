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
        
    print("Reading temp_ghn.xlsx...")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        ws = wb.active
        print(f"Active Sheet: {ws.title}")
        for i, row in enumerate(ws.iter_rows(max_row=30, values_only=True), 1):
            clean_row = [x for x in row if x is not None and x != '']
            if clean_row:
                print(f"Row {i:2d}: {clean_row[:15]}")
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
