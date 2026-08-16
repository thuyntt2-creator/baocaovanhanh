import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

FILE_PATH = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch\data.xlsx"

def main():
    if not os.path.exists(FILE_PATH):
        print(f"File not found: {FILE_PATH}")
        return
        
    xl = pd.ExcelFile(FILE_PATH)
    print("Sheets in data.xlsx:")
    for s in xl.sheet_names:
        print(f"- {s}")
        try:
            df = pd.read_excel(xl, s, nrows=2)
            print(f"  Columns: {df.columns.tolist()}")
        except Exception as e:
            print(f"  Error reading: {e}")

if __name__ == "__main__":
    main()
