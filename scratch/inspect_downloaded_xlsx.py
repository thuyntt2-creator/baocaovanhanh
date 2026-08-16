import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

FILE_PATH = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch\Rớt LC 22_6_2026 - Full sàn.xlsx"

def main():
    if not os.path.exists(FILE_PATH):
        print(f"File not found: {FILE_PATH}")
        return
        
    xl = pd.ExcelFile(FILE_PATH)
    print("Sheets in Excel:")
    for s in xl.sheet_names:
        print(f"- {s}")
        df = pd.read_excel(xl, s, nrows=5)
        print(f"  Columns: {df.columns.tolist()}")
        print(f"  First rows:\n{df.head(2)}")

if __name__ == "__main__":
    main()
