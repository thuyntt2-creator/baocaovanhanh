import os
import sys
import io
import pandas as pd

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    # Let's find any excel file in the Desktop folder
    desktop_dir = r"c:\Users\lap4all\Desktop\Backlog_Automation"
    excel_file = os.path.join(desktop_dir, "temp_ghn.xlsx")
    if not os.path.exists(excel_file):
        print("Excel file not found")
        return
        
    print(f"Reading {excel_file}...")
    try:
        # Let's read first 10 rows
        df = pd.read_excel(excel_file).fillna("")
        print(f"Columns: {df.columns.tolist()}")
        print(f"Types before cleaning:\n{df.dtypes}")
        print("\nFirst 3 rows:")
        print(df.head(3))
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
