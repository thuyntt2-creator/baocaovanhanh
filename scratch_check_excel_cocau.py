import os
import sys
import io
import openpyxl
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    excel_path = os.path.join(BASE_DIR, 'manual_raw.xlsx')
    if not os.path.exists(excel_path):
        print("manual_raw.xlsx not found.")
        return
        
    print("Reading manual_raw.xlsx 'Cơ cấu' sheet...")
    try:
        df_cc = pd.read_excel(excel_path, sheet_name='Cơ cấu')
        print(f"Cơ cấu row count in Excel: {len(df_cc)}")
        print("Cơ cấu columns in Excel:", df_cc.columns.tolist())
        print("Cơ cấu unique provinces in Excel:")
        print(df_cc['Tỉnh'].value_counts() if 'Tỉnh' in df_cc.columns else "No 'Tỉnh' column")
        print("\nFirst 10 rows in Excel:")
        print(df_cc.head(10).to_string())
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
