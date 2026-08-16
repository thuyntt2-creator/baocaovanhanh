import os
import sys
import io
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
        
    try:
        df = pd.read_excel(excel_path, sheet_name='DB')
        # Check rows where Cấp Quản Lý is null, empty or whitespace
        df['Cấp Quản Lý_str'] = df['Cấp Quản Lý'].astype(str).str.strip()
        empty_rows = df[df['Cấp Quản Lý_str'].isin(['', 'nan', 'None'])]
        print(f"Found {len(empty_rows)} rows with empty 'Cấp Quản Lý'")
        if len(empty_rows) > 0:
            print("Sample empty rows:")
            print(empty_rows[['Cấp Quản Lý', 'Chi tiết', 'Loại Hàng', 'Time', 'Volume']].head(20).to_string())
            print("\nUnique 'Chi tiết' in these empty rows:")
            print(empty_rows['Chi tiết'].value_counts())
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
