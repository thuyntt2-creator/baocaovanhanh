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
        
    df = pd.read_excel(excel_path, sheet_name='DB')
    # Filter for date 2026-06-11
    df_11 = df[df['Time'].astype(str).str.startswith('2026-06-11')]
    print(f"Total rows on 2026-06-11: {len(df_11)}")
    print(df_11[['Cấp Quản Lý', 'Chi tiết', 'Loại Hàng', 'Time', 'Volume']].to_string())

if __name__ == "__main__":
    main()
