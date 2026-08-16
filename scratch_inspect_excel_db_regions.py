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
        
    print("Reading manual_raw.xlsx...")
    for sname in ['DB', 'Data']:
        try:
            df = pd.read_excel(excel_path, sheet_name=sname)
            print(f"\n=== Sheet '{sname}' in Excel ===")
            print(f"Unique 'Cấp Quản Lý' in {sname}:")
            print(df['Cấp Quản Lý'].value_counts() if 'Cấp Quản Lý' in df.columns else "No 'Cấp Quản Lý'")
            print(f"Unique 'Chi tiết' count in {sname}:")
            print(df['Chi tiết'].nunique() if 'Chi tiết' in df.columns else "No 'Chi tiết'")
        except Exception as e:
            print(f"Error reading {sname}: {e}")

if __name__ == "__main__":
    main()
