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
            print(f"  Rows: {len(df)}")
            if len(df) > 0:
                time_col = 'Time' if 'Time' in df.columns else df.columns[3]
                unique_times = df[time_col].dropna().unique()
                print(f"  Unique times: {list(unique_times)[:5]}")
                print(f"  Latest times: {list(unique_times)[-5:]}")
                vol_col = 'Volume' if 'Volume' in df.columns else df.columns[4]
                df[vol_col] = pd.to_numeric(df[vol_col], errors='coerce').fillna(0)
                print(f"  Total Volume: {df[vol_col].sum()}")
                
                # Check for 2026-06-25
                df_25 = df[df[time_col].astype(str).str.startswith('2026-06-25')]
                print(f"  Rows for 2026-06-25: {len(df_25)}")
                print(f"  Volume for 2026-06-25: {df_25[vol_col].sum()}")
                if len(df_25) > 0:
                    print("  Sample rows for 2026-06-25:")
                    print(df_25[['Cấp Quản Lý', 'Chi tiết', time_col, vol_col]].head(10).to_string())
        except Exception as e:
            print(f"  Error reading sheet '{sname}': {e}")

if __name__ == "__main__":
    main()
