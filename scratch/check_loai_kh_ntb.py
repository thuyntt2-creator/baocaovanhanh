import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

FILE_PATH = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch\Rớt LC 22_6_2026 - Full sàn.xlsx"

def main():
    if not os.path.exists(FILE_PATH):
        print(f"File not found: {FILE_PATH}")
        return
        
    df = pd.read_excel(FILE_PATH, sheet_name="Sheet1")
    print("Unique vung_xuat values:")
    print(df['vung_xuat'].unique())
    print("\nUnique tinh_xuat values:")
    print(df['tinh_xuat'].unique())
    
    # Check if any row matches vung_xuat == 'NTB' or tinh_xuat == 'NTB'
    ntb_vung = df[df['vung_xuat'] == 'NTB']
    print(f"\nRows with vung_xuat == 'NTB': {len(ntb_vung)}")
    
    ntb_tinh = df[df['tinh_xuat'] == 'NTB']
    print(f"Rows with tinh_xuat == 'NTB': {len(ntb_tinh)}")
    
    # If not NTB, what does it look like? Let's check where vung_xuat starts with NTB
    ntb_starts = df[df['vung_xuat'].astype(str).str.startswith('NTB')]
    print(f"Rows with vung_xuat starting with NTB: {len(ntb_starts)}")
    
    ntb_starts_tinh = df[df['tinh_xuat'].astype(str).str.startswith('NTB')]
    print(f"Rows with tinh_xuat starting with NTB: {len(ntb_starts_tinh)}")

if __name__ == "__main__":
    main()
