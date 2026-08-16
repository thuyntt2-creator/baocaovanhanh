import os
import io
import sys
import pandas as pd

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

DOWNLOADS_DIR = r"C:\Users\lap4all\Downloads"

def inspect():
    # Find any excel file matching the pattern in Downloads
    matched_file = None
    for f in os.listdir(DOWNLOADS_DIR):
        if "BÁO CÁO XỬ LÝ THÀNH CÔNG TRONG NGÀY" in f and f.endswith(".xlsx"):
            matched_file = os.path.join(DOWNLOADS_DIR, f)
            print(f"Found file: {f}")
            break
            
    if not matched_file:
        print("❌ Không tìm thấy file Excel nào trong Downloads.")
        return
        
    df = pd.read_excel(matched_file).fillna("")
    print("\nColumns in local Excel file:")
    print(list(df.columns))
    
    print("\nCam Linh 2026-07-04 row values and types in local Excel:")
    for idx, r in df.iterrows():
        if "Cam Linh" in str(r.iloc[1]) and "2026-07-04" in str(r.iloc[3]):
            vals = list(r)
            print("Values:", vals)
            print("Types :", [type(x) for x in vals])

if __name__ == "__main__":
    inspect()
