import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

temp_ghn_path = r"c:\Users\lap4all\Documents\Auto report\temp_ghn.xlsx"
if os.path.exists(temp_ghn_path):
    try:
        df = pd.read_excel(temp_ghn_path, nrows=5)
        print("Columns of temp_ghn.xlsx:")
        print(repr(df.columns.tolist()))
        print("\nFirst 3 rows:")
        for idx, row in df.head(3).iterrows():
            print(repr(row.to_dict()))
    except Exception as e:
        print(f"Error reading temp_ghn.xlsx: {e}")
else:
    print(f"File not found: {temp_ghn_path}")
