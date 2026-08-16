import os
import sys
import json
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
INSPECT_JSON = os.path.join(BASE_DIR, "scratch", "sheet_data_inspect.json")

def main():
    with open(INSPECT_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    odr = pd.DataFrame(data["ODR"]["rows"], columns=data["ODR"]["headers"])
    
    print("Value counts for 'Quản lý':")
    print(odr["Quản lý"].value_counts())
    
    print("\nValue counts for 'Chi tiết':")
    print(odr["Chi tiết"].value_counts())
    
    print("\nMin/Max Time:")
    print(odr["Time"].min(), "to", odr["Time"].max())
    
    print("\nFirst 20 rows of ODR:")
    print(odr.head(20).to_string())

if __name__ == "__main__":
    main()
