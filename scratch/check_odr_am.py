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
    cocau = pd.DataFrame(data["Cocau"]["rows"], columns=data["Cocau"]["headers"])
    odr_tts = pd.DataFrame(data["ODR - TTS"]["rows"], columns=data["ODR - TTS"]["headers"])
    
    print("--- ODR Unique 'Quản lý' ---")
    print(odr["Quản lý"].unique())
    
    print("\n--- ODR Unique 'Chi tiết' (First 15) ---")
    print(odr["Chi tiết"].unique()[:15])
    
    print("\n--- Cocau Unique 'AM' ---")
    print(cocau["AM"].unique())
    
    print("\n--- ODR - TTS Unique 'AM' ---")
    print(odr_tts["AM"].unique())
    
    # Check if we can map Chi tiết to Cocau's Bưu cục
    print("\nMapping analysis:")
    unmapped = []
    for po in odr["Chi tiết"].unique():
        # Match po in Cocau
        match = cocau[cocau["Bưu cục"].str.lower() == po.lower()]
        if match.empty:
            unmapped.append(po)
            
    print(f"Total unique POs in ODR: {len(odr['Chi tiết'].unique())}")
    print(f"Unmapped POs count: {len(unmapped)}")
    if unmapped:
        print(f"Unmapped examples: {unmapped[:10]}")

if __name__ == "__main__":
    main()
