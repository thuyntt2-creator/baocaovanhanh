import sys
import io
import os
import pandas as pd
import unicodedata

sys.stdout.reconfigure(encoding='utf-8')

FILE_PATH = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch\Rớt LC 22_6_2026 - Full sàn.xlsx"

def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())

def main():
    if not os.path.exists(FILE_PATH):
        print(f"File not found: {FILE_PATH}")
        return
        
    df = pd.read_excel(FILE_PATH, sheet_name="Sheet1")
    df_ntb = df[df['vung_xuat'] == 'NTB'].copy()
    
    print(f"Total NTB rows in raw Excel: {len(df_ntb)}")
    print("Unique loai_kh values in NTB:")
    print(df_ntb['loai_kh'].value_counts())
    
    # Let's count by post office for all NTB rows
    po_counts_all = df_ntb['tenbcxuat'].value_counts()
    print("\nTop 10 post offices by ALL rows in NTB:")
    print(po_counts_all.head(10))
    
    # Let's count by post office for TTS NTB rows
    df_tts = df_ntb[df_ntb['loai_kh'] == 'TTS']
    po_counts_tts = df_tts['tenbcxuat'].value_counts()
    print("\nTop 10 post offices by TTS rows in NTB:")
    print(po_counts_tts.head(10))
    
if __name__ == "__main__":
    main()
