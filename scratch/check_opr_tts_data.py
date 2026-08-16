import os
import sys
import pandas as pd
import gspread
import unicodedata
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def clean_bc_name(name):
    name = unicodedata.normalize('NFC', str(name).strip().lower())
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    return " ".join(name.split())

def match_po_name(raw_name, standard_list):
    raw_norm = unicodedata.normalize('NFC', str(raw_name).strip().lower())
    for std in standard_list:
        if unicodedata.normalize('NFC', std).lower() == raw_norm:
            return std
            
    raw_clean = clean_bc_name(raw_name)
    if not raw_clean:
        return None
        
    cleaned_std_list = [(std, clean_bc_name(std)) for std in standard_list]
    for std, std_clean in cleaned_std_list:
        if std_clean == raw_clean:
            return std
            
    matches = []
    for std, std_clean in cleaned_std_list:
        if std_clean and (raw_clean in std_clean or std_clean in raw_clean):
            matches.append(std)
            
    if matches:
        matches.sort(key=lambda x: abs(len(clean_bc_name(x)) - len(raw_clean)))
        return matches[0]
        
    return None

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(MAIN_SPREADSHEET_ID)
    
    # Load Master pos
    ws_cc = sh.worksheet("CoCauVung")
    cc_rows = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_rows[1:], columns=cc_rows[0])
    df_cc = df_cc[df_cc['Bưu cục'].str.strip() != '']
    master_pos = df_cc['Bưu cục'].str.strip().tolist()
    
    # Load OPR sheet rows
    ws_opr = sh.worksheet("OPR")
    opr_rows = ws_opr.get_all_values()
    df_opr = pd.DataFrame(opr_rows[1:], columns=opr_rows[0])
    df_opr_day = df_opr[df_opr['NgayLTC'] == '2026-06-27']
    
    print(f"Total rows in 'OPR' sheet for 2026-06-27: {len(df_opr_day)}")
    
    po_opr = {}
    for idx, row in df_opr_day.iterrows():
        raw_po = row['KhoLay']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
            
        vol = int(row['Don_ltc']) if row['Don_ltc'] else 0
        ot = int(row['Don_ontime']) if row['Don_ontime'] else 0
        late = max(0, vol - ot)
        
        if std_po not in po_opr:
            po_opr[std_po] = {'vol': 0, 'ontime': 0, 'late': 0}
            
        po_opr[std_po]['vol'] += vol
        po_opr[std_po]['ontime'] += ot
        po_opr[std_po]['late'] += late

    print("\nListing post offices violating OPR TTS thresholds (%OPR < 80% and late > 5):")
    count = 0
    for po, metrics in po_opr.items():
        if metrics['vol'] > 0:
            rate = (metrics['ontime'] / metrics['vol']) * 100.0
            if rate < 80.0 and metrics['late'] > 5:
                count += 1
                print(f"  • {po}: %OPR = {rate:.1f}%, Vol = {metrics['vol']}, Late = {metrics['late']}")
                
    if count == 0:
        print("  • No post offices violated the OPR TTS alert threshold!")

if __name__ == "__main__":
    main()
