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

# Re-use string helpers from original script
def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip().lower())

def clean_bc_name(name):
    name = normalize_str(name)
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    name = name.replace("-", " ").replace("_", " ")
    return " ".join(name.split())

def match_po_name(raw_name, standard_list):
    raw_norm = normalize_str(raw_name)
    for std in standard_list:
        if normalize_str(std) == raw_norm:
            return std, "Exact Match"
            
    raw_clean = clean_bc_name(raw_name)
    if not raw_clean:
        return None, None
        
    cleaned_std_list = [(std, clean_bc_name(std)) for std in standard_list]
    
    for std, std_clean in cleaned_std_list:
        if std_clean == raw_clean:
            return std, "Clean Exact Match"
            
    matches = []
    for std, std_clean in cleaned_std_list:
        if std_clean and (raw_clean in std_clean or std_clean in raw_clean):
            matches.append(std)
            
    if matches:
        matches.sort(key=lambda x: abs(len(clean_bc_name(x)) - len(raw_clean)))
        return matches[0], "Fuzzy Substring Match"
        
    return None, None

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
    
    # Load FD sheet rows
    ws_fd = sh.worksheet("FD ")
    fd_rows = ws_fd.get_all_values()
    fd_headers = ['Bưu Cục', 'AM', '%FD (N)', '%FD (N-1)', 'vs N-1', '%FD (N-7)', 'vs N-7', 'Vol giao', 'Vol trả', 'Tỷ trọng trả']
    df_fd = pd.DataFrame(fd_rows[3:], columns=fd_headers + [f'Col_{i}' for i in range(10, len(fd_rows[3]))])
    
    print("Debugging FD sheet name matching for matching errors...")
    for idx, row in df_fd.iterrows():
        raw_name = row['Bưu Cục']
        if not raw_name or raw_name.strip() == "":
            continue
        matched_po, match_type = match_po_name(raw_name, master_pos)
        
        # We only care about transit centers showing up or matching results containing "Kho" or "Trung Chuyển" or "Chuyển Tiếp"
        if matched_po and any(x in matched_po for x in ["Kho", "Trung chuyển", "Chuyển tiếp"]):
            print(f"Row {idx+4}: Raw: '{raw_name}' -> Matched: '{matched_po}' (Type: {match_type})")
            print(f"  • Raw cleaned: '{clean_bc_name(raw_name)}'")
            print(f"  • Matched cleaned: '{clean_bc_name(matched_po)}'")
            print(f"  • Metrics: %FD={row['%FD (N)']}, Vol={row['Vol giao']}")
            
if __name__ == "__main__":
    main()
