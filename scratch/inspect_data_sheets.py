import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
output_path = os.path.join(BASE_DIR, "scratch", "inspect_data_sheets_output.txt")

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def log(msg):
    with open(output_path, "a", encoding="utf-8") as out:
        out.write(msg + "\n")

# Clear file
with open(output_path, "w", encoding="utf-8") as out:
    out.write("Data Sheets Detailed Inspection\n\n")

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    
    # 1. Main Spreadsheet
    sh_main = gc_client.open_by_key("1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ")
    
    # Check 'Data' sheet values
    log("--- Sheet 'Data' ---")
    ws_data = sh_main.worksheet("Data")
    rows = ws_data.get_all_values()
    df_data = pd.DataFrame(rows[1:], columns=rows[0])
    log(f"Unique 'Loại Hàng' in Data: {df_data['Loại Hàng'].unique().tolist()}")
    log(f"Time column samples in Data: {df_data['Time'].unique()[-5:].tolist()}")
    log(f"Shape of Data: {df_data.shape}")
    
    # Check 'TTS' sheet values
    log("\n--- Sheet 'TTS' ---")
    ws_tts = sh_main.worksheet("TTS")
    rows_tts = ws_tts.get_all_values()
    df_tts = pd.DataFrame(rows_tts[1:], columns=rows_tts[0])
    log(f"Unique 'Loại Hàng' in TTS: {df_tts['Loại Hàng'].unique().tolist()}")
    log(f"Shape of TTS: {df_tts.shape}")

    # Check 'FD ' sheet values
    log("\n--- Sheet 'FD ' (first 15 rows) ---")
    ws_fd = sh_main.worksheet("FD ")
    fd_rows = ws_fd.get_all_values()
    for idx, r in enumerate(fd_rows[:20]):
        log(f"Row {idx+1}: {r[:12]}")
        
    # Check 'Aging trên 5 ngày' values
    log("\n--- Sheet 'Aging trên 5 ngày' ---")
    ws_aging = sh_main.worksheet("Aging trên 5 ngày")
    aging_rows = ws_aging.get_all_values()
    df_aging = pd.DataFrame(aging_rows[1:], columns=aging_rows[0])
    log(f"Headers of Aging: {df_aging.columns.tolist()}")
    log(f"Sample Row 2: {df_aging.iloc[0].to_dict() if len(df_aging) > 0 else 'Empty'}")
    log(f"Aging categories count: {df_aging['Nhóm BL'].value_counts().to_dict() if 'Nhóm BL' in df_aging.columns else 'N/A'}")
    log(f"Total Aging Rows: {len(df_aging)}")
    
    # Check 'Treo LC' values
    log("\n--- Sheet 'Treo LC' ---")
    ws_treo = sh_main.worksheet("Treo LC")
    treo_rows = ws_treo.get_all_values()
    df_treo = pd.DataFrame(treo_rows[1:], columns=treo_rows[0])
    log(f"Headers of Treo LC: {df_treo.columns.tolist()}")
    log(f"Sample Row 2: {df_treo.iloc[0].to_dict() if len(df_treo) > 0 else 'Empty'}")
    log(f"Total Treo LC Rows: {len(df_treo)}")

    # Check 'OPR' values
    log("\n--- Sheet 'OPR' ---")
    ws_opr = sh_main.worksheet("OPR")
    opr_rows = ws_opr.get_all_values()
    df_opr = pd.DataFrame(opr_rows[1:], columns=opr_rows[0])
    log(f"Headers of OPR: {df_opr.columns.tolist()}")
    log(f"Sample Row 2: {df_opr.iloc[0].to_dict() if len(df_opr) > 0 else 'Empty'}")
    log(f"Unique NgayLTC in OPR: {df_opr['NgayLTC'].unique()[-5:].tolist() if 'NgayLTC' in df_opr.columns else 'N/A'}")
    log(f"Total OPR Rows: {len(df_opr)}")
    
    # 2. Rot LC Spreadsheet
    sh_rot = gc_client.open_by_key("14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4")
    log("\n--- Sheet 'Data TTS' in Rot LC Spreadsheet ---")
    ws_rot_tts = sh_rot.worksheet("Data TTS")
    rot_tts_rows = ws_rot_tts.get_all_values()
    df_rot_tts = pd.DataFrame(rot_tts_rows[1:], columns=rot_tts_rows[0])
    log(f"Unique 'Loại ngày' in Data TTS: {df_rot_tts['Loại ngày'].unique()[-5:].tolist() if 'Loại ngày' in df_rot_tts.columns else 'N/A'}")
    log(f"Sample Row 2: {df_rot_tts.iloc[0].to_dict() if len(df_rot_tts) > 0 else 'Empty'}")
    log(f"Shape of Data TTS: {df_rot_tts.shape}")

if __name__ == "__main__":
    main()
