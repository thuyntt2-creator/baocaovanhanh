import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
output_path = os.path.join(BASE_DIR, "scratch", "inspect_po_data_output.txt")

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def log(msg):
    with open(output_path, "a", encoding="utf-8") as out:
        out.write(msg + "\n")

# Clear file
with open(output_path, "w", encoding="utf-8") as out:
    out.write("Post Office Names Alignment & Sample Data for 2026-06-23\n\n")

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    
    # Open spreadsheets
    sh_main = gc_client.open_by_key("1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ")
    sh_rot = gc_client.open_by_key("14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4")
    
    # Load CoCauVung mapping
    ws_cc = sh_main.worksheet("CoCauVung")
    cc_rows = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_rows[1:], columns=cc_rows[0])
    log("--- CoCauVung Sample ---")
    log(df_cc.head(10).to_string())
    
    # Load a few rows from 'Data' sheet for target date 2026-06-23
    ws_data = sh_main.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    df_data_yesterday = df_data[df_data['Time'].str.startswith('2026-06-23')]
    log(f"\n--- Data Sheet on 2026-06-23 ({len(df_data_yesterday)} rows) ---")
    log(df_data_yesterday.head(10).to_string())
    
    # Load 'TTS' sheet for target date 2026-06-23
    ws_tts = sh_main.worksheet("TTS")
    tts_rows = ws_tts.get_all_values()
    df_tts = pd.DataFrame(tts_rows[1:], columns=tts_rows[0])
    df_tts_yesterday = df_tts[df_tts['Time'].str.startswith('2026-06-23')]
    log(f"\n--- TTS Sheet on 2026-06-23 ({len(df_tts_yesterday)} rows) ---")
    log(df_tts_yesterday.head(10).to_string())

    # Load 'FD ' sheet rows (skip first 3 rows headers)
    ws_fd = sh_main.worksheet("FD ")
    fd_rows = ws_fd.get_all_values()
    df_fd = pd.DataFrame(fd_rows[3:], columns=['Bưu Cục', 'AM', '%FD (N)', '%FD (N-1)', 'vs N-1', '%FD (N-7)', 'vs N-7', 'Vol giao', 'Vol trả', 'Tỷ trọng trả'] + [f'C{i}' for i in range(10, len(fd_rows[3]))])
    log(f"\n--- FD Sheet ({len(df_fd)} rows) ---")
    log(df_fd.head(10).to_string())

    # Load 'Aging trên 5 ngày'
    ws_aging = sh_main.worksheet("Aging trên 5 ngày")
    aging_rows = ws_aging.get_all_values()
    df_aging = pd.DataFrame(aging_rows[1:], columns=aging_rows[0])
    log(f"\n--- Aging Sheet Sample ({len(df_aging)} rows) ---")
    log(df_aging.head(5).to_string())

    # Load 'Treo LC'
    ws_treo = sh_main.worksheet("Treo LC")
    treo_rows = ws_treo.get_all_values()
    df_treo = pd.DataFrame(treo_rows[1:], columns=treo_rows[0])
    log(f"\n--- Treo LC Sheet Sample ({len(df_treo)} rows) ---")
    log(df_treo.head(5).to_string())

    # Load 'OPR'
    ws_opr = sh_main.worksheet("OPR")
    opr_rows = ws_opr.get_all_values()
    df_opr = pd.DataFrame(opr_rows[1:], columns=opr_rows[0])
    df_opr_yesterday = df_opr[df_opr['NgayLTC'] == '2026-06-23']
    log(f"\n--- OPR Sheet on 2026-06-23 ({len(df_opr_yesterday)} rows) ---")
    log(df_opr_yesterday.head(10).to_string())

    # Load Rot LC 'Data TTS'
    ws_rot_tts = sh_rot.worksheet("Data TTS")
    rot_tts_rows = ws_rot_tts.get_all_values()
    df_rot_tts = pd.DataFrame(rot_tts_rows[1:], columns=rot_tts_rows[0])
    df_rot_tts_yesterday = df_rot_tts[df_rot_tts['Loại ngày'] == '2026-06-23']
    log(f"\n--- Rot LC Data TTS on 2026-06-23 ({len(df_rot_tts_yesterday)} rows) ---")
    log(df_rot_tts_yesterday.head(10).to_string())

if __name__ == "__main__":
    main()
