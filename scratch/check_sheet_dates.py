import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(MAIN_SPREADSHEET_ID)
    
    print("--- Data Sheet ---")
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    times = df_data['Time'].unique()[:10]
    for t in times:
        print(f"Time: {t!r}")
    
    print("--- TTS Sheet ---")
    ws_tts = sh.worksheet("TTS")
    tts_rows = ws_tts.get_all_values()
    df_tts = pd.DataFrame(tts_rows[1:], columns=tts_rows[0])
    times_tts = df_tts['Time'].unique()[:10]
    for t in times_tts:
        print(f"Time TTS: {t!r}")
    
    print("--- OPR Sheet ---")
    ws_opr = sh.worksheet("OPR")
    opr_rows = ws_opr.get_all_values()
    df_opr = pd.DataFrame(opr_rows[1:], columns=opr_rows[0])
    dates = df_opr['NgayLTC'].unique()[:10]
    for d in dates:
        print(f"Date OPR: {d!r}")

if __name__ == "__main__":
    main()
