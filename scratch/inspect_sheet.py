import sys
import json
import pandas as pd
from google.oauth2.credentials import Credentials as UserCredentials
import gspread

sys.stdout.reconfigure(encoding='utf-8')

try:
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = UserCredentials.from_authorized_user_file('authorized_user.json', scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key('1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk')

    ws_gtc = sh.worksheet('gtc')
    data_gtc = ws_gtc.get_all_values()
    print('gtc rows:', len(data_gtc))
    if data_gtc:
        print('gtc Header:', data_gtc[0])
        df_gtc = pd.DataFrame(data_gtc[1:], columns=data_gtc[0])
        print('gtc columns:', df_gtc.columns.tolist())
        for col in df_gtc.columns:
            if any(k in col.lower() for k in ['time', 'ngày', 'ca']):
                print(f'gtc unique values in {col}:', df_gtc[col].unique()[:20])

    ws_data = sh.worksheet('Data')
    data = ws_data.get_all_values()
    print('Data rows:', len(data))
    if data:
        print('Data Header:', data[0])
        df_data = pd.DataFrame(data[1:], columns=data[0])
        print('Data columns:', df_data.columns.tolist())
        for col in df_data.columns:
            if any(k in col.lower() for k in ['time', 'ngày', 'ca']):
                print(f'Data unique values in {col}:', df_data[col].unique()[:30])
except Exception as e:
    print("Error:", e)
