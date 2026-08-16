import sys
import pandas as pd
from google.oauth2.credentials import Credentials as UserCredentials
import gspread

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = UserCredentials.from_authorized_user_file('authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)
sh = gc.open_by_key('1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk')

ws_gtc = sh.worksheet('gtc')
data_gtc = ws_gtc.get_all_values()
df_gtc = pd.DataFrame(data_gtc[1:], columns=data_gtc[0])
print('=== ALL UNIQUE DATES IN gtc TAB ===')
gtc_dates = sorted(df_gtc['Time'].unique())
for d in gtc_dates:
    print(d)

ws_data = sh.worksheet('Data')
data = ws_data.get_all_values()
df_data = pd.DataFrame(data[1:], columns=data[0])
print('\n=== ALL UNIQUE DATES IN Data TAB ===')
data_dates = sorted(df_data['Time'].unique())
for d in data_dates:
    print(d)
