import sys, time, gspread
import pandas as pd
import numpy as np
import unicodedata
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

t0 = time.time()
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
t_auth = time.time() - t0
print(f"✔️ Auth & open sheet: {t_auth:.2f}s")

t1 = time.time()
ws_lm = sh.worksheet('LM')
lm_data = ws_lm.get_all_values()
df_lm = pd.DataFrame(lm_data[1:], columns=lm_data[0])
t_read = time.time() - t1
print(f"✔️ Read LM tab directly ({len(df_lm):,} rows): {t_read:.2f}s")

print("LM columns:", df_lm.columns.tolist())
print("Sample row 1:", df_lm.iloc[0].to_dict())
