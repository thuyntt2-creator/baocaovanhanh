import sys, json, gspread
import pandas as pd
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
ws_lm = sh.worksheet('LM')
ws_co = sh.worksheet('cocaumoi')

lm_vals = ws_lm.get_all_values()
co_vals = ws_co.get_all_values()

df_lm = pd.DataFrame(lm_vals[1:], columns=lm_vals[0])
df_co = pd.DataFrame(co_vals[1:], columns=co_vals[0])

print(f"LM rows: {len(df_lm)}")
print(f"co_df rows: {len(df_co)}")

# Clean warehouse IDs
df_lm['Mã bưu cục clean'] = df_lm['Mã bưu cục'].astype(str).str.strip()
df_co['warehouse_id clean'] = df_co['warehouse_id'].astype(str).str.strip()

merged = pd.merge(df_lm, df_co, left_on='Mã bưu cục clean', right_on='warehouse_id clean', how='left')
matched_count = merged['AM_y'].notna().sum()
unmatched_count = merged['AM_y'].isna().sum()

print(f"Matched AM count in Python: {matched_count} / {len(df_lm)}")
print(f"Unmatched AM count: {unmatched_count}")

print("\nSample unmatched Mã bưu cục values:")
print(merged[merged['AM_y'].isna()]['Mã bưu cục clean'].unique()[:10])

print("\nSample matched AM values:")
print(merged[merged['AM_y'].notna()][['Mã bưu cục clean', 'Bưu cục_y', 'AM_y']].head(5))
