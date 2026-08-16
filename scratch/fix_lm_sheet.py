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

print("1. Đang tải dữ liệu tab LM và cocaumoi...")
lm_vals = ws_lm.get_all_values()
co_vals = ws_co.get_all_values()

df_lm = pd.DataFrame(lm_vals[1:], columns=lm_vals[0])
df_co = pd.DataFrame(co_vals[1:], columns=co_vals[0])

# Clean warehouse IDs for 100% exact string matching
df_lm['bc_clean'] = df_lm['Mã bưu cục'].astype(str).str.strip()
df_co['wh_clean'] = df_co['warehouse_id'].astype(str).str.strip()

merged = pd.merge(df_lm, df_co[['wh_clean', 'AM']], left_on='bc_clean', right_on='wh_clean', how='left')
am_list = merged['AM_y'].fillna("Không xác định").values.tolist()

print(f"2. Đang điền dữ liệu tĩnh AM cho {len(am_list):,} dòng vào cột H tab 'LM'...")
am_column_grid = [["AM"]] + [[val] for val in am_list]

# Write all AM values to H1:H{len(am_column_grid)} in 1 single RAW API call
ws_lm.update(range_name=f"H1:H{len(am_column_grid)}", values=am_column_grid, value_input_option="RAW")

print(f"3. Đang thu gọn tab 'LM' từ 81,000 dòng xuống {len(am_column_grid) + 10} dòng...")
try:
    ws_lm.resize(rows=len(am_column_grid) + 10, cols=8)
    print("✔️ Đã thu gọn kích thước tab 'LM' thành công!")
except Exception as e:
    print(f"⚠️ Không thể thu gọn kích thước: {e}")

print("🎉 HOÀN THÀNH SỬA TAB LM!")
