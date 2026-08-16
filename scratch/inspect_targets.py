import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

targets = ['20942000', '22830000', 'Di Linh', 'Cam Linh']

# Inspect Gửi sếp sheet
print("=== SHEET: Gửi sếp ===")
ws_gs = spreadsheet.worksheet("Gửi sếp")
gs_vals = ws_gs.get_all_values()
df_gs = pd.DataFrame(gs_vals)
print("Gửi sếp shape:", df_gs.shape)
print(df_gs.head(30).to_string())

# Inspect CơCấuVùng sheet
print("\n=== SHEET: CơCấuVùng ===")
ws_cc = spreadsheet.worksheet("CơCấuVùng")
cc_vals = ws_cc.get_all_values()
df_cc = pd.DataFrame(cc_vals)
print("CơCấuVùng shape:", df_cc.shape)
for idx, row in df_cc.iterrows():
    row_str = " | ".join([str(c) for c in row if c])
    if any(t in row_str for t in targets):
        print(f"Row {idx}: {row_str}")

# Filter Thu Nhập sheet
print("\n=== SHEET: thu nhập (Target Rows) ===")
ws_tn = spreadsheet.worksheet("thu nhập")
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
print("Columns in thu nhập:", list(df_tn.columns))

for target in ['20942000', '22830000']:
    sub_tn = df_tn[df_tn['Bưu cục'].str.contains(target, na=False)]
    print(f"\n--- Target {target} in thu nhập ({len(sub_tn)} rows) ---")
    print(sub_tn.to_string())

# Filter Data sheet
print("\n=== SHEET: data (Target Rows) ===")
ws_data = spreadsheet.worksheet("data")
data_vals = ws_data.get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])
print("Columns in data:", list(df_data.columns))

for target in ['20942000', '22830000', 'Di Linh', 'Cam Linh']:
    sub_data = df_data[df_data.apply(lambda r: r.astype(str).str.contains(target).any(), axis=1)]
    print(f"\n--- Target {target} in data ({len(sub_data)} rows) ---")
    print(sub_data.head(10).to_string())

