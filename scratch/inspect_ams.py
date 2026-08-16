import sys, gspread, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
gc = gspread.service_account(filename='credentials.json')
sh = gc.open_by_key('1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ')
ws = sh.worksheet('CoCauVung')
data = ws.get_all_values()
df = pd.DataFrame(data[1:], columns=data[0])

print("=== Unique AMs in CoCauVung ===")
for am in df['AM'].unique():
    print(f"- '{am}'")

print("\n=== Checking rows for Hiền, Lực, Hoàng ===")
target_ams = ['Hiền', 'Lực', 'Hoàng']
for am_sub in target_ams:
    sub_df = df[df['AM'].str.contains(am_sub, case=False, na=False)]
    print(f"\n--- Matching '{am_sub}' ({len(sub_df)} rows) ---")
    print(sub_df[['warehouse_id', 'Bưu cục', 'Tỉnh', 'AM']])
