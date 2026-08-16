import sys, gspread, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
gc = gspread.service_account(filename='credentials.json')
sh = gc.open_by_key('1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ')
ws = sh.worksheet('CoCauVung')
data = ws.get_all_values()
df = pd.DataFrame(data[1:], columns=data[0])

print("=== All Rows in CoCauVung ===")
for am, group in df.groupby('AM'):
    bcs = group['Bưu cục'].tolist()
    provinces = group['Tỉnh'].unique().tolist()
    print(f"\nAM: '{am}' (Tỉnh: {provinces})")
    for bc in bcs:
        print(f"  - {bc}")
