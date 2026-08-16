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

ws_ns = spreadsheet.worksheet("năng suất")
vals_ns = ws_ns.get_all_values()
df_ns = pd.DataFrame(vals_ns[1:], columns=vals_ns[0])

ws_tn = spreadsheet.worksheet("thu nhập")
vals_tn = ws_tn.get_all_values()
df_tn = pd.DataFrame(vals_tn[1:], columns=vals_tn[0])

# Clean helper
def clean_num(val):
    if not val:
        return 0
    val_str = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0

df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)
df_ns['%GTC_num'] = df_ns['%GTC'].str.replace('%', '').str.replace(',', '.').apply(clean_num)

# Extract code from NhanVien string in năng suất (e.g. 3179025_Đoàn Minh Đại -> 3179025)
df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].strip())
df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].strip())

hubs = [
    ('(LDO) Di Linh', '20942000'),
    ('(KHO) Cam Linh', '22830000')
]

date = '22 thg 7, 2026'

for hub_name, hub_code in hubs:
    print(f"\n=======================================================")
    print(f"ANALYSIS FOR: {hub_name} ({hub_code}) - NGÀY {date}")
    print(f"=======================================================")
    
    sub_ns = df_ns[(df_ns['Bưu cục'] == hub_name) & (df_ns['Ngay'] == date)]
    sub_tn = df_tn[df_tn['Bưu cục'].str.contains(hub_code, na=False)]
    
    merged = pd.merge(sub_ns, sub_tn[['Emp_Code', 'Ngày vào làm', 'Thâm niên']], on='Emp_Code', how='left')
    
    print(f"Merged {len(merged)} employees for {date}:")
    print(merged[['NhanVien', 'Thâm niên', 'Ngày vào làm', 'TongDon_num', 'TongDonGTC_num', '%GTC_num']].to_string())
    
    print(f"\n--- GTC PERFORMANCE BY SENIORITY GROUP ({hub_name}) ---")
    for group, grp_df in merged.groupby('Thâm niên'):
        avg_gtc = grp_df['TongDonGTC_num'].mean()
        avg_rate = grp_df['%GTC_num'].mean()
        count = len(grp_df)
        low_count = len(grp_df[grp_df['TongDonGTC_num'] < 40])
        print(f"Group: {group:25s} | Số NV: {count:2d} | Avg GTC: {avg_gtc:5.1f} | Avg %GTC: {avg_rate:5.1f}% | Số NV GTC thấp (<40 đơn): {low_count}")
        for _, r in grp_df.iterrows():
            print(f"    - {r['NhanVien']:32s} (Vào làm: {r['Ngày vào làm']}) => GTC: {r['TongDonGTC_num']:2.0f} đơn | %GTC: {r['%GTC_num']:.1f}%")

