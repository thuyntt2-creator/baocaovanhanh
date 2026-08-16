import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb_path = r'C:\Users\lap4all\Downloads\NTB - BÁO CÁO KINH DOANH.xlsx'
xl = pd.ExcelFile(wb_path)

df_1dt = xl.parse('1_doanhthu')
df_1vol = xl.parse('1_volume')
df_2 = xl.parse('2')
df_3 = xl.parse('3')
df_khm = xl.parse('KHM ')

def parse_val(v):
    if pd.isna(v): return 0.0
    s = str(v).replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

df_1dt['DT_num'] = df_1dt['DT'].apply(parse_val)
df_1vol['Vol_num'] = df_1vol['DT'].apply(parse_val)

print("=== 1. W31 vs W30 SUMMARY ===")
print("W30 Revenue:", df_1dt[df_1dt['Tuan']=='2026/30']['DT_num'].sum())
print("W31 Revenue:", df_1dt[df_1dt['Tuan']=='2026/31']['DT_num'].sum())
print("W30 Volume:", df_1vol[df_1vol['Tuan']=='2026/30']['Vol_num'].sum())
print("W31 Volume:", df_1vol[df_1vol['Tuan']=='2026/31']['Vol_num'].sum())

print("\n=== 2. GROUP BREAKDOWN W31 vs W30 ===")
for nhom in ['A', 'BCD', 'EF', 'G']:
    dt_30 = df_1dt[(df_1dt['NhomKH']==nhom) & (df_1dt['Tuan']=='2026/30')]['DT_num'].values
    dt_31 = df_1dt[(df_1dt['NhomKH']==nhom) & (df_1dt['Tuan']=='2026/31')]['DT_num'].values
    vol_30 = df_1vol[(df_1vol['NhomKH']==nhom) & (df_1vol['Tuan']=='2026/30')]['Vol_num'].values
    vol_31 = df_1vol[(df_1vol['NhomKH']==nhom) & (df_1vol['Tuan']=='2026/31')]['Vol_num'].values
    dt30_v = dt_30[0] if len(dt_30)>0 else 0
    dt31_v = dt_31[0] if len(dt_31)>0 else 0
    vol30_v = vol_30[0] if len(vol_30)>0 else 0
    vol31_v = vol_31[0] if len(vol_31)>0 else 0
    pct_dt = (dt31_v - dt30_v) / dt30_v * 100 if dt30_v>0 else 0
    pct_vol = (vol31_v - vol30_v) / vol30_v * 100 if vol30_v>0 else 0
    print(f"Group {nhom:3s} | DT W30: {dt30_v:14,.0f} | DT W31: {dt31_v:14,.0f} ({pct_dt:+6.2f}%) | Vol W30: {vol30_v:7,.0f} | Vol W31: {vol31_v:7,.0f} ({pct_vol:+6.2f}%)")

print("\n=== 3. GROUP A CUSTOMERS IN W31 (Sheet 3) ===")
df_3_a = df_3[(df_3['Nhom'] == 'A') & (df_3['TuanCN'] == '2026/31')]
print(f"Total Group A customers in W31: {len(df_3_a)}")
for idx, r in df_3_a.iterrows():
    print(f"MaKH: {r['MaKH']} | {r['TenKH']} | DT W31: {r['DT']} | CamKet: {r['Cam ket thang']} | % WTD-1: {r['% sv WTD-1']} | AM: {r['AM']} | Tỉnh: {r['Tỉnh']}")

print("\n=== 4. SHEET 2 CUSTOMERS FOR W31 (% sv WTD-1 < 0.70) ===")
s2_w31 = df_2[df_2['Tuan_1'] == '2026/31']
print(f"Total customers in Sheet 2 for W31: {len(s2_w31)}")
print(s2_w31['Nhom'].value_counts())

print("\n=== 5. KHM SUMMARY ===")
print(f"Total KHM rows: {len(df_khm)}")
print(f"KHM Revenue sum (NoVAT): {df_khm['DoanhThu_NoVAT'].sum():,.2f}")
print(f"KHM Volume sum: {df_khm['Volume'].sum():,}")
print("\nTop AMs in KHM:")
print(df_khm.groupby('AM').agg(count=('Mã KH', 'count'), dt=('DoanhThu_NoVAT', 'sum'), vol=('Volume', 'sum')).sort_values(by='dt', ascending=False).head(8))
print("\nTop Provinces in KHM:")
print(df_khm.groupby('Tinh').agg(count=('Mã KH', 'count'), dt=('DoanhThu_NoVAT', 'sum'), vol=('Volume', 'sum')).sort_values(by='dt', ascending=False))
