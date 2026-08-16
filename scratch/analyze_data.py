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

print('=== 1. GROUP A CUSTOMERS ANALYSIS ===')
df_3_a = df_3[(df_3['Nhom'] == 'A') & (df_3['TuanCN'] == '2026/30')]
print(f"Total Group A customers in W30: {len(df_3_a)}")
for idx, row in df_3_a.iterrows():
    print(f"MaKH: {row['MaKH']} | TenKH: {row['TenKH']} | DT W30: {row['DT']} | CamKet: {row['Cam ket thang']} | % WTD-1: {row['% sv WTD-1']} | AM: {row['AM']} | Tỉnh: {row['Tỉnh']}")

print('\n=== Group A comparison W29 vs W30 ===')
df_3_a_all = df_3[df_3['Nhom'] == 'A']
for makh, group in df_3_a_all.groupby('MaKH'):
    name = group['TenKH'].iloc[0]
    am = group['AM'].iloc[0]
    tinh = group['Tỉnh'].iloc[0]
    w29_dt = group[group['TuanCN']=='2026/29']['DT'].values
    w30_dt = group[group['TuanCN']=='2026/30']['DT'].values
    w29_str = w29_dt[0] if len(w29_dt)>0 else "0"
    w30_str = w30_dt[0] if len(w30_dt)>0 else "0"
    pct = group['% sv WTD-1'].iloc[0]
    print(f"MaKH: {makh} | {name} | AM: {am} | Tỉnh: {tinh} | W29: {w29_str} | W30: {w30_str} | % sv WTD-1: {pct}")

print('\n=== 2. SHEET 2 CUSTOMERS (RISK / DECREASING CUSTOMERS) ===')
s2_w30 = df_2[df_2['Tuan_1'] == '2026/30']
print(f"Total customers with declining performance (% sv WTD-1 < 0.70) in Sheet 2 for W30: {len(s2_w30)}")
print(s2_w30['Nhom'].value_counts())

print('\n=== 3. SHEET 3 SUMMARY BY GROUP FOR W29 vs W30 ===')
for nhom, group in df_3.groupby('Nhom'):
    w29 = group[group['TuanCN']=='2026/29']
    w30 = group[group['TuanCN']=='2026/30']
    print(f"Group {nhom}: Unique KHs: {group['MaKH'].nunique()} | W29 rows: {len(w29)} | W30 rows: {len(w30)}")
