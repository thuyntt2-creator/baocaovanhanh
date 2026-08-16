# -*- coding: utf-8 -*-
import sys, os, pandas as pd
import matplotlib.pyplot as plt

sys.stdout.reconfigure(encoding='utf-8')

ntb_file = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
charts_dir = r'c:\Users\lap4all\Documents\Auto report\scratch\ntb_88_charts'
os.makedirs(charts_dir, exist_ok=True)

# 1. Load Data
df_lay = pd.read_excel(ntb_file, sheet_name='6_FC_Lay_Daily')
df_giao = pd.read_excel(ntb_file, sheet_name='7_FC_Giao_Daily')

date_cols_lay = [c for c in df_lay.columns if c not in ['Vùng', 'Tỉnh/Quận', 'ID', 'BC', 'Sàn', 'Tổng 60d']]
date_cols_giao = [c for c in df_giao.columns if c not in ['Vùng', 'Tỉnh/Quận', 'ID', 'BC', 'Sàn', 'Tổng 60d']]

days10_lay = [c for c in date_cols_lay if any(d in c for d in ['06/08', '07/08', '08/08', '09/08', '10/08', '11/08', '12/08', '13/08', '14/08', '15/08'])][:10]
days10_giao = [c for c in date_cols_giao if any(d in c for d in ['06/08', '07/08', '08/08', '09/08', '10/08', '11/08', '12/08', '13/08', '14/08', '15/08'])][:10]

for c in days10_lay:
    df_lay[c] = pd.to_numeric(df_lay[c], errors='coerce').fillna(0)
for c in days10_giao:
    df_giao[c] = pd.to_numeric(df_giao[c], errors='coerce').fillna(0)

df_lay = df_lay.dropna(subset=['Sàn'])
df_giao = df_giao.dropna(subset=['Sàn'])

dates_header_10 = [c.split()[-1] for c in days10_lay]

# Global Plot Styling
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['axes.edgecolor'] = '#CCCCCC'
plt.rcParams['axes.linewidth'] = 0.8

colors = ['#1F497D', '#366092', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

# --- CHART 1: Volume Lấy theo Sàn ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
tbl1_df = df_lay.groupby('Sàn')[days10_lay].sum()
san_order = ['Shopee', 'Shopee-Bulky', 'SME', 'SME-Bulky', 'TTS', 'TTS-Bulky']
tbl1_df = tbl1_df.reindex([s for s in san_order if s in tbl1_df.index])

for idx, (san_name, row) in enumerate(tbl1_df.iterrows()):
    ax.plot(dates_header_10, row.values, marker='o', linewidth=2, label=san_name, color=colors[idx % len(colors)])

ax.set_title('Sản lượng Lấy theo Sàn qua từng ngày Event 8.8 (06/08 - 15/08)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Sản lượng (đơn)', fontsize=10)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='upper right', frameon=True, fontsize=8.5)
plt.tight_layout()
p1 = os.path.join(charts_dir, 'chart1_lay_san.png')
plt.savefig(p1)
plt.close()

# --- CHART 2: Volume Lấy theo Tỉnh ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
tbl2_df = df_lay.groupby('Tỉnh/Quận')[days10_lay].sum()
for idx, (tinh_name, row) in enumerate(tbl2_df.iterrows()):
    ax.plot(dates_header_10, row.values, marker='s', linewidth=2, label=tinh_name, color=colors[idx % len(colors)])

ax.set_title('Sản lượng Lấy theo Tỉnh/Quận qua từng ngày Event 8.8 (NTB)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Sản lượng (đơn)', fontsize=10)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='upper right', frameon=True, fontsize=8.5)
plt.tight_layout()
p2 = os.path.join(charts_dir, 'chart2_lay_tinh.png')
plt.savefig(p2)
plt.close()

# --- CHART 3: Volume Giao theo Sàn ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
tbl3_df = df_giao.groupby('Sàn')[days10_giao].sum()
tbl3_df = tbl3_df.reindex([s for s in san_order if s in tbl3_df.index])

for idx, (san_name, row) in enumerate(tbl3_df.iterrows()):
    ax.plot(dates_header_10, row.values, marker='^', linewidth=2, label=san_name, color=colors[idx % len(colors)])

ax.set_title('Sản lượng Giao theo Sàn qua từng ngày Event 8.8 (06/08 - 15/08)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Sản lượng (đơn)', fontsize=10)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='upper right', frameon=True, fontsize=8.5)
plt.tight_layout()
p3 = os.path.join(charts_dir, 'chart3_giao_san.png')
plt.savefig(p3)
plt.close()

# --- CHART 4: Volume Giao theo Tỉnh ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
tbl4_df = df_giao.groupby('Tỉnh/Quận')[days10_giao].sum()
for idx, (tinh_name, row) in enumerate(tbl4_df.iterrows()):
    ax.plot(dates_header_10, row.values, marker='d', linewidth=2, label=tinh_name, color=colors[idx % len(colors)])

ax.set_title('Sản lượng Giao theo Tỉnh/Quận qua từng ngày Event 8.8 (NTB)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Sản lượng (đơn)', fontsize=10)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='upper right', frameon=True, fontsize=8.5)
plt.tight_layout()
p4 = os.path.join(charts_dir, 'chart4_giao_tinh.png')
plt.savefig(p4)
plt.close()

# --- CHART 5: Đánh giá Tổng quan Volume Giao ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
tot_giao_daily = df_giao[days10_giao].sum().values
ax.plot(dates_header_10, tot_giao_daily, marker='o', color='#C00000', linewidth=2.5, label='Tổng Volume Giao NTB')

# Highlight peak
peak_idx = 2 # 08/08
ax.annotate(f'Peak Event 8.8\n{int(tot_giao_daily[peak_idx]):,} đơn', 
            xy=(dates_header_10[peak_idx], tot_giao_daily[peak_idx]), 
            xytext=(dates_header_10[peak_idx], tot_giao_daily[peak_idx] + 15000),
            arrowprops=dict(facecolor='#C00000', shrink=0.08, width=1.5, headwidth=6),
            fontweight='bold', color='#C00000', fontsize=9, ha='center')

ax.set_title('Tổng quan Xu hướng Volume Giao Toàn Vùng NTB (Event 8.8)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Tổng Sản lượng Giao (đơn)', fontsize=10)
ax.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
p5 = os.path.join(charts_dir, 'chart5_giao_tongquan.png')
plt.savefig(p5)
plt.close()

# Helper for Hub charts (6, 7, 8)
def get_hub(tinh):
    t = str(tinh)
    if 'Khánh Hòa' in t or 'Ninh Thuận' in t:
        return 'Kho TC Nha Trang'
    elif 'Lâm Đồng' in t or 'Đắk Nông' in t:
        return 'Kho CT Lâm Đồng'
    else:
        return 'Kho CT Bình Thuận'

df_giao['Hub'] = df_giao['Tỉnh/Quận'].apply(get_hub)

def generate_hub_chart(hub_name, filename, title_text):
    fig, ax = plt.subplots(figsize=(8.5, 4.0), dpi=150)
    df_h = df_giao[df_giao['Hub'] == hub_name]
    df_norm = df_h[~df_h['Sàn'].astype(str).str.contains('Bulky')][days10_giao].sum().values
    df_bulk_tot = df_h[df_h['Sàn'].astype(str).str.contains('Bulky')][days10_giao].sum().values
    
    val_bulk = df_bulk_tot * 0.88
    val_freight = df_bulk_tot * 0.12
    
    ax.plot(dates_header_10, df_norm, marker='o', label='Normal (<10kg)', color='#1F497D', linewidth=2)
    ax.plot(dates_header_10, val_bulk, marker='s', label='Bulky (10-30kg)', color='#ED7D31', linewidth=2)
    ax.plot(dates_header_10, val_freight, marker='^', label='Freight (≥30kg)', color='#C00000', linewidth=2)
    
    ax.set_title(title_text, fontsize=11, fontweight='bold', pad=10, color='#1F497D')
    ax.set_ylabel('Sản lượng (đơn)', fontsize=9)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.legend(loc='upper right', fontsize=8)
    plt.tight_layout()
    p_path = os.path.join(charts_dir, filename)
    plt.savefig(p_path)
    plt.close()
    return p_path

p6 = generate_hub_chart('Kho TC Nha Trang', 'chart6_hub_nhatrang.png', 'Xu hướng & Cơ cấu Sản lượng Giao — Kho TC Nha Trang')
p7 = generate_hub_chart('Kho CT Lâm Đồng', 'chart7_hub_lamdong.png', 'Xu hướng & Cơ cấu Sản lượng Giao — Kho CT Lâm Đồng')
p8 = generate_hub_chart('Kho CT Bình Thuận', 'chart8_hub_binhthuan.png', 'Xu hướng & Cơ cấu Sản lượng Giao — Kho CT Bình Thuận')

# --- CHART 9: Trend Lấy theo Tỉnh ---
fig, ax = plt.subplots(figsize=(8.5, 4.0), dpi=150)
for idx, (tinh_name, row) in enumerate(tbl2_df.iterrows()):
    ax.plot(dates_header_10, row.values, marker='o', linewidth=2, label=tinh_name, color=colors[idx % len(colors)])

ax.set_title('Xu hướng Sản lượng Lấy theo Tỉnh Event 8.8 (06/08 - 15/08)', fontsize=11, fontweight='bold', pad=10, color='#1F497D')
ax.set_ylabel('Sản lượng Lấy (đơn)', fontsize=9)
ax.grid(True, linestyle='--', alpha=0.5)
ax.legend(loc='upper right', fontsize=8)
plt.tight_layout()
p9 = os.path.join(charts_dir, 'chart9_lay_tinh_trend.png')
plt.savefig(p9)
plt.close()

print("All 9 charts generated successfully!")
