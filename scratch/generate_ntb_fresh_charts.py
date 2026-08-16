# -*- coding: utf-8 -*-
import sys, os, pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

ntb_file = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
charts_dir = r'c:\Users\lap4all\Documents\Auto report\scratch\ntb_fresh_charts'
os.makedirs(charts_dir, exist_ok=True)

# 1. Load Data
df_lay = pd.read_excel(ntb_file, sheet_name='6_FC_Lay_Daily')
df_giao = pd.read_excel(ntb_file, sheet_name='7_FC_Giao_Daily')
df_sort = pd.read_excel(ntb_file, sheet_name='FC Sorting 60d')

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

# FRESH MODERN PALETTE (Ocean & Teal Theme for Nam Trung Bộ)
PRIMARY_BAR_COLOR = '#0F4C81'  # Deep Ocean Blue
PEAK_BAR_COLOR = '#E67E22'     # Warm Amber Orange (for Peak Day 08/08)
AVG_LINE_COLOR = '#C0392B'     # Crimson Red
GROWTH_LINE_COLOR = '#2ECC71'   # Emerald Green
PALETTE_MULTI = ['#0F4C81', '#E67E22', '#2ECC71', '#8E44AD', '#3498DB', '#16A085', '#D35400']

plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['axes.edgecolor'] = '#E0E0E0'
plt.rcParams['axes.linewidth'] = 0.8

def render_fresh_combo_chart(x_labels, y_vals, title_text, legend_bar_label, filename):
    fig, ax1 = plt.subplots(figsize=(10, 5.2), dpi=200)
    fig.patch.set_facecolor('#FCFDFD')
    ax1.set_facecolor('#FCFDFD')

    avg_val = np.mean(y_vals)
    pct_diffs = [(v - avg_val) / avg_val * 100 for v in y_vals]
    x_indices = np.arange(len(x_labels))
    bar_width = 0.52

    # Color array: Highlight peak day (08/08 index 2) with PEAK_BAR_COLOR
    bar_colors = [PEAK_BAR_COLOR if i == 2 else PRIMARY_BAR_COLOR for i in range(len(x_labels))]

    # Primary Axis: Bars with dual colors
    bars = ax1.bar(x_indices, y_vals, width=bar_width, color=bar_colors, label=legend_bar_label, zorder=2, edgecolor='none')

    # Primary Axis: Dashed Red Line (Trung bình)
    ax1.axhline(y=avg_val, color=AVG_LINE_COLOR, linestyle='--', linewidth=2, zorder=3, label='Trung bình đợt')
    ax1.plot(x_indices, [avg_val]*len(x_labels), marker='o', color=AVG_LINE_COLOR, markersize=6, linestyle='', zorder=4)

    # Secondary Axis: Emerald Green Line (% Tăng/giảm)
    ax2 = ax1.twinx()
    ax2.plot(x_indices, pct_diffs, color=GROWTH_LINE_COLOR, marker='D', linewidth=2.2, markersize=5.5, label='% Tăng/Giảm FC', zorder=5)

    # Axis limits
    max_y = max(y_vals)
    ax1.set_ylim(0, max_y * 1.35)
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{int(x):,}'))
    
    max_pct = max(abs(min(pct_diffs)), abs(max(pct_diffs)))
    sec_limit = max(30, np.ceil(max_pct / 10) * 10 + 10)
    ax2.set_ylim(-sec_limit, sec_limit)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:.1f}%'))

    # Hide Spines
    for spine in ['top', 'left', 'right', 'bottom']:
        ax1.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)

    # Title & Legend
    plt.title(title_text, fontsize=14, fontweight='bold', color='#1F497D', pad=25)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', bbox_to_anchor=(0.5, 1.08), ncol=3, frameon=False, fontsize=9)

    # Annotate Numbers on Bars
    for idx, rect in enumerate(bars):
        height = rect.get_height()
        col = '#FFFFFF'
        ax1.annotate(f'{int(round(height)):,}',
                     xy=(rect.get_x() + rect.get_width() / 2, height / 2),
                     xytext=(0, 0), textcoords="offset points",
                     ha='center', va='center', color=col, fontweight='bold', fontsize=8.5)

    # Annotate Average Value above line
    for idx in range(len(x_labels)):
        ax1.annotate(f'{int(round(avg_val)):,}',
                     xy=(idx, avg_val),
                     xytext=(0, 8), textcoords="offset points",
                     ha='center', va='bottom', color=AVG_LINE_COLOR, fontweight='bold', fontsize=8)

    # Annotate % Growth Badges
    for idx, pct in enumerate(pct_diffs):
        offset_y = -18 if pct < 0 else 14
        pct_str = f'{pct:+.1f}%' if pct != 0 else '0.0%'
        ax2.annotate(pct_str,
                     xy=(idx, pct),
                     xytext=(0, offset_y), textcoords="offset points",
                     ha='center', va='center', color=GROWTH_LINE_COLOR, fontweight='bold', fontsize=8,
                     bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', edgecolor=GROWTH_LINE_COLOR, alpha=0.9, linewidth=0.8))

    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(x_labels, fontweight='bold', fontsize=9.5, color='#333333')

    plt.tight_layout()
    file_path = os.path.join(charts_dir, filename)
    plt.savefig(file_path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f'Saved fresh combo chart: {file_path}')
    return file_path

# Generate Fresh Combo Charts
lay_daily = df_lay[days10_lay].sum().values
giao_daily = df_giao[days10_giao].sum().values

render_fresh_combo_chart(dates_header_10, lay_daily, 'FC Volume Lấy Event 08.08 - Vùng Nam Trung Bộ', 'Volume Lấy', 'fresh_combo_lay_88.png')
render_fresh_combo_chart(dates_header_10, giao_daily, 'FC Volume Giao Event 08.08 - Vùng Nam Trung Bộ', 'Volume Giao', 'fresh_combo_giao_88.png')

# Breakdown line charts
def render_fresh_line_chart(df, groupby_col, cols_10, title_text, filename):
    fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=180)
    fig.patch.set_facecolor('#FCFDFD')
    ax.set_facecolor('#FCFDFD')

    piv = df.groupby(groupby_col)[cols_10].sum()
    for idx, (name, row) in enumerate(piv.iterrows()):
        ax.plot(dates_header_10, row.values, marker='o', linewidth=2, label=name, color=PALETTE_MULTI[idx % len(PALETTE_MULTI)])

    ax.set_title(title_text, fontsize=12, fontweight='bold', pad=12, color='#1F497D')
    ax.set_ylabel('Sản lượng (đơn)', fontsize=9.5)
    ax.grid(True, linestyle='--', alpha=0.4)
    ax.legend(loc='upper right', frameon=True, fontsize=8.5)
    for spine in ['top', 'left', 'right', 'bottom']:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    path = os.path.join(charts_dir, filename)
    plt.savefig(path, dpi=180)
    plt.close()
    return path

render_fresh_line_chart(df_lay, 'Sàn', days10_lay, 'Sản lượng Lấy theo Sàn (Event 8.8 NTB)', 'fresh_lay_san.png')
render_fresh_line_chart(df_lay, 'Tỉnh/Quận', days10_lay, 'Sản lượng Lấy theo Tỉnh (Event 8.8 NTB)', 'fresh_lay_tinh.png')
render_fresh_line_chart(df_giao, 'Sàn', days10_giao, 'Sản lượng Giao theo Sàn (Event 8.8 NTB)', 'fresh_giao_san.png')
render_fresh_line_chart(df_giao, 'Tỉnh/Quận', days10_giao, 'Sản lượng Giao theo Tỉnh (Event 8.8 NTB)', 'fresh_giao_tinh.png')

# KTC Sorting Combo Chart
target_cols_sort = list(range(24, 34))
hubs_list = [
    'Kho Trung Chuyển Khánh Hòa',
    'Kho Chuyển Tiếp Bình Thuận',
    'Kho Chuyển Tiếp Đức Trọng-Lâm Đồng',
    'Kho Chuyển Tiếp Bảo Lộc-Lâm Đồng',
    'Kho Chuyển Tiếp Đắk Nông'
]

hub_data_dict = {}
for h in hubs_list:
    hub_rows = df_sort[df_sort.iloc[:, 0] == h]
    hub_data_dict[h] = {}
    for idx, r in hub_rows.iterrows():
        cat = str(r.iloc[1]).strip()
        vals = [int(round(r.iloc[c])) if pd.notna(r.iloc[c]) and isinstance(r.iloc[c], (int, float)) else 0 for c in target_cols_sort]
        hub_data_dict[h][cat] = np.array(vals)

total_ktc_daily = np.zeros(10)
for h in hubs_list:
    total_ktc_daily += hub_data_dict[h]['Total']

render_fresh_combo_chart(dates_header_10, total_ktc_daily, 'FC Volume Sorting 5 Kho KTC/Chuyển Tiếp NTB (Event 8.8)', 'Volume Sorting KTC', 'fresh_combo_ktc_sorting.png')

# 5 KTC Breakdown Line Chart
fig, ax = plt.subplots(figsize=(9, 4.5), dpi=180)
fig.patch.set_facecolor('#FCFDFD')
ax.set_facecolor('#FCFDFD')
for idx, h in enumerate(hubs_list):
    short_name = h.replace('Kho Trung Chuyển ', '').replace('Kho Chuyển Tiếp ', '')
    ax.plot(dates_header_10, hub_data_dict[h]['Total'], marker='s', linewidth=2, label=short_name, color=PALETTE_MULTI[idx])

ax.set_title('Phân bổ sản lượng Sorting 5 Kho KTC / Chuyển tiếp NTB (Event 8.8)', fontsize=12, fontweight='bold', pad=12, color='#1F497D')
ax.set_ylabel('Sản lượng (đơn)', fontsize=9.5)
ax.grid(True, linestyle='--', alpha=0.4)
ax.legend(loc='upper right', frameon=True, fontsize=8.5)
for spine in ['top', 'left', 'right', 'bottom']:
    ax.spines[spine].set_visible(False)
plt.tight_layout()
chart_ktc_hubs_path = os.path.join(charts_dir, 'fresh_ktc_hubs.png')
plt.savefig(chart_ktc_hubs_path, dpi=180)
plt.close()

print('All fresh modern charts rendered successfully!')
