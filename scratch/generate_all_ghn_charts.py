# -*- coding: utf-8 -*-
import sys, os, pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

ntb_file = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
charts_dir = r'c:\Users\lap4all\Documents\Auto report\scratch\ghn_charts'
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

plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['axes.edgecolor'] = '#E0E0E0'

# Combo Charts
def render_ghn_combo_chart(x_labels, y_vals, title_text, legend_bar_label, filename):
    fig, ax1 = plt.subplots(figsize=(10, 5.0), dpi=200)
    fig.patch.set_facecolor('#FFFFFF')
    ax1.set_facecolor('#FFFFFF')

    avg_val = np.mean(y_vals)
    pct_diffs = [(v - avg_val) / avg_val * 100 for v in y_vals]
    x_indices = np.arange(len(x_labels))

    bar_colors = ['#FA6400' if i == 2 else '#0072BC' for i in range(len(x_labels))]

    bars = ax1.bar(x_indices, y_vals, width=0.52, color=bar_colors, label=legend_bar_label, zorder=2)
    ax1.axhline(y=avg_val, color='#C0392B', linestyle='--', linewidth=2, zorder=3, label='Trung bình đợt')
    ax1.plot(x_indices, [avg_val]*len(x_labels), marker='o', color='#C0392B', markersize=6, linestyle='', zorder=4)

    ax2 = ax1.twinx()
    ax2.plot(x_indices, pct_diffs, color='#009688', marker='D', linewidth=2.2, markersize=5.5, label='% Tăng/Giảm FC', zorder=5)

    ax1.set_ylim(0, max(y_vals) * 1.35)
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{int(x):,}'))

    max_pct = max(abs(min(pct_diffs)), abs(max(pct_diffs)))
    sec_limit = max(30, np.ceil(max_pct / 10) * 10 + 10)
    ax2.set_ylim(-sec_limit, sec_limit)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:.1f}%'))

    for spine in ['top', 'left', 'right', 'bottom']:
        ax1.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)

    plt.title(title_text, fontsize=14, fontweight='bold', color='#1B365D', pad=22)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', bbox_to_anchor=(0.5, 1.08), ncol=3, frameon=False, fontsize=9)

    for idx, rect in enumerate(bars):
        height = rect.get_height()
        ax1.annotate(f'{int(round(height)):,}',
                     xy=(rect.get_x() + rect.get_width() / 2, height / 2),
                     ha='center', va='center', color='#FFFFFF', fontweight='bold', fontsize=8.5)

    for idx in range(len(x_labels)):
        ax1.annotate(f'{int(round(avg_val)):,}',
                     xy=(idx, avg_val), xytext=(0, 8), textcoords="offset points",
                     ha='center', va='bottom', color='#C0392B', fontweight='bold', fontsize=8)

    for idx, pct in enumerate(pct_diffs):
        offset_y = -18 if pct < 0 else 14
        pct_str = f'{pct:+.1f}%' if pct != 0 else '0.0%'
        ax2.annotate(pct_str,
                     xy=(idx, pct), xytext=(0, offset_y), textcoords="offset points",
                     ha='center', va='center', color='#009688', fontweight='bold', fontsize=8,
                     bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', edgecolor='#009688', alpha=0.9, linewidth=0.8))

    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(x_labels, fontweight='bold', fontsize=9.5, color='#333333')

    plt.tight_layout()
    path = os.path.join(charts_dir, filename)
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path

# Line Breakdown Charts
def render_ghn_line_chart(df, groupby_col, cols_10, title_text, filename):
    fig, ax = plt.subplots(figsize=(9, 4.5), dpi=180)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    piv = df.groupby(groupby_col)[cols_10].sum()
    palette = ['#0072BC', '#FA6400', '#009688', '#8E44AD', '#3498DB', '#E74C3C', '#2ECC71']
    for idx, (name, row) in enumerate(piv.iterrows()):
        ax.plot(dates_header_10, row.values, marker='o', linewidth=2, label=name, color=palette[idx % len(palette)])

    ax.set_title(title_text, fontsize=13, fontweight='bold', pad=12, color='#1B365D')
    ax.set_ylabel('Sản lượng (đơn)', fontsize=9.5)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{int(x):,}'))
    ax.grid(True, linestyle='--', alpha=0.4)
    ax.legend(loc='upper right', frameon=True, fontsize=8.5)
    for spine in ['top', 'left', 'right', 'bottom']:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    path = os.path.join(charts_dir, filename)
    plt.savefig(path, dpi=180)
    plt.close()
    return path

lay_daily = df_lay[days10_lay].sum().values
giao_daily = df_giao[days10_giao].sum().values

render_ghn_combo_chart(dates_header_10, lay_daily, 'FC Volume Lấy Event 08.08 - Vùng Nam Trung Bộ', 'Volume Lấy', 'ghn_combo_lay.png')
render_ghn_combo_chart(dates_header_10, giao_daily, 'FC Volume Giao Event 08.08 - Vùng Nam Trung Bộ', 'Volume Giao', 'ghn_combo_giao.png')

render_ghn_line_chart(df_lay, 'Sàn', days10_lay, 'Chi tiết Sản lượng Lấy theo Sàn (Event 8.8 NTB)', 'ghn_lay_san.png')
render_ghn_line_chart(df_lay, 'Tỉnh/Quận', days10_lay, 'Sản lượng Lấy theo Tỉnh/Quận & Tỷ trọng (Event 8.8 NTB)', 'ghn_lay_tinh.png')
render_ghn_line_chart(df_giao, 'Sàn', days10_giao, 'Chi tiết Sản lượng Giao theo Sàn (Event 8.8 NTB)', 'ghn_giao_san.png')
render_ghn_line_chart(df_giao, 'Tỉnh/Quận', days10_giao, 'Sản lượng Giao theo Tỉnh/Quận & Tỷ trọng (Event 8.8 NTB)', 'ghn_giao_tinh.png')

print('All GHN Corporate charts generated successfully!')
