# -*- coding: utf-8 -*-
import sys, os, pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

ntb_file = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
charts_dir = r'c:\Users\lap4all\Documents\Auto report\scratch\ntb_88_charts'
os.makedirs(charts_dir, exist_ok=True)

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

dates = [c.split()[-1] for c in days10_lay]

def render_tnb_style_combo_chart(x_labels, y_vals, title_text, legend_bar_label, filename, y_max_limit=None):
    fig, ax1 = plt.subplots(figsize=(10, 5.2), dpi=200)
    fig.patch.set_facecolor('#FFFFFF')
    ax1.set_facecolor('#FFFFFF')

    avg_val = np.mean(y_vals)
    pct_diffs = [(v - avg_val) / avg_val * 100 for v in y_vals]

    x_indices = np.arange(len(x_labels))
    bar_width = 0.52

    # Primary Axis: Bars (Dark Blue)
    bars = ax1.bar(x_indices, y_vals, width=bar_width, color='#0A2540', label=legend_bar_label, zorder=2)

    # Primary Axis: Red Line & Red Dots (Trung bình)
    ax1.axhline(y=avg_val, color='#E53935', linestyle='-', linewidth=2, zorder=3)
    ax1.plot(x_indices, [avg_val]*len(x_labels), marker='o', color='#E53935', markersize=7, linestyle='', label='Trung bình', zorder=4)

    # Secondary Axis: Yellow Line (% Tăng/ giảm FC)
    ax2 = ax1.twinx()
    ax2.plot(x_indices, pct_diffs, color='#FFB300', marker='o', linewidth=2.2, markersize=6, label='Tăng/ giảm FC', zorder=5)

    # Axis limits & formatting
    max_y = max(y_vals)
    top_limit = y_max_limit if y_max_limit else max_y * 1.35
    ax1.set_ylim(0, top_limit)
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{int(x):,}'))
    
    # Secondary Y limit
    max_pct = max(abs(min(pct_diffs)), abs(max(pct_diffs)))
    sec_limit = max(30, np.ceil(max_pct / 10) * 10 + 10)
    ax2.set_ylim(-sec_limit, sec_limit)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:.2f}%'))

    # Grid & Spines
    ax1.grid(False)
    ax2.grid(False)
    for spine in ['top', 'left', 'right', 'bottom']:
        ax1.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)

    # Title
    plt.title(title_text, fontsize=15, fontweight='bold', color='#000000', pad=25)

    # Custom Legend on Top Center
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', bbox_to_anchor=(0.5, 1.08), ncol=3, frameon=False, fontsize=9.5)

    # Annotate Bar Labels (Numbers on bars)
    for idx, rect in enumerate(bars):
        height = rect.get_height()
        ax1.annotate(f'{int(round(height)):,}',
                     xy=(rect.get_x() + rect.get_width() / 2, height / 2),
                     xytext=(0, 0), textcoords="offset points",
                     ha='center', va='center', color='#FFFFFF', fontweight='bold', fontsize=8.5)

    # Annotate Red Line Average Labels (Red numbers above line)
    for idx in range(len(x_labels)):
        ax1.annotate(f'{int(round(avg_val)):,}',
                     xy=(idx, avg_val),
                     xytext=(0, 10), textcoords="offset points",
                     ha='center', va='bottom', color='#E53935', fontweight='bold', fontsize=8)

    # Annotate Yellow Line % Labels
    for idx, pct in enumerate(pct_diffs):
        offset_y = -18 if pct < 0 else 14
        pct_str = f'{pct:+.2f}%' if pct != 0 else '0.00%'
        ax2.annotate(pct_str,
                     xy=(idx, pct),
                     xytext=(0, offset_y), textcoords="offset points",
                     ha='center', va='center', color='#FFB300', fontweight='bold', fontsize=8,
                     bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', edgecolor='#FFB300', alpha=0.9, linewidth=0.8))

    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(x_labels, fontweight='bold', fontsize=9.5)

    plt.tight_layout()
    file_path = os.path.join(charts_dir, filename)
    plt.savefig(file_path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f'Combo chart saved: {file_path}')
    return file_path

# Generate Lay & Giao combo charts
lay_daily = df_lay[days10_lay].sum().values
giao_daily = df_giao[days10_giao].sum().values

render_tnb_style_combo_chart(dates, lay_daily, 'FC Volume Lấy event 08.08', 'Volume lấy', 'chart_tnb_combo_lay_88.png')
render_tnb_style_combo_chart(dates, giao_daily, 'FC Volume Giao event 08.08', 'Volume giao', 'chart_tnb_combo_giao_88.png')

