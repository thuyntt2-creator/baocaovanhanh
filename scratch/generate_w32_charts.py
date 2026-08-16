import matplotlib.pyplot as plt
import numpy as np
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Values for W31 vs W32 from 00_Tong quan
categories = ['%GTC Full hàng', '%GTC TTS', '%ODR Full hàng', '%LTC Full hàng']
w31_vals = [59.2, 60.4, 93.7, 90.1]
w32_vals = [59.3, 59.7, 93.5, 89.3]

plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['axes.edgecolor'] = '#E0E0E0'

fig, ax = plt.subplots(figsize=(8, 4.5), dpi=200)
fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FFFFFF')

x = np.arange(len(categories))
width = 0.32

rects1 = ax.bar(x - width/2, w31_vals, width, label='W31', color='#4A90E2')
rects2 = ax.bar(x + width/2, w32_vals, width, label='W32', color='#D0021B')

ax.set_title('So sánh W31 vs W32', fontsize=14, fontweight='bold', pad=15, color='#000000')
ax.set_xticks(x)
ax.set_xticklabels(categories, fontweight='bold', fontsize=10, color='#333333')
ax.set_ylim(0, 115)
ax.yaxis.grid(True, linestyle='--', alpha=0.5, color='#CCCCCC')
ax.set_axisbelow(True)

# Remove top/right/left spines for clean modern aesthetic
for spine in ['top', 'right', 'left']:
    ax.spines[spine].set_visible(False)

# Add data labels
def autolabel(rects):
    for rect in rects:
        height = rect.get_height()
        ax.annotate(f'{height:.1f}%',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight='bold')

autolabel(rects1)
autolabel(rects2)

ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=2, frameon=False, fontsize=10)

plt.tight_layout()
out_png = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image3_w31_vs_w32.png'
plt.savefig(out_png, dpi=200, bbox_inches='tight')
plt.close()

print(f"Generated new chart at: {out_png}")
