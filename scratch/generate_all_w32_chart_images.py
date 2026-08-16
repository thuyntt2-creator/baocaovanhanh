import matplotlib.pyplot as plt
import numpy as np
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['axes.edgecolor'] = '#E0E0E0'

# 1. Chart 1: So sánh W31 vs W32 (image3.png)
categories = ['%GTC Full hàng', '%GTC TTS', '%ODR Full hàng', '%LTC Full hàng']
w31_vals = [59.2, 60.4, 93.7, 90.1]
w32_vals = [59.3, 59.7, 93.5, 89.3]

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

for spine in ['top', 'right', 'left']:
    ax.spines[spine].set_visible(False)

for rect in rects1:
    h = rect.get_height()
    ax.annotate(f'{h:.1f}%', xy=(rect.get_x() + rect.get_width() / 2, h), xytext=(0, 3),
                textcoords="offset points", ha='center', va='bottom', fontsize=9, fontweight='bold')

for rect in rects2:
    h = rect.get_height()
    ax.annotate(f'{h:.1f}%', xy=(rect.get_x() + rect.get_width() / 2, h), xytext=(0, 3),
                textcoords="offset points", ha='center', va='bottom', fontsize=9, fontweight='bold')

ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=2, frameon=False, fontsize=10)
plt.tight_layout()
path3 = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image3_w31_vs_w32.png'
plt.savefig(path3, dpi=200, bbox_inches='tight')
plt.close()

# 2. Chart 2: Xu hướng Sản lượng 4 tuần W29 - W32 (image21.png)
weeks = ['W29', 'W30', 'W31', 'W32']
full_sl = [361.763, 323.175, 326.075, 344.835] # in thousands
tts_sl = [96.110, 89.158, 87.844, 92.317]

fig, ax1 = plt.subplots(figsize=(8, 4.2), dpi=200)
fig.patch.set_facecolor('#FFFFFF')
ax1.set_facecolor('#FFFFFF')

x2 = np.arange(len(weeks))
l1 = ax1.plot(x2, full_sl, marker='o', color='#0072BC', linewidth=2.5, markersize=7, label='Full hàng (đơn)')
l2 = ax1.plot(x2, tts_sl, marker='s', color='#FA6400', linewidth=2.5, markersize=7, label='TTS (đơn)')

ax1.set_title('Xu hướng Sản lượng 4 tuần (W29 - W32)', fontsize=13, fontweight='bold', pad=15, color='#1B365D')
ax1.set_xticks(x2)
ax1.set_xticklabels(weeks, fontweight='bold', fontsize=10, color='#333333')
ax1.yaxis.grid(True, linestyle='--', alpha=0.5, color='#E0E0E0')
ax1.set_axisbelow(True)

for spine in ['top', 'right', 'left']:
    ax1.spines[spine].set_visible(False)

for idx, val in enumerate(full_sl):
    ax1.annotate(f'{val*1000:,.0f}', xy=(idx, val), xytext=(0, 7), textcoords="offset points",
                 ha='center', va='bottom', fontsize=8.5, fontweight='bold', color='#0072BC')

for idx, val in enumerate(tts_sl):
    ax1.annotate(f'{val*1000:,.0f}', xy=(idx, val), xytext=(0, -14), textcoords="offset points",
                 ha='center', va='top', fontsize=8.5, fontweight='bold', color='#FA6400')

ax1.legend(loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=2, frameon=False, fontsize=9.5)
plt.tight_layout()
path21 = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image21_trend_w29_w32.png'
plt.savefig(path21, dpi=200, bbox_inches='tight')
plt.close()

print("Generated new chart images for image3.png and image21.png!")
