# -*- coding: utf-8 -*-
"""修复图3-4：降低标注密度，提升信息区分度"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.font_manager as fm
from matplotlib.patches import Patch
from matplotlib.lines import Line2D

FONT_PATH = '/System/Library/Fonts/Hiragino Sans GB.ttc'
fm.fontManager.addfont(FONT_PATH)
rcParams['font.family'] = fm.FontProperties(fname=FONT_PATH).get_name()
rcParams['axes.unicode_minus'] = False
rcParams['font.size'] = 11
rcParams['figure.dpi'] = 300
rcParams['savefig.dpi'] = 300
rcParams['savefig.bbox'] = 'tight'

C_MAIN = '#2E5F8A'
C_ACCENT = '#C0392B'
C_ORANGE = '#E67E22'
C_GREEN = '#27AE60'

OUT = "/Users/liuxinyi18/Desktop/FreshRetailNet-50K/论文图表"

df = pd.read_parquet("/Users/liuxinyi18/Desktop/FreshRetailNet-50K/data/train.parquet")

# 复用图2选中的门店0-商品4序列
g = df[(df['store_id'] == 0) & (df['product_id'] == 4)].sort_values('dt').reset_index(drop=True)

fig, ax = plt.subplots(figsize=(11, 4.8))
xs = pd.to_datetime(g['dt'])

# 主线：观测销量
ax.plot(xs, g['sale_amount'], color=C_MAIN, lw=1.6, marker='o', ms=3.5, zorder=3, alpha=0.9)

# 严重删失日（缺货≥4小时）：红色倒三角 —— 这些天的观测值被系统性压低
severe = g['stock_hour6_22_cnt'] >= 4
ax.scatter(xs[severe], g.loc[severe, 'sale_amount'], color=C_ACCENT, s=60,
           zorder=4, marker='v')

# 节假日背景色块
hol = g['holiday_flag'] == 1
for _, row in g[hol].iterrows():
    ax.axvspan(pd.to_datetime(row['dt']) - pd.Timedelta(hours=12),
               pd.to_datetime(row['dt']) + pd.Timedelta(hours=12),
               alpha=0.12, color=C_GREEN, zorder=1)

# 显著促销日（折扣力度位于该SKU前10%）：橙色圈
strong_promo = g['discount'] <= g['discount'].quantile(0.10)
ax.scatter(xs[strong_promo], g.loc[strong_promo, 'sale_amount'],
           facecolors='none', edgecolors=C_ORANGE, s=110, lw=1.8, zorder=5)

# 注释一个典型的"促销后断货"案例
cand = g[(g['stock_hour6_22_cnt'] >= 8)]
if len(cand) > 0:
    row = cand.iloc[cand['sale_amount'].argmax()]
    ax.annotate('促销高峰后连续缺货：\n观测销量≠真实需求', xy=(pd.to_datetime(row['dt']), row['sale_amount']),
                xytext=(pd.to_datetime(row['dt']) + pd.Timedelta(days=6), row['sale_amount'] + 2.2),
                fontsize=9.5, color=C_ACCENT,
                arrowprops=dict(arrowstyle='->', color=C_ACCENT, lw=1.2))

ax.set_xlabel('日期（2024年）')
ax.set_ylabel('日销量（归一化）')
ax.set_title('典型门店-SKU 90天销量时序（门店0 · 商品4）：缺货日的观测销量被系统性压低')

legend_elements = [
    Line2D([0], [0], color=C_MAIN, lw=1.8, marker='o', ms=5, label='日销量（观测值）'),
    Line2D([0], [0], color=C_ACCENT, marker='v', ms=8, lw=0, label='严重删失日（缺货≥4小时）'),
    Line2D([0], [0], marker='o', ms=10, lw=0, markerfacecolor='none',
           markeredgecolor=C_ORANGE, markeredgewidth=1.8, label='显著促销日'),
    Patch(facecolor=C_GREEN, alpha=0.15, label='节假日'),
]
ax.legend(handles=legend_elements, fontsize=9.5, loc='upper right', framealpha=0.95)
ax.grid(alpha=0.3, ls='--')
fig.autofmt_xdate()
plt.savefig(f'{OUT}/图3-4_90天时序删失日.png')
plt.close()
print("✅ 图3-4 已修复")
