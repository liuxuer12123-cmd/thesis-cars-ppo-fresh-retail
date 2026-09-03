# -*- coding: utf-8 -*-
"""重画图3-3为分桶版：解决答辩时'16个小时怎么17根柱子'的困惑，并标注百分比口径"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.font_manager as fm

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
C_GREEN = '#27AE60'
C_ORANGE = '#E67E22'

OUT = "/Users/liuxinyi18/Desktop/FreshRetailNet-50K/论文图表"

df = pd.read_parquet("/Users/liuxinyi18/Desktop/FreshRetailNet-50K/data/train.parquet")
cnt = df['stock_hour6_22_cnt'].value_counts().sort_index()
pct = cnt / len(df) * 100

# 分桶：0(全天有货) / 1-3(轻微) / 4-7(中度) / 8-11(重度) / 12-15(极重) / 16(全天缺货)
buckets = [
    ('全天有货\n(0小时)', pct.loc[0], C_GREEN),
    ('轻微缺货\n(1-3小时)', pct.loc[1:3].sum(), '#8FBF8F'),
    ('中度缺货\n(4-7小时)', pct.loc[4:7].sum(), C_ORANGE),
    ('重度缺货\n(8-11小时)', pct.loc[8:11].sum(), '#D35400'),
    ('极重缺货\n(12-15小时)', pct.loc[12:15].sum(), C_ACCENT),
    ('全天缺货\n(16小时)', pct.loc[16], '#7B241C'),
]

fig, ax = plt.subplots(figsize=(8.5, 4.8))
labels = [b[0] for b in buckets]
values = [b[1] for b in buckets]
colors = [b[2] for b in buckets]
bars = ax.bar(range(len(buckets)), values, color=colors, alpha=0.9, width=0.62)
for i, v in enumerate(values):
    ax.text(i, v + 1.0, f'{v:.1f}%', ha='center', fontsize=11.5, fontweight='bold', color='#333')

# 累计缺货标注
total_oos = sum(values[1:])
ax.annotate(f'存在缺货合计 {total_oos:.1f}%', xy=(3, values[3]), xytext=(3.6, 42),
            fontsize=11, color=C_ACCENT, fontweight='bold',
            arrowprops=dict(arrowstyle='->', color=C_ACCENT, lw=1.3))

ax.set_xticks(range(len(buckets)))
ax.set_xticklabels(labels, fontsize=10.5)
ax.set_ylabel('样本占比（%）')
ax.set_xlabel('当日营业时段（6:00-22:00）缺货小时数区间')
ax.set_title('营业时段缺货程度分布（样本日 n = 4,500,000 = 898家门店 × 865个SKU × 90天）', fontsize=12.5)
ax.set_ylim(0, 63)
ax.grid(alpha=0.3, axis='y', ls='--')
# 口径注释
ax.text(0.99, 0.02, '口径：样本占比 = 落在该区间的（门店×商品×日）样本数 ÷ 全部样本日数',
        transform=ax.transAxes, ha='right', va='bottom', fontsize=8.5, color='#777', style='italic')
plt.savefig(f'{OUT}/图3-3_缺货小时数分布.png')
plt.close()
print("✅ 图3-3 分桶版已生成")
print("分桶明细:", {l.replace(chr(10), ''): f'{v:.1f}%' for l, v, _ in buckets})
