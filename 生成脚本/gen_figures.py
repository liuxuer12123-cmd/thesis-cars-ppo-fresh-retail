# -*- coding: utf-8 -*-
"""FreshRetailNet-50K 论文图表生成脚本
生成论文第三章"数据分析"部分可用的高清图表（300 DPI）
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import os

# ===== 中文字体与学术风格设置 =====
import matplotlib.font_manager as fm
# 直接注册字体文件（CLT Python 找不到系统字体名，用路径最稳妥）
FONT_PATH = '/System/Library/Fonts/Hiragino Sans GB.ttc'
fm.fontManager.addfont(FONT_PATH)
_zh = fm.FontProperties(fname=FONT_PATH)
rcParams['font.family'] = _zh.get_name()
rcParams['axes.unicode_minus'] = False
rcParams['font.size'] = 11
rcParams['axes.titlesize'] = 13
rcParams['axes.labelsize'] = 12
rcParams['figure.dpi'] = 300
rcParams['savefig.dpi'] = 300
rcParams['savefig.bbox'] = 'tight'

OUT = "/Users/liuxinyi18/Desktop/FreshRetailNet-50K/论文图表"
os.makedirs(OUT, exist_ok=True)

# 配色（学术蓝灰系）
C_MAIN = '#2E5F8A'    # 主色 深蓝
C_ACCENT = '#C0392B'  # 强调 红（缺货）
C_GRAY = '#7F8C8D'
C_GREEN = '#27AE60'
C_ORANGE = '#E67E22'

print("加载数据...")
df = pd.read_parquet("/Users/liuxinyi18/Desktop/FreshRetailNet-50K/data/train.parquet")
print(f"  train: {len(df):,} 行")

# ============================================================
# 图1: 逐小时缺货率曲线（U型）
# ============================================================
print("图1: 逐小时缺货率曲线...")
stock_matrix = np.stack(df['hours_stock_status'].values)
hourly_oos = stock_matrix.mean(axis=0) * 100

fig, ax = plt.subplots(figsize=(8, 4.5))
hours = np.arange(24)
ax.plot(hours, hourly_oos, color=C_MAIN, lw=2.2, marker='o', ms=5, zorder=3)
ax.fill_between(hours, hourly_oos, alpha=0.15, color=C_MAIN)
# 标注营业时段
ax.axvspan(6, 22, alpha=0.06, color=C_GREEN)
ax.text(14, 44, '营业时段 6:00-22:00', ha='center', fontsize=10, color=C_GREEN)
# 标注早晚峰
ax.annotate(f'凌晨缺货率 {hourly_oos[0]:.1f}%\n（前日库存耗尽）', xy=(0, hourly_oos[0]),
            xytext=(1.2, 38), fontsize=9.5, color=C_ACCENT,
            arrowprops=dict(arrowstyle='->', color=C_ACCENT, lw=1.2))
ax.annotate(f'补货后骤降至 {hourly_oos[6]:.1f}%', xy=(6, hourly_oos[6]),
            xytext=(7.5, 26), fontsize=9.5, color=C_MAIN,
            arrowprops=dict(arrowstyle='->', color=C_MAIN, lw=1.2))
ax.annotate(f'晚间爬升至 {hourly_oos[21]:.1f}%\n（日间销售消耗库存）', xy=(21, hourly_oos[21]),
            xytext=(16.5, 35), fontsize=9.5, color=C_ACCENT,
            arrowprops=dict(arrowstyle='->', color=C_ACCENT, lw=1.2))
ax.set_xlabel('小时')
ax.set_ylabel('缺货率（%）')
ax.set_title('叮咚买菜数据逐小时缺货率分布（"U型"曲线）')
ax.set_xticks(range(0, 24, 2))
ax.set_xlim(-0.5, 23.5)
ax.set_ylim(0, 50)
ax.grid(alpha=0.3, ls='--')
plt.savefig(f'{OUT}/图3-1_逐小时缺货率U型曲线.png')
plt.close()

# ============================================================
# 图2: 正常日 vs 缺货日 逐小时销量对比（删失问题的直观展示）
# ============================================================
print("图2: 正常日 vs 缺货日对比...")
# 找同一个门店-SKU的相邻两天：一天正常、一天严重缺货
df_sorted = df.sort_values(['store_id', 'product_id', 'dt'])
grp = df.groupby(['store_id', 'product_id'])
found = None
for (sid, pid), g in grp:
    g = g.sort_values('dt').reset_index(drop=True)
    if len(g) < 30:
        continue
    oos_days = g[g['stock_hour6_22_cnt'] >= 10]
    normal_days = g[(g['stock_hour6_22_cnt'] == 0) & (g['sale_amount'] > g['sale_amount'].median())]
    if len(oos_days) > 0 and len(normal_days) > 0:
        found = (g, normal_days.iloc[0], oos_days.iloc[0])
        break

g, nd, od = found
fig, axes = plt.subplots(1, 2, figsize=(11, 4.2), sharey=True)
xs = np.arange(24)

# 左：正常日
ax = axes[0]
ax.bar(xs, nd['hours_sale'], color=C_MAIN, alpha=0.85)
ax.set_title(f'正常日（门店{nd["store_id"]} · 商品{nd["product_id"]} · {nd["dt"]}）\n日销量 = {nd["sale_amount"]:.2f}，营业时段零缺货', fontsize=11)
ax.set_xlabel('小时')
ax.set_ylabel('小时销量（归一化）')
ax.set_xticks(range(0, 24, 4))
ax.grid(alpha=0.3, axis='y', ls='--')

# 右：缺货日
ax = axes[1]
colors = [C_ACCENT if s == 1 else C_MAIN for s in od['hours_stock_status']]
ax.bar(xs, od['hours_sale'], color=colors, alpha=0.85)
# 缺货时段阴影
status = od['hours_stock_status']
in_oos = False
for h in range(24):
    if status[h] == 1 and not in_oos:
        start = h; in_oos = True
    if (status[h] == 0 or h == 23) and in_oos:
        end = h if status[h] == 0 else h + 1
        ax.axvspan(start - 0.5, end - 0.5, alpha=0.12, color=C_ACCENT)
        in_oos = False
ax.set_title(f'缺货日（同门店同商品 · {od["dt"]}）\n日销量 = {od["sale_amount"]:.2f}（被删失压低！），缺货{od["stock_hour6_22_cnt"]}小时', fontsize=11)
ax.set_xlabel('小时')
ax.set_xticks(range(0, 24, 4))
ax.grid(alpha=0.3, axis='y', ls='--')
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=C_ACCENT, alpha=0.85, label='缺货时段'),
                   Patch(color=C_MAIN, alpha=0.85, label='有货时段')], fontsize=9, loc='upper right')

fig.suptitle('需求删失的直观表现：缺货时段销量被系统性归零', fontsize=13, y=1.02)
plt.tight_layout()
plt.savefig(f'{OUT}/图3-2_正常日vs缺货日对比.png')
plt.close()

# ============================================================
# 图3: 缺货小时数分布直方图
# ============================================================
print("图3: 缺货小时数分布...")
cnt = df['stock_hour6_22_cnt'].value_counts().sort_index()
pct = cnt / len(df) * 100

fig, ax = plt.subplots(figsize=(8, 4.5))
colors = [C_GREEN if v == 0 else (C_ACCENT if v >= 8 else C_ORANGE) for v in cnt.index]
bars = ax.bar(cnt.index, pct.values, color=colors, alpha=0.88)
for i, v in enumerate(pct.values):
    ax.text(cnt.index[i], v + 0.6, f'{v:.1f}', ha='center', fontsize=8.5, color='#333')
ax.axvline(7.5, color=C_ACCENT, ls='--', lw=1.2, alpha=0.7)
ax.text(7.6, 48, '严重缺货分界线（≥8小时）\n占17.5%', fontsize=9.5, color=C_ACCENT)
ax.set_xlabel('营业时段缺货小时数 stock_hour6_22_cnt')
ax.set_ylabel('样本占比（%）')
ax.set_title('营业时段缺货小时数分布（n = 4,500,000）')
ax.set_xticks(range(0, 17))
ax.grid(alpha=0.3, axis='y', ls='--')
plt.savefig(f'{OUT}/图3-3_缺货小时数分布.png')
plt.close()

# ============================================================
# 图4: 典型门店-SKU 90天销量时序（标注删失日）
# ============================================================
print("图4: 90天时序与删失日...")
g90 = g.sort_values('dt').reset_index(drop=True)
fig, ax = plt.subplots(figsize=(11, 4.5))
xs = pd.to_datetime(g90['dt'])
ax.plot(xs, g90['sale_amount'], color=C_MAIN, lw=1.8, marker='o', ms=3.5, label='日销量（观测值）', zorder=3)
oos_mask = g90['stock_hour6_22_cnt'] > 0
ax.scatter(xs[oos_mask], g90.loc[oos_mask, 'sale_amount'],
           color=C_ACCENT, s=55, zorder=4, marker='v', label='删失日（存在缺货）')
# 标注促销和节假日
promo = g90['discount'] < 1.0
if promo.sum() > 0:
    ax.scatter(xs[promo], g90.loc[promo, 'sale_amount'],
               facecolors='none', edgecolors=C_ORANGE, s=90, lw=1.5, zorder=4, label='促销日')
ax.set_xlabel('日期（2024年）')
ax.set_ylabel('日销量（归一化）')
ax.set_title(f'典型门店-SKU 90天销量时序（门店{g90["store_id"].iloc[0]} · 商品{g90["product_id"].iloc[0]}）：红点处真实需求被低估')
ax.legend(fontsize=10, loc='upper right')
ax.grid(alpha=0.3, ls='--')
fig.autofmt_xdate()
plt.savefig(f'{OUT}/图3-4_90天时序删失日.png')
plt.close()

# ============================================================
# 图5: 促销/节假日/天气对需求与缺货的影响（4子图）
# ============================================================
print("图5: 外部因素影响分析...")
fig, axes = plt.subplots(2, 2, figsize=(11, 7.5))

# 5a: 促销 vs 非促销
ax = axes[0, 0]
promo_sales = df[df['discount'] < 1.0]['sale_amount']
nopromo_sales = df[df['discount'] == 1.0]['sale_amount']
bp = ax.boxplot([np.log1p(nopromo_sales.sample(50000, random_state=1)),
                 np.log1p(promo_sales.sample(min(50000, len(promo_sales)), random_state=1))],
                labels=['非促销日', '促销日'], patch_artist=True, widths=0.5,
                medianprops=dict(color='black'))
bp['boxes'][0].set_facecolor(C_GRAY); bp['boxes'][0].set_alpha(0.6)
bp['boxes'][1].set_facecolor(C_ORANGE); bp['boxes'][1].set_alpha(0.6)
ax.set_title('(a) 促销对销量的影响', fontsize=11.5)
ax.set_ylabel('log(1+日销量)')
ax.grid(alpha=0.3, axis='y', ls='--')

# 5b: 节假日 vs 平日缺货率
ax = axes[0, 1]
hol = df.groupby('holiday_flag')['stock_hour6_22_cnt'].apply(lambda s: (s > 0).mean() * 100)
bars = ax.bar(['平日', '节假日'], hol.values, color=[C_MAIN, C_ACCENT], alpha=0.85, width=0.5)
for i, v in enumerate(hol.values):
    ax.text(i, v + 0.8, f'{v:.1f}%', ha='center', fontsize=11)
ax.set_title('(b) 节假日缺货发生率对比', fontsize=11.5)
ax.set_ylabel('缺货发生率（%）')
ax.set_ylim(0, max(hol.values) * 1.2)
ax.grid(alpha=0.3, axis='y', ls='--')

# 5c: 温度与销量关系
ax = axes[1, 0]
tmp = df[['avg_temperature', 'sale_amount']].sample(100000, random_state=1)
tbins = pd.cut(tmp['avg_temperature'], bins=12)
tmean = tmp.groupby(tbins, observed=True)['sale_amount'].mean()
xpos = [interval.mid for interval in tmean.index]
ax.plot(xpos, tmean.values, color=C_MAIN, marker='s', ms=5, lw=2)
ax.set_title('(c) 平均气温与销量关系', fontsize=11.5)
ax.set_xlabel('平均气温（℃）')
ax.set_ylabel('平均日销量（归一化）')
ax.grid(alpha=0.3, ls='--')

# 5d: 降水与缺货率
ax = axes[1, 1]
pr = df[['precpt', 'stock_hour6_22_cnt']].sample(200000, random_state=1)
pr['rain'] = (pr['precpt'] > 0.5).astype(int)
rain_oos = pr.groupby('rain')['stock_hour6_22_cnt'].apply(lambda s: (s > 0).mean() * 100)
bars = ax.bar(['无降水', '有降水'], rain_oos.values, color=[C_MAIN, '#5B9BD5'], alpha=0.85, width=0.5)
for i, v in enumerate(rain_oos.values):
    ax.text(i, v + 0.8, f'{v:.1f}%', ha='center', fontsize=11)
ax.set_title('(d) 降水与缺货发生率', fontsize=11.5)
ax.set_ylabel('缺货发生率（%）')
ax.set_ylim(0, max(rain_oos.values) * 1.2)
ax.grid(alpha=0.3, axis='y', ls='--')

fig.suptitle('外部协变量对需求与缺货的影响（RL状态空间设计依据）', fontsize=13)
plt.tight_layout(rect=[0, 0, 1, 0.97])
plt.savefig(f'{OUT}/图3-5_外部协变量影响.png')
plt.close()

# ============================================================
# 图6: 缺货持续性（连续缺货天数分布）—— RL状态设计的核心依据
# ============================================================
print("图6: 连续缺货持续性分析...")
g_sorted = df.sort_values(['store_id', 'product_id', 'dt'])
g_sorted['oos'] = (g_sorted['stock_hour6_22_cnt'] >= 4).astype(int)  # 缺货≥4小时记为缺货日
# 计算每个store-product的连续缺货段长度
streaks = []
for _, gg in g_sorted.groupby(['store_id', 'product_id']):
    arr = gg['oos'].values
    cur = 0
    for v in arr:
        if v == 1:
            cur += 1
        else:
            if cur > 0:
                streaks.append(cur)
            cur = 0
    if cur > 0:
        streaks.append(cur)
streaks = np.array(streaks)

fig, ax = plt.subplots(figsize=(8, 4.5))
maxlen = min(15, streaks.max())
vals = [(streaks == k).sum() for k in range(1, maxlen + 1)]
bars = ax.bar(range(1, maxlen + 1), vals, color=C_MAIN, alpha=0.88)
ax.set_yscale('log')
for i, v in enumerate(vals):
    if v > 0:
        ax.text(i + 1, v * 1.15, f'{v:,}', ha='center', fontsize=8, rotation=45)
ax.set_xlabel('连续缺货天数（缺货≥4小时/天）')
ax.set_ylabel('缺货段数量（对数刻度）')
ax.set_title(f'连续缺货段长度分布：{int((streaks >= 3).sum()):,} 段缺货持续≥3天\n—— "缺货积累效应"真实存在，是RL状态空间纳入缺货历史的实证依据', fontsize=12)
ax.grid(alpha=0.3, axis='y', ls='--')
plt.tight_layout()
plt.savefig(f'{OUT}/图3-6_连续缺货段分布.png')
plt.close()

print()
print("=" * 60)
print(f"全部图表已生成至: {OUT}")
for f in sorted(os.listdir(OUT)):
    size = os.path.getsize(f'{OUT}/{f}') / 1024
    print(f"  {f}  ({size:.0f} KB)")
