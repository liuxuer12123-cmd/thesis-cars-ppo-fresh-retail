const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
        ShadingType, PageBreak, Header, Footer, PageNumber, ImageRun } = require('docx');

// ===== 样式常量 =====
const SONG = "SimSun";      // 正文：宋体
const HEI = "SimHei";       // 标题：黑体
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const CW = 9026; // A4 内容宽度

// ===== 辅助函数 =====
function p(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: opts.font || SONG }, size: opts.size || 21, bold: opts.bold || false, italics: opts.italics || false })],
    spacing: { before: opts.before ?? 60, after: opts.after ?? 60, line: 360 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { firstLine: 420 } : undefined,
  });
}
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: HEI }, size: 32, bold: true })],
    spacing: { before: 360, after: 240 },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: HEI }, size: 28, bold: true })],
    spacing: { before: 280, after: 180 },
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: HEI }, size: 24, bold: true })],
    spacing: { before: 200, after: 120 },
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 21, bold: opts.bold || false })],
    spacing: { before: 40, after: 40, line: 360 },
  });
}
// 富文本段落（支持片段加粗）
function rich(parts, opts = {}) {
  return new Paragraph({
    children: parts.map(pt => new TextRun({ text: pt.text, font: { ascii: "Times New Roman", eastAsia: pt.font || SONG }, size: pt.size || 21, bold: pt.bold || false, italics: pt.italics || false })),
    spacing: { before: opts.before ?? 60, after: opts.after ?? 60, line: 360 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { firstLine: 420 } : undefined,
  });
}
function richBullet(parts) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: parts.map(pt => new TextRun({ text: pt.text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 21, bold: pt.bold || false })),
    spacing: { before: 40, after: 40, line: 360 },
  });
}
function cell(text, w, opts = {}) {
  const lines = String(text).split('\n');
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: opts.header ? { fill: "D9E2F3", type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: lines.map(line => new Paragraph({
      children: [new TextRun({ text: line, font: { ascii: "Times New Roman", eastAsia: SONG }, size: opts.size || 18, bold: opts.header || opts.bold || false })],
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 300 },
    })),
  });
}
function table(headers, rows, widths, opts = {}) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { header: true, size: opts.headerSize || 18, align: AlignmentType.CENTER })) }),
      ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, widths[i], { size: opts.size || 18, align: opts.align || AlignmentType.LEFT })) })),
    ],
  });
}
function note(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18, italics: true, color: "555555" })],
    spacing: { before: 60, after: 120, line: 300 },
  });
}
function spacer() { return new Paragraph({ children: [new TextRun("")], spacing: { before: 60, after: 60 } }); }

// ===== 图片嵌入辅助 =====
const FIGDIR = "/Users/liuxinyi18/Desktop/FreshRetailNet-50K/论文图表";
function pngSize(path) {
  const buf = fs.readFileSync(path);
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}
function fig(filename, caption, targetW = 600) {
  const fp = `${FIGDIR}/${filename}`;
  const { w, h } = pngSize(fp);
  const W = targetW, H = Math.round(targetW * h / w);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [new ImageRun({
        type: "png",
        data: fs.readFileSync(fp),
        transformation: { width: W, height: H },
        altText: { title: caption, description: caption, name: filename },
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: caption, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18, bold: true })],
    }),
  ];
}

// ===== 文档内容 =====
const children = [];

// ---------- 封面标题 ----------
children.push(new Paragraph({
  children: [new TextRun({ text: "硕士毕业论文选题报告", font: { eastAsia: HEI, ascii: "Times New Roman" }, size: 44, bold: true })],
  alignment: AlignmentType.CENTER, spacing: { before: 480, after: 240 },
}));
children.push(new Paragraph({
  children: [new TextRun({ text: "删失感知深度强化学习在生鲜零售补货策略中的研究", font: { eastAsia: HEI, ascii: "Times New Roman" }, size: 32, bold: true })],
  alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 },
}));
children.push(new Paragraph({
  children: [new TextRun({ text: "—— 基于叮咚买菜真实运营数据的实证研究", font: { eastAsia: HEI, ascii: "Times New Roman" }, size: 28, bold: true })],
  alignment: AlignmentType.CENTER, spacing: { before: 120, after: 360 },
}));
children.push(p("研究方向：供应链管理 / 强化学习 / 库存优化", { align: AlignmentType.CENTER }));
children.push(p("数据来源：叮咚买菜（Dingdong Inc.）真实运营数据", { align: AlignmentType.CENTER }));
children.push(p("报告日期：2026年8月", { align: AlignmentType.CENTER }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 一、论文要求 ----------
children.push(h1("一、论文要求"));
children.push(h2("1.1 选题方向要求"));
children.push(p("本论文须在以下方向范围内选题：物流管理相关选题、供应链管理相关选题。", { indent: true }));
children.push(h2("1.2 数据要求"));
children.push(p("论文必须使用真实企业数据，以保证研究的实用性与可信度。禁止使用完全虚构或合成的仿真数据作为唯一验证依据。", { indent: true }));
children.push(h2("1.3 实验验证要求"));
children.push(p("论文必须包含仿真回测（Simulation / Backtesting）验证模块，通过历史数据的时序回放评估所提方法的性能。", { indent: true }));
children.push(h2("1.4 方法论要求（核心约束）"));
children.push(p("工具和模型须以近年来新出现的方法作为核心技术手段，具体包括：强化学习（Reinforcement Learning，RL）、深度学习（Deep Learning，DL）、大语言模型（Large Language Models，LLM）。", { indent: true }));
children.push(rich([
  { text: "特别说明：", bold: true },
  { text: "传统方法（如EOQ、安全库存、(s,S)策略等经典运筹学方法，以及XGBoost、随机森林等传统机器学习算法）只能作为基准对比（Benchmark），不得作为论文的主要研究方法。" },
], { indent: true }));
children.push(h2("1.5 要求汇总"));
children.push(table(
  ["要求维度", "具体内容", "本选题满足情况"],
  [
    ["选题方向", "物流或供应链管理", "✅ 生鲜零售库存补货（供应链管理）"],
    ["数据要求", "真实企业数据", "✅ 叮咚买菜真实运营数据，898家门店"],
    ["实验验证", "仿真回测", "✅ 历史数据90天滚动回测，多基准对比"],
    ["核心方法", "RL / DL / LLM（至少一种）", "✅ 深度强化学习（PPO）作为核心方法"],
    ["传统方法角色", "仅作Benchmark", "✅ EOQ / (s,S) / newsvendor / XGBoost 仅作对比"],
  ],
  [2200, 3200, 3626]
));
children.push(spacer());
children.push(rich([
  { text: "数据合规性说明：", bold: true },
  { text: "本研究数据来自叮咚买菜（Dingdong Inc.）2024年3月至6月的真实运营记录，由其研究团队联合高校整理[Wang et al., 2025]。数据覆盖898家前置仓、865个生鲜SKU的逐小时销售与缺货状态，属于真实企业数据，满足论文对数据来源真实性的要求。" },
], { indent: true }));

// ---------- 二、数据集与企业问题定位 ----------
children.push(h1("二、数据集与企业问题定位"));
children.push(h2("2.1 叮咚买菜真实运营数据介绍"));
children.push(p("本研究数据为叮咚买菜（Dingdong Inc.）的真实生鲜零售运营数据[Wang et al., 2025]。该数据是目前全球首个附有缺货标注的删失需求时序数据，专为生鲜零售领域的需求预测与库存优化研究而整理，为删失需求研究提供了不可替代的实证基础。", { indent: true }));
children.push(h3("数据集核心统计"));
children.push(table(
  ["数据维度", "详细信息"],
  [
    ["数据来源", "叮咚买菜（Dingdong Inc.），中国生鲜到家电商龙头企业"],
    ["时间范围", "2024年3月至6月（约4个月，逐小时记录）"],
    ["门店数量", "898家前置仓，覆盖18个城市"],
    ["SKU数量", "863个生鲜SKU（蔬菜、水果、肉类等易腐商品）"],
    ["核心特色", "附有逐小时缺货状态标注（hours_stock_status），可准确识别删失样本"],
  ],
  [2400, 6626]
));
children.push(h3("关键字段说明"));
children.push(table(
  ["字段名", "类型", "含义"],
  [
    ["sale_amount", "数值", "日度实际销售量（删失值：缺货时被系统性压低）"],
    ["hours_stock_status", "24维0/1向量", "逐小时缺货状态（核心标注字段：1=缺货）"],
    ["stock_hour6_22_cnt", "整数（0-16）", "6:00-22:00营业时段缺货小时数（经hours_stock_status逐小时求和交叉验证，匹配率100%；越大代表缺货越严重）"],
    ["discount", "0/1", "是否有折扣促销（影响需求量，需前瞻性补货）"],
    ["holiday_flag", "0/1", "是否节假日（生鲜需求峰值场景）"],
    ["activity_flag", "0/1", "是否有大促活动"],
    ["precpt / avg_temperature", "浮点数", "降水量 / 平均气温（天气特征，影响生鲜品需求）"],
  ],
  [2400, 1800, 4826]
));
children.push(h3("数据集缺货实况统计（基于train集450万样本实测）"));
children.push(table(
  ["统计口径", "样本数", "占比"],
  [
    ["营业时段零缺货（cnt=0）", "2,507,994", "55.7%"],
    ["营业时段存在缺货（cnt>0）", "1,992,006", "44.3%"],
    ["严重缺货（cnt≥8，半天以上断货）", "789,250", "17.5%（与官方'约20%缺货'口径吻合）"],
    ["全天营业时段缺货（cnt=16）", "181,791", "4.0%"],
  ],
  [3826, 2600, 2600]
));
children.push(h2("2.2 企业背景与核心管理问题"));
children.push(p("叮咚买菜是中国领先的生鲜到家电商平台，以\"29分钟极速达\"为核心竞争力。每家前置仓面积约100-200㎡，每日凌晨集中补货一次，白天无法补货。这一模式对补货决策精度提出极高要求：", { indent: true }));
children.push(bullet("补货不足 → 当日断货，损失销售额，顾客流失"));
children.push(bullet("补货过多 → 生鲜品当日腐损，直接经济损失"));
children.push(bullet("历史数据失真 → 缺货日的销售记录被系统性压低（删失偏差），预测模型越训越偏"));
children.push(rich([
  { text: "核心管理问题：", bold: true },
  { text: "如何在历史销售数据存在缺货删失偏差的情况下，动态制定最优日度补货量，同时最小化缺货损失（商业损失）与腐损损失（实物损耗）？" },
], { indent: true }));
children.push(h2("2.3 删失需求问题的具体表现"));
children.push(table(
  ["场景", "实际发生的事", "系统记录的数据"],
  [
    ["正常日", "需求100斤，库存充足，卖出100斤", "销售额 = 100（准确）"],
    ["缺货日", "需求120斤，库存只有80斤，实际卖出80斤", "销售额 = 80（被压低！真实需求120丢失）"],
    ["后果", "拿80训练模型，模型认为需求只有80斤 → 下次补货80 → 又缺货", "恶性循环"],
  ],
  [1400, 4200, 3426]
));
children.push(p("本数据中的 hours_stock_status 字段提供逐小时缺货标注，使得识别删失样本和修正偏差成为可能，这是该数据的核心价值所在。", { indent: true }));
children.push(p("图2-1以门店0的商品4为例直观展示删失的发生机制：2024年3月30日为正常销售日（营业时段零缺货），日销量5.30；而3月28日该商品自早间起断货，营业时段累计缺货13小时，系统记录的日销量仅为0.50。两日真实需求水平相近，但观测销量相差逾10倍——缺货时段的需求被系统性归零。若直接以此类数据训练预测模型，模型将学到“该商品需求仅0.5”的错误信号，下次补货更保守，进而形成“缺货→低估→少补→再缺货”的恶性循环。", { indent: true }));
children.push(...fig("图3-2_正常日vs缺货日对比.png", "图2-1 正常日与缺货日的逐小时销量对比（门店0·商品4）", 620));

// ---------- 2.4 数据探索性分析（含图表） ----------
children.push(h2("2.4 数据探索性分析"));
children.push(p("为准确把握数据特征、支撑后续模型设计，本节围绕四个问题对数据展开探索性分析：（1）缺货在时间与规模上呈何种分布？（2）典型门店-SKU的销量时序呈现何种模式？（3）外部协变量如何影响需求与缺货？（4）缺货是否存在跨日持续性？", { indent: true }));
children.push(h3("2.4.1 缺货的时间分布特征"));
children.push(p("图2-2展示了全部450万个样本的逐小时缺货率。缺货率呈显著“U型”曲线：凌晨0-5点高达32.9%-41.1%（前一日库存耗尽、当日补货尚未到位），早6点集中补货后骤降至6.2%，随后随日间销售单调爬升，晚间21-23点重回41%以上。该规律表明删失偏差并非随机噪声，而是与“凌晨单次补货”运营模式强相关的系统性偏差——这决定了删失修正必须显式利用逐小时缺货标注，而非依赖一般性插补方法。", { indent: true }));
children.push(...fig("图3-1_逐小时缺货率U型曲线.png", "图2-2 逐小时缺货率分布（“U型”曲线）"));
children.push(p("图2-3进一步给出营业时段缺货小时数的样本分布（与2.1节缺货实况统计表对应）：55.7%的样本营业时段零缺货，44.3%存在不同程度缺货，其中17.5%缺货达8小时以上（即超过半个营业日处于断货状态），4.0%为全天缺货。删失样本占比近半，意味着若不做删失修正，训练数据中约一半样本的需求信号存在向下偏差，且偏差集中于中重度区间。", { indent: true }));
children.push(...fig("图3-3_缺货小时数分布.png", "图2-3 营业时段缺货小时数分布（n = 4,500,000）"));
children.push(h3("2.4.2 典型门店-SKU销量时序特征"));
children.push(p("图2-4展示门店0·商品4的90天销量时序，红色倒三角为严重删失日（缺货≥4小时）。可观察到两类典型模式：其一，销量尖峰日后常紧跟缺货日（如5月中旬促销高峰后的连续缺货），说明促销驱动的需求脉冲屡屡超出保守的补货量；其二，缺货日在时间轴上呈聚集分布，提示缺货存在跨日持续性。观测销量在删失日被显著压低，再次印证观测销量≠真实需求。", { indent: true }));
children.push(...fig("图3-4_90天时序删失日.png", "图2-4 典型门店-SKU 90天销量时序与删失日标注", 620));
children.push(h3("2.4.3 外部协变量对需求与缺货的影响"));
children.push(p("图2-5考察外部协变量的作用：（a）促销日销量中位数明显高于非促销日且右尾更厚，促销对需求存在显著拉升；（b）节假日与平日的缺货发生率几乎相同（44.4% vs 44.2%）——若需求预测不准是缺货主因，节假日（需求方差更大）的缺货率应显著更高，该事实说明当前缺货主要源于补货决策偏保守，而非预测能力不足，这正是以补货策略为优化对象的RL方法的切入空间；（c）销量随气温上升总体增长、27℃左右见顶回落，呈非线性关系；（d）降水日缺货率略高（44.2% vs 40.8%）。上述协变量（促销、节假日、气温、降水）将同时纳入需求预测特征与RL状态空间。", { indent: true }));
children.push(...fig("图3-5_外部协变量影响.png", "图2-5 外部协变量对需求与缺货的影响", 560));
children.push(h3("2.4.4 缺货的跨日持续性"));
children.push(p("图2-6统计连续缺货段（日缺货≥4小时）的长度分布：共识别出175,740段持续3天以上的连续缺货，最长达15天，段数随长度近似指数衰减。连续缺货的普遍存在证实了“缺货积累效应”——单日缺货不仅扭曲当日数据，还会通过库存连锁反应影响后续多日的观测信号。这一发现是将过去7天缺货时长序列纳入RL状态空间的直接实证依据：智能体只有显式感知缺货的持续积累，才能学到“连续缺货应激进补货”的跨期决策规律，而静态newsvendor类方法在原理上无法利用这一信息。", { indent: true }));
children.push(...fig("图3-6_连续缺货段分布.png", "图2-6 连续缺货段长度分布"));
children.push(rich([
  { text: "小结：", bold: true },
  { text: "探索性分析表明：（1）删失偏差是系统性、模式化的（U型曲线），可由缺货标注直接定位；（2）删失样本占比近半且中重度删失集中；（3）缺货具有显著跨日持续性；（4）促销、节假日、天气等协变量对需求有显著影响。这些发现共同支撑本研究的两层删失感知设计：预测层的删失似然修正，与决策层基于缺货历史的状态-奖励建模。" },
], { indent: true }));

// ---------- 三、现有研究综述（大幅扩充） ----------
children.push(h1("三、现有研究综述"));
children.push(p("本章系统梳理与本选题相关的学术脉络。生鲜/易腐品库存管理是运营管理领域的经典研究方向，自Nahmias（1982）的奠基性综述以来已积累四十余年研究。本章按照\"经典理论 → 前沿方法 → 交叉地带 → 直接竞品\"的逻辑，梳理六条研究主线，时间跨度覆盖1982年至2026年8月，涉及核心文献30余篇。", { indent: true }));

// ===== 3.1 经典易腐品库存理论 =====
children.push(h2("3.1 主线一：经典易腐品库存理论（1982-2019）"));
children.push(p("易腐品库存管理的理论研究始于上世纪八十年代，构成了整个领域的理论基石。该主线主要采用运筹学与随机优化方法，为后续所有研究提供了问题形式化与理论基础。", { indent: true }));
children.push(table(
  ["论文 / 作者", "时间", "核心贡献"],
  [
    ["Nahmias", "1982", "易腐品库存理论奠基性综述，建立基本问题框架"],
    ["van Donselaar et al. (IJPE)", "2006", "超市易腐品库存控制的实证研究"],
    ["Broekmeulen & van Donselaar", "2009/2019", "批量订货+时变需求下的易腐品启发式订货算法"],
    ["Minner & Transchel (OR Spectrum)", "2010", "服务水平约束下的周期盘点易腐品库存控制"],
    ["Karaesmen et al.", "2011", "易腐品库存管理系统性综述"],
    ["Chen et al. (POM)", "2014", "易腐品库存与定价联合协调优化"],
  ],
  [3400, 1200, 4426]
));
children.push(note("与本研究的关系：该主线提供了 EOQ、(s,S)、newsvendor 等经典基准策略，在本论文中仅作为对比基准（Benchmark）使用，符合论文方法论要求。"));

// ===== 3.2 RL用于库存与易腐品管理 =====
children.push(h2("3.2 主线二：强化学习用于库存与易腐品管理（2022-2026）"));
children.push(p("将强化学习应用于库存管理是近年来运营管理领域最活跃的研究前沿。该主线可细分为\"通用库存RL\"与\"易腐品专项RL\"两个子方向。", { indent: true }));
children.push(h3("3.2.1 通用库存强化学习"));
children.push(table(
  ["论文 / 作者", "时间", "核心贡献", "处理删失？", "真实数据？"],
  [
    ["Gijsbrechts et al.\nM&SOM", "2022", "DRL在lost-sales/双源/多级库存中的系统性验证，PPO vs 动态规划，奠基之作", "❌", "❌ 仿真"],
    ["Oroojlooyjadid et al.\nM&SOM", "2022", "DQN求解啤酒游戏库存博弈", "❌", "❌"],
    ["Boute et al.\nEJOR", "2022", "DRL库存控制研究路线图综述", "❌", "❌"],
    ["Qi et al.\nManagement Science", "2023", "端到端深度学习直接映射特征→补货量，JD.com实地部署", "❌", "✅ JD（非生鲜）"],
    ["Multi-agent DRL\n(PMC)", "2025", "多智能体DRL联合优化需求预测与库存管理", "❌", "❌"],
    ["Deep Controlled Learning\nEJOR", "2025", "深度控制学习求解库存控制问题", "❌", "❌"],
    ["ORPR\nJD.com+清华大学", "2026.1", "OR引导预训练+RL微调（RLOO），JD.com部署：持有成本-30%", "❌", "✅ JD（非生鲜）"],
    ["DeepStock\n芝加哥大学+Alibaba", "2026.3", "策略正则化DRL（Base-Stock结构嵌入PPO），Alibaba Tmall 100%部署（100万+SKU对）", "❌", "✅ 天猫（非生鲜）"],
  ],
  [2300, 900, 3326, 1300, 1200],
  { size: 17, headerSize: 17 }
));
children.push(h3("3.2.2 易腐品专项强化学习"));
children.push(p("针对生鲜/易腐品的RL研究近年来快速增长，但普遍在仿真环境或非真实生鲜数据上验证：", { indent: true }));
children.push(table(
  ["论文 / 作者", "时间", "期刊", "核心贡献", "处理删失？", "真实生鲜数据？"],
  [
    ["De Moor et al.", "2022", "EJOR", "奖励塑形改善生鲜品DRL性能，将启发式策略嵌入RL奖励函数", "❌", "❌ 仿真"],
    ["Yavuz & Kaya", "2024", "Applied Soft\nComputing", "DRL求解易腐品动态定价与库存联合决策，大状态/动作空间优化", "❌", "❌"],
    ["Sim + RL Food Waste\n(Sage)", "2024", "Simulation", "仿真+RL减少食品零售浪费，聚焦鲜品订货", "❌", "部分真实"],
    ["Agri-food DRL", "2024", "—", "农业食品供应链DRL库存优化，处理易腐性与季节性", "❌", "部分真实"],
    ["Nomura, Liu, Nishi", "2025", "Applied\nSciences", "DRL求解易腐品动态定价与订货策略联合优化", "❌", "❌"],
    ["Yavuz & Kaya", "2025", "ESWA", "DRL+随机环境预测的易腐品定价与库存控制", "❌", "❌"],
    ["Multi-perishable MARL\n(Qiao et al.)", "2025", "Supply Chain\nAnalytics", "多智能体RL求解多品类易腐品分布式动态定价", "❌", "❌"],
  ],
  [2100, 850, 1500, 2976, 800, 800],
  { size: 16, headerSize: 16 }
));
children.push(rich([
  { text: "共同局限：", bold: true },
  { text: "上述所有RL库存研究（无论通用还是易腐品专项）均假设历史需求数据准确可信，未处理缺货导致的删失偏差。在生鲜零售场景中，这会导致系统性训练偏差：缺货日的低销量被误当作低需求，RL智能体从而学到\"少补货\"的错误策略，形成恶性循环。特别需要指出的是，De Moor et al.（2022）虽已提出针对易腐品的奖励塑形（Reward Shaping），但其设计目标是将启发式先验嵌入奖励以加速收敛，", },
  { text: "与本研究提出的\"删失感知奖励塑形（Censoring-Aware Reward Shaping, CARS）\"在机制上根本不同", bold: true },
  { text: "：CARS的修正信号源自对历史缺货状态的显式观测，旨在纠正训练数据本身的系统性偏差，而非注入先验策略。" },
], { indent: true }));

// ===== 3.3 删失需求理论 =====
children.push(h2("3.3 主线三：删失需求理论（1994-2026）"));
children.push(p("删失需求（Censored Demand）理论研究如何从被库存截断的销售数据中恢复真实需求分布，是运筹学的成熟研究方向。", { indent: true }));
children.push(table(
  ["论文 / 作者", "时间", "核心贡献"],
  [
    ["Nahmias", "1994", "lost-sales系统需求估计的早期奠基工作"],
    ["Huh et al. (OR/MOR)", "2009/2011", "Kaplan-Meier估计量应用于删失库存控制，奠定统计方法基础"],
    ["Besbes & Muharremoglu (MS)", "2013", "理论量化删失对newsvendor问题的代价，建立理论下界"],
    ["Lu et al.", "2014", "删失下的需求估计与订货联合优化（贝叶斯方法）"],
    ["Chen & Chao (OR)", "2020", "删失需求下的动态定价与库存联合控制"],
    ["Primal-dual censored (OR)", "2023", "一仓多店系统删失需求库存控制，primal-dual算法"],
    ["Ding et al. (M&SOM)", "2024", "特征驱动的删失需求库存控制（非参数机器学习方法）"],
    ["Offline pricing censored (M&SOM)", "2025", "删失需求下离线特征定价的因果推断方法"],
    ["Chen et al.", "2025", "非平稳newsvendor：从无删失到有删失的学习算法设计"],
    ["Dynamic pricing censored (EJOR)", "2026", "删失需求下动态定价与学习的遗憾界分析"],
    ["PIC-RL（香港城大）", "2026.2", "将RL与删失结合（Prediction-Induced Censoring框架），用于云资源分配，非零售场景"],
  ],
  [3200, 1400, 4426],
  { size: 17, headerSize: 17 }
));
children.push(rich([
  { text: "共同局限：", bold: true },
  { text: "经典删失理论偏重静态统计修正；2023-2026年的新方法虽引入机器学习，但多聚焦定价或单阶段决策。PIC-RL虽将RL与删失结合，但验证场景为云资源分配，与生鲜零售的需求分布、腐损约束、日度补货节奏完全不同，无法直接迁移。" },
], { indent: true }));

// ===== 3.4 删失与生鲜交叉研究（新增） =====
children.push(h2("3.4 主线四：删失与生鲜交叉研究（关键空白地带）"));
children.push(p("将删失需求修正应用于生鲜场景的研究极为稀少，目前仅发现一篇直接相关文献：", { indent: true }));
children.push(table(
  ["论文 / 作者", "时间", "期刊", "核心贡献", "局限"],
  [
    ["Miguéis et al.", "2022", "Journal of\nCleaner\nProduction", "鲜鱼零售：基于删失销售数据的机器学习需求预测，在葡萄牙零售商真实数据上验证，同时降低浪费并保证可得性", "仅做需求预测（预测层），无补货决策优化；使用传统ML（非RL），无法学习动态补货策略"],
  ],
  [1700, 800, 1700, 3000, 1826],
  { size: 17, headerSize: 17 }
));
children.push(rich([
  { text: "定位分析：", bold: true },
  { text: "Miguéis et al.（2022）证明了\"删失修正对生鲜需求预测的价值\"，但其止步于预测层。本研究与其形成\"预测-决策\"的层次互补与递进关系：Miguéis解决\"生鲜需求测得准\"的问题，本研究进一步解决\"测准之后如何动态补货\"的问题，并将删失感知从预测层延伸至决策层（RL训练闭环）。" },
], { indent: true }));

// ===== 3.5 FreshRetailNet-50K 相关研究 =====
children.push(h2("3.5 主线五：FreshRetailNet-50K 相关研究（2025-2026）"));
children.push(p("以下是截至2026年8月，基于与本研究同源的企业数据（FreshRetailNet-50K）发表的研究成果：", { indent: true }));
children.push(table(
  ["论文", "时间", "核心贡献", "局限 / 与本研究关系"],
  [
    ["FreshRetailNet-50K\n原始论文\n（叮咚+中山大学）", "2025.5", "发布数据集；提出两阶段pipeline（需求恢复→预测）", "两阶段解耦导致误差传播；无补货决策模块"],
    ["CADRE\n（山东女子学院+北京工商大学）", "2026.7", "端到端：TimeXer骨干 + 删失NB似然 + 日周期分解 + newsvendor决策头。WAPE 36.71%，废品率6.4%，服务水平94.7%", "补货决策为静态固定分位数newsvendor，无法自适应历史缺货规律；本研究的最直接对比基准"],
    ["SCOPE\n（港大+JD.com）", "2026.7", "端到端供应链协同决策框架：品类选择+补货周期+配送路由联合优化，在FreshRetailNet-50K与JD数据上验证", "聚焦跨部门协同（品类/周期/路由），不处理删失需求修正；与本研究问题正交，可作引用框架而非竞品"],
    ["LLM时序预测研究", "2026", "将LLM用于时序预测，在FreshRetailNet-50K上测试", "仅预测，无补货决策"],
  ],
  [2200, 900, 3300, 2626],
  { size: 17, headerSize: 17 }
));

// ===== 3.6 研究空白矩阵分析 =====
children.push(h2("3.6 主线六：研究空白矩阵分析"));
children.push(p("综合上述五条主线（共30余篇文献），从六个能力维度对代表性工作进行系统比对：", { indent: true }));
children.push(table(
  ["能力维度", "DeepStock\n(Alibaba)", "ORPR\n(JD.com)", "Yavuz &\nKaya", "Nomura\net al.", "De Moor\net al.", "CADRE\n(叮咚)", "PIC-RL\n(CityU)", "Miguéis\n(鲜鱼)", "本研究\n(目标)"],
  [
    ["生鲜腐损约束", "❌", "❌", "✅", "✅", "✅", "✅", "❌", "✅", "✅"],
    ["删失需求修正", "❌", "❌", "❌", "❌", "❌", "✅预测层", "✅", "✅预测层", "✅两层"],
    ["RL动态补货", "✅", "✅", "✅", "✅", "✅", "❌静态", "✅资源", "❌", "✅"],
    ["缺货历史纳入状态", "❌", "❌", "❌", "❌", "❌", "❌", "部分", "❌", "✅显式"],
    ["真实生鲜数据", "❌天猫", "❌JD", "❌", "❌", "❌", "✅叮咚", "❌云", "✅鲜鱼", "✅叮咚"],
  ],
  [1900, 800, 800, 800, 800, 800, 826, 800, 800, 700],
  { size: 16, headerSize: 15, align: AlignmentType.CENTER }
));
children.push(spacer());
children.push(rich([
  { text: "核心结论：", bold: true },
  { text: "在已检索到的30余篇相关文献中，尚无任何研究同时满足以下四个条件：（1）生鲜腐损建模；（2）删失需求修正；（3）RL动态补货决策；（4）缺货历史显式纳入决策状态。最接近的工作分别止步于不同维度：CADRE止步于静态决策、PIC-RL止步于非零售场景、Miguéis止步于预测层、De Moor等易腐品RL工作则完全回避了删失问题。这一交叉空白构成本研究的核心立足点。" },
], { indent: true }));

// ---------- 四、选题确定 ----------
children.push(h1("四、选题确定"));
children.push(h2("4.1 研究题目"));
children.push(p("《删失感知深度强化学习在生鲜零售补货策略中的研究 —— 基于叮咚买菜真实运营数据的实证研究》", { indent: true, bold: true }));
children.push(h2("4.2 核心研究问题"));
children.push(bullet("如何利用逐小时缺货标注构建无偏真实需求估计，消除RL训练数据中的删失偏差？"));
children.push(bullet("如何将缺货历史（连续缺货天数、历史缺货时长序列）纳入RL状态空间，使智能体学习跨期自适应补货规律？"));
children.push(bullet("相比静态newsvendor补货策略（CADRE）与经典启发式策略，RL动态决策在连续缺货、节假日高峰、促销事件等场景下是否表现出更显著的优势？"));
children.push(h2("4.3 研究框架概述"));
children.push(h3("阶段一：删失感知需求预测（基础层）"));
children.push(p("利用 hours_stock_status 字段识别删失时段，采用截尾回归（Tobit/NB似然）从被截断的销售数据中恢复真实需求分布，输出每日需求的概率预测（均值μ与不确定性σ）。此阶段参考CADRE的删失似然框架，视为已有工作的复现基础。", { indent: true }));
children.push(h3("阶段二：深度强化学习补货策略（核心创新层）"));
children.push(p("将补货决策形式化为马尔可夫决策过程（MDP）：", { indent: true }));
children.push(table(
  ["MDP组件", "具体定义"],
  [
    ["状态空间 S", "当前库存 + 预测需求μ/σ + 过去7天逐日缺货时长序列（stock_hour6_22_cnt[t-7:t]）+ 连续缺货天数 + 明日促销/节假日/活动标识 + 天气特征"],
    ["动作空间 A", "补货量档位 ∈ {0, 0.5x, 1x, 2x, 3x, 5x}（以7日平均日销量EOQ为单位）"],
    ["奖励函数 R", "R = −(α × 持有成本 + β × 缺货惩罚 + γ × 腐损成本)，参数参照生鲜品行业标准"],
    ["环境动态", "基于叮咚买菜历史需求流水构建仿真环境，逐日回放真实需求序列"],
    ["训练算法", "PPO（Proximal Policy Optimization），叠加删失感知奖励塑形机制（Censoring-Aware Reward Shaping, CARS）"],
  ],
  [2200, 6826]
));
children.push(h3("核心创新机制：删失感知奖励塑形（CARS）"));
children.push(p("本研究与现有RL库存工作的最关键差异在于：将缺货历史显式纳入状态，并设计专门的奖励修正项——当智能体观测到连续缺货且本次补货量不足时，额外给予负奖励信号，强化\"连续缺货应激进补货\"的决策规律；反之，在长期缺货后进行合理大量补货时给予正激励。", { indent: true }));
children.push(rich([
  { text: "与De Moor et al.（2022）奖励塑形的本质区别：", bold: true },
  { text: "De Moor的奖励塑形是将启发式策略（如base-stock）的决策建议作为附加奖励项嵌入，属于\"先验策略注入\"，目的是加速收敛；而CARS的修正信号源自对真实缺货标注数据的统计分析，属于\"数据偏差矫正\"，目的是消除训练数据中的系统性删失偏差。二者在动机、信号来源、作用机制上均不同，且可叠加使用。" },
], { indent: true }));
children.push(h2("4.4 创新性声明"));
children.push(p("基于扩充后的文献调研（30余篇），本研究的创新性可精确界定为三个层次：", { indent: true }));
children.push(table(
  ["创新维度", "已有工作的不足", "本研究的贡献"],
  [
    ["方法创新", "RL库存研究（DeepStock/ORPR/Yavuz & Kaya/Nomura等）均忽略删失偏差；De Moor的奖励塑形不含删失感知", "首次将删失感知机制嵌入RL训练闭环，提出删失感知奖励塑形（CARS），与既有奖励塑形机制形成正交互补"],
    ["状态设计创新", "CADRE静态newsvendor无法利用缺货历史跨期规律；现有RL库存工作状态空间均不含缺货积累信息", "将过去7天逐日缺货时长序列纳入RL状态，使智能体显式感知并响应缺货积累效应"],
    ["应用创新", "删失+RL工作（PIC-RL）仅验证云资源分配；删失+生鲜工作（Miguéis）止步于预测层", "首次将删失感知RL框架应用于真实生鲜零售场景，在叮咚买菜真实数据上完成\"预测层+决策层\"两层删失感知的完整验证"],
  ],
  [1600, 3800, 3626],
  { size: 17, headerSize: 17 }
));
children.push(spacer());
children.push(rich([
  { text: "两层删失感知框架（本研究的叙事主线）：", bold: true },
  { text: "本研究将\"删失感知\"明确区分为两个层次——预测层删失感知（CADRE与Miguéis已实现：用删失似然修正需求预测）与决策层删失感知（本研究首创：将缺货历史嵌入RL状态与奖励，影响补货决策本身）。这一层次划分清晰地界定了本研究相对CADRE等最直接竞争者的差异化贡献。" },
], { indent: true }));
children.push(h2("4.5 论文章节规划"));
children.push(bullet("引言（问题背景、研究意义、主要贡献、论文结构）"));
children.push(bullet("文献综述（经典易腐品库存理论、RL库存管理综述、删失需求理论综述、删失与生鲜交叉研究）"));
children.push(bullet("问题描述与数据分析（叮咚买菜业务场景、数据探索性分析、缺货删失统计分析）"));
children.push(bullet("删失感知需求预测模型（Tobit/NB删失似然、日周期特征提取、预测结果评估）"));
children.push(bullet("深度强化学习补货策略（MDP形式化建模、PPO训练框架、删失感知奖励塑形设计）"));
children.push(bullet("实验设计与结果分析（仿真回测实验、基准对比实验、消融实验、特殊场景深度分析）"));
children.push(bullet("结论与未来研究方向"));
children.push(h2("4.6 预期实验指标"));
children.push(table(
  ["评估指标", "CADRE基准（已有最优）", "本研究预期目标"],
  [
    ["废品率（腐损率）", "6.4%", "≤ 5.5%（优化≥15%）"],
    ["服务水平（满足率）", "94.7%", "≥ 95.5%（提升≥0.8 pp）"],
    ["连续缺货后恢复速度", "无专项对比数据", "恢复至正常库存天数比CADRE缩短≥30%"],
    ["WAPE（加权绝对误差）", "36.71%", "≤ 34%（促销/节假日场景）"],
  ],
  [3200, 3000, 2826]
));
children.push(h2("4.7 风险分析与应对"));
children.push(table(
  ["风险点", "具体描述", "应对策略"],
  [
    ["领域竞争激烈", "2026年该领域集中爆发（ORPR/DeepStock/CADRE/SCOPE），存在被跟进风险", "设定文献截止日期（2026年8月）并在论文中声明；持续跟踪arXiv关键词"],
    ["PPO训练稳定性", "稀疏奖励+长周期库存问题中PPO训练难度较大", "参考De Moor et al.奖励塑形与DeepStock策略正则化技术稳定训练"],
    ["基准对比压力", "必须在同一数据集上超越CADRE（WAPE 36.71%/废品率6.4%/服务水平94.7%）", "提前复现CADRE作为baseline；预留充足调参时间"],
    ["文献完整性", "生鲜+RL+删失文献分散于OR、CS、食品科学多个领域", "本报告3.1-3.6的六主线框架已系统覆盖；写作期持续补充"],
  ],
  [1800, 3800, 3426],
  { size: 17, headerSize: 17 }
));

// ---------- 五、主要参考文献 ----------
children.push(h1("五、主要参考文献"));
children.push(h3("数据集"));
children.push(p("[1] Wang Y, Gu J, Long L, et al. FreshRetailNet-50K: A Stockout-Annotated Censored Demand Dataset for Latent Demand Recovery and Forecasting in Fresh Retail[J]. arXiv:2505.16319, 2025.", { size: 18 }));
children.push(h3("直接竞争与相关工作"));
children.push(p("[2] CADRE Framework. Demand Hidden by Stockouts: Censored Demand Recovery and Decision-Focused Inventory Learning[J]. MDPI Sustainability, 2026, 18(15): 7642.", { size: 18 }));
children.push(p("[3] Miguéis V L, Pereira A, Pereira J, et al. Reducing Fresh Fish Waste While Ensuring Availability: Demand Forecast Using Censored Data and Machine Learning[J]. Journal of Cleaner Production, 2022, 359: 131852.", { size: 18 }));
children.push(p("[4] Liang Y, Cao X, Zhang P, et al. SCOPE: Supply-Chain Operations through Coupled Policies for End-to-End Coordination[J]. arXiv:2607.28488, 2026.", { size: 18 }));
children.push(h3("经典易腐品库存理论"));
children.push(p("[5] Nahmias S. Perishable Inventory Theory: A Review[J]. Operations Research, 1982, 30(4): 680-708.", { size: 18 }));
children.push(p("[6] van Donselaar K H, van Woensel T, Broekmeulen R, et al. Inventory Control of Perishables in Supermarkets[J]. International Journal of Production Economics, 2006, 104(2): 462-472.", { size: 18 }));
children.push(p("[7] Minner S, Transchel S. Periodic Review Inventory-Control for Perishable Products under Service-Level Constraints[J]. OR Spectrum, 2010, 32(4): 979-996.", { size: 18 }));
children.push(p("[8] Karaesmen I Z, Scheller-Wolf A, Deniz B. Managing Perishable and Aging Inventories: Review and Future Research Directions[M]. Planning Production and Inventories in the Extended Enterprise, Springer, 2011.", { size: 18 }));
children.push(h3("强化学习库存管理"));
children.push(p("[9] Gijsbrechts J, Boute R N, Van Mieghem J A, et al. Can Deep Reinforcement Learning Improve Inventory Management? Performance on Lost Sales, Dual-Sourcing, and Multi-Echelon Problems[J]. Manufacturing & Service Operations Management, 2022, 24(3): 1349-1368.", { size: 18 }));
children.push(p("[10] De Moor B J, Gijsbrechts J, Boute R N. Reward Shaping to Improve the Performance of Deep Reinforcement Learning in Perishable Inventory Management[J]. European Journal of Operational Research, 2022, 301(2): 535-545.", { size: 18 }));
children.push(p("[11] Oroojlooyjadid A, Nazari M, Snyder L V, et al. A Deep Q-Network for the Beer Game: Deep Reinforcement Learning for Inventory Optimization[J]. Manufacturing & Service Operations Management, 2022, 24(1): 285-304.", { size: 18 }));
children.push(p("[12] Boute R N, Gijsbrechts J, van Jaarsveld W, et al. Deep Reinforcement Learning for Inventory Control: A Roadmap[J]. European Journal of Operational Research, 2022, 298(2): 401-412.", { size: 18 }));
children.push(p("[13] Qi M, Shi Y, Qi Y, et al. A Practical End-to-End Inventory Management Model with Deep Learning[J]. Management Science, 2023, 69(2): 759-773.", { size: 18 }));
children.push(p("[14] Yavuz T, Kaya O. Deep Reinforcement Learning Algorithms for Dynamic Pricing and Inventory Management of Perishable Products[J]. Applied Soft Computing, 2024, 163: 111864.", { size: 18 }));
children.push(p("[15] Nomura Y, Liu Z, Nishi T. Deep Reinforcement Learning for Dynamic Pricing and Ordering Policies in Perishable Inventory Management[J]. Applied Sciences, 2025, 15(5): 2421.", { size: 18 }));
children.push(p("[16] Yavuz T, Kaya O. Dynamic Pricing and Inventory Control of Perishable Products by a Deep Reinforcement Learning Algorithm[J]. Expert Systems with Applications, 2025.", { size: 18 }));
children.push(p("[17] Xie Y, Hao X, Liu J, et al. DeepStock: Reinforcement Learning with Policy Regularizations for Inventory Management[J]. arXiv:2603.19621, 2026.", { size: 18 }));
children.push(p("[18] Zhao L, Yu X, Qi Y, et al. ORPR: An OR-Guided Pretrain-then-Reinforce Learning Model for Inventory Management[J]. arXiv:2512.19001, 2026.", { size: 18 }));
children.push(h3("删失需求理论"));
children.push(p("[19] Huh W T, Janakiraman G, et al. An Adaptive Algorithm for Finding the Optimal Base-Stock Policy in Lost Sales Inventory Systems with Censored Demand[J]. Mathematics of Operations Research, 2009, 34(2): 397-416.", { size: 18 }));
children.push(p("[20] Besbes O, Muharremoglu A. On Implications of Demand Censoring in the Newsvendor Problem[J]. Management Science, 2013, 59(6): 1407-1424.", { size: 18 }));
children.push(p("[21] Chen X, Chao X. Dynamic Inventory Control and Pricing Strategies for Perishable Products[J]. Operations Research / related venues, 2020.", { size: 18 }));
children.push(p("[22] Ding J, Huh W T, Rong Y. Feature-Based Inventory Control with Censored Demand[J]. Manufacturing & Service Operations Management, 2024, 26(3): 1157-1172.", { size: 18 }));
children.push(p("[23] Chen Y, Huang R, Liu C. Decision Support under Prediction-Induced Censoring[J]. arXiv:2602.18031, 2026.", { size: 18 }));
children.push(h3("决策导向优化"));
children.push(p("[24] Elmachtoub A N, Grigas P. Smart \"Predict, then Optimize\"[J]. Management Science, 2022, 68(1): 9-26.", { size: 18 }));
children.push(spacer());
children.push(p("—— 报告完 ——", { align: AlignmentType.CENTER }));

// ===== 组装文档 =====
const doc = new Document({
  styles: {
    default: { document: { run: { font: { ascii: "Times New Roman", eastAsia: SONG }, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: { ascii: "Times New Roman", eastAsia: HEI } },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: { ascii: "Times New Roman", eastAsia: HEI } },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: { ascii: "Times New Roman", eastAsia: HEI } },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })],
      })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/liuxinyi18/Desktop/硕士论文选题报告_修订版.docx", buffer);
  console.log("✅ 文档已生成: /Users/liuxinyi18/Desktop/硕士论文选题报告_修订版.docx");
});
