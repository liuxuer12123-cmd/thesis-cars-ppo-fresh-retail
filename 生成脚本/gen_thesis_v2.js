/**
 * 硕士论文选题报告生成脚本 v2 —— 答辩修订版
 * 调整：新增"问题分析"章（专硕三段式）+ 缺货口径合理性验证 + 敏感性分析设计
 *      + 缺陷驱动的RL论证 + 图3-3分桶版 + 保留完整版各项修改
 * 运行：NODE_PATH=$(npm root -g) node gen_thesis_v2.js
 */
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType, PageBreak,
        Header, Footer, PageNumber, ImageRun } = require("docx");

const SONG = "宋体", HEI = "黑体";
const OUT = "/Users/liuxinyi18/Desktop/硕士论文选题报告_答辩修订版.docx";
const FIGDIR = "/Users/liuxinyi18/Desktop/FreshRetailNet-50K/论文图表";

function pngSize(path) {
  const buf = fs.readFileSync(path);
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}
function fig(filename, caption, targetW = 600) {
  const fp = `${FIGDIR}/${filename}`;
  const { w, h } = pngSize(fp);
  const W = targetW, H = Math.round(targetW * h / w);
  const data = fs.readFileSync(fp);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 },
      children: [new ImageRun({ type: "png", data, transformation: { width: W, height: H },
        altText: { title: caption, description: caption, name: filename } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: caption, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18, bold: true })] })
  ];
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { line: 400, before: opts.before || 0, after: opts.after || 120 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.noIndent ? undefined : { firstLine: 480 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24, bold: !!opts.bold, color: opts.color })]
  });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: HEI }, size: 32, bold: true })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: HEI }, size: 28, bold: true })] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24, bold: true })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { line: 400, after: 80 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24 })] });
}
function rich(parts, opts = {}) {
  return new Paragraph({ spacing: { line: 400, after: 120 }, alignment: AlignmentType.JUSTIFIED,
    indent: opts.noIndent ? undefined : { firstLine: 480 },
    children: parts.map(t => typeof t === "string"
      ? new TextRun({ text: t, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24 })
      : new TextRun({ font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24, ...t })) });
}
function cell(text, width, opts = {}) {
  return new TableCell({ width: { size: width, type: WidthType.DXA }, shading: opts.head ? { fill: "D9E2F3", type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), font: { ascii: "Times New Roman", eastAsia: SONG }, size: 21, bold: !!opts.head })] })] });
}
function table(headers, rows, widths) {
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: { style: BorderStyle.SINGLE, size: 4 }, bottom: { style: BorderStyle.SINGLE, size: 4 },
      left: { style: BorderStyle.SINGLE, size: 4 }, right: { style: BorderStyle.SINGLE, size: 4 },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2 }, insideVertical: { style: BorderStyle.SINGLE, size: 2 } },
    rows: [ new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { head: true, align: AlignmentType.CENTER })) }),
      ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, widths[i])) })) ] });
}
function caption(text) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    children: [new TextRun({ text, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18, bold: true })] });
}
function note(text) { return p(text, { noIndent: true }); }
function spacer() { return new Paragraph({ spacing: { after: 200 }, children: [] }); }
function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

const children = [];

/* ============ 封面 ============ */
children.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2000, after: 400 },
    children: [new TextRun({ text: "硕士专业学位论文选题报告", font: { ascii: "Times New Roman", eastAsia: HEI }, size: 44, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({ text: "删失感知深度强化学习在生鲜零售补货策略中的研究", font: { ascii: "Times New Roman", eastAsia: HEI }, size: 36, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1200 },
    children: [new TextRun({ text: "——基于叮咚买菜真实运营数据的实证", font: { ascii: "Times New Roman", eastAsia: SONG }, size: 28 })] }),
  ...["专　　业：物流工程与管理", "研究方向：智慧供应链", "数据来源：叮咚买菜", "日　　期：2026年3月"]
    .map(t => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
      children: [new TextRun({ text: t, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 28 })] })),
  pageBreak()
);

/* ============ 一、论文要求 ============ */
children.push(h1("一、论文要求"));
children.push(h2("（一）现实背景"));
children.push(
  p("生鲜零售是高频刚需赛道，2024年中国生鲜电商交易规模达7367.9亿元、同比增长14.7%，前置仓模式已成为核心履约形态。但生鲜商品保质期以天计、需求波动剧烈，库存决策长期面临两难：订少了缺货，损失即时销售并伤害用户体验与复购；订多了当日废弃，腐损成本直接侵蚀毛利。行业生鲜腐损率普遍在5%—10%，补货决策水平是前置仓企业盈利能力的决定性变量。"),
  p("补货决策依赖需求预测，但需求数据存在一个长期被忽视的结构性缺陷——需求删失（Demand Censoring）：缺货时段没有成交，系统只能记录到“销量=0”，真实需求无从观测。观测销量是真实需求的右删失样本，用其训练的预测模型会系统性低估需求，进而压低订货量、加剧缺货，形成“越缺越低估、越低估越缺”的自强化恶性循环。"),
  p("本研究合作企业（国内头部前置仓生鲜电商）的逐小时运营数据显示：44.3%的样本日存在缺货，17.5%缺货超过8小时（重度），4.0%全天缺货。需求删失不是边缘噪声，而是该企业库存决策必须正面处理的核心数据特征，这为删失感知的智能补货研究提供了真实、必要且紧迫的业务土壤。")
);
children.push(h2("（二）数据"));
children.push(p("真实企业数据。叮咚买菜2024年第二季度的真实运营记录，覆盖898家前置仓、865个生鲜SKU、90天的逐小时粒度销售与缺货状态，含真实缺货标注（观测销量的右删失），直接支撑删失感知建模与仿真回测。无合成数据。"));
children.push(h2("（三）方法"));
children.push(p("以近端策略优化算法（PPO）为决策核心，属于典型的深度学习与强化学习方法，符合论文方法要求；创新点Censoring-Aware Reward Shaping（CARS，删失感知奖励塑形）建立于PPO框架之上。线性回归、指数平滑等经典统计模型仅作为基准对比，以验证智能方法的增益。"));
children.push(h2("（四）仿真回测验证"));
children.push(p("构建数字孪生仿真环境进行离线回测：基于历史真实数据构建生鲜零售仿真环境，模型输出的补货策略在仿真环境中闭环执行、滚动推演，以废品率、服务水平、WAPE、缺货恢复速度四项指标与基准策略对比，验证其在真实业务场景中的有效性。"));
children.push(h2("（五）规范性与合规性"));
children.push(p("本研究数据来自叮咚买菜2024年3月至6月的真实运营记录，数据覆盖898家前置仓、865个生鲜SKU的逐小时销售与缺货状态。"));
children.push(pageBreak());

/* ============ 二、企业背景与核心管理问题提出 ============ */
children.push(h1("二、企业背景与核心管理问题提出"));

children.push(h2("（一）企业背景与现行补货模式"));
children.push(
  p("研究对象为国内头部前置仓生鲜电商企业。其商业模式为“城市分选中心—前置仓—即时配送”：在全国多个城市布局前置仓，每仓覆盖周边3公里范围，主营蔬菜、水果、肉禽蛋、水产等短保生鲜品类，承诺最快30分钟送达。该模式下，每个前置仓每晚需对数百个SKU做出订货决策，订货在当日凌晨到仓、当日销售、当日清库，隔夜库存基本报废。"),
  p("企业现行的补货模式为典型的“预测→规则”两阶段范式：第一阶段由销量预测模型（基于历史销量的统计/机器学习模型）输出各SKU次日需求点估计；第二阶段在点估计基础上叠加人工设定的静态规则（安全库存系数、品类加成、促销临时上调等）生成订货量。该范式结构清晰、易于执行，在业务早期支撑了快速扩张，但随着仓网密度与SKU复杂度上升，其决策质量瓶颈日益凸显。")
);

children.push(h2("（二）现行模式下的管理困境"));
children.push(
  p("从管理视角看，企业当前面临的核心困境是缺货与腐损的“双高”挤压："),
  bullet("缺货侧：运营数据显示，44.3%的（门店×商品×日）样本存在缺货，17.5%缺货超过8小时、4.0%全天缺货；且缺货呈明显U型分布，晚间18—21点销售高峰恰是缺货高发时段，缺货对GMV与用户体验的伤害被进一步放大。"),
  bullet("腐损侧：为对抗缺货，部分门店采取“多订保险”策略，短保商品当日未售即废弃，腐损成本直接侵蚀毛利，行业内头部企业废品率水平约6.4%，每降低一个百分点都是可观的利润释放。"),
  bullet("管理侧：缺货与腐损此消彼长，门店与品类之间的订货宽严程度依赖店长个人经验，缺乏统一、可复制的决策标准；大促、节假日等外生冲击期间，人工规则调整滞后，波动进一步加剧。")
);

children.push(h2("（三）核心管理问题提出"));
children.push(
  p("上述困境可凝练为本研究的核心管理问题：在需求高度不确定、商品当日清库的约束下，企业如何在“缺货损失”与“腐损成本”之间实现动态最优平衡，使补货决策从依赖人工经验的静态规则，升级为数据驱动的动态智能决策？"),
  p("该问题在学术上对应“易腐品库存的动态补货决策”这一经典问题，但其现实特殊性在于：决策所依赖的需求数据本身因缺货而被系统性删失，使得“先预测、再优化”的传统范式在数据源头上就存在偏差。下一章基于企业真实运营数据，对这一问题的根源进行逐层剖析。")
);
children.push(pageBreak());

/* ============ 三、问题分析（新增章） ============ */
children.push(h1("三、问题分析：基于真实运营数据的归因诊断"));
children.push(p("本章遵循“先说明数据、再诊断根因、后映射对策”的逻辑展开：首先说明分析所用的数据；其次讲清数据中缺货是如何判定的，以及这一判定在数据中的具体表现；然后从数据中诊断出现行补货模式的三个深层根因；最后将根因映射为对解决方案的能力要求，从而自然引出第六章的技术路线。"));

children.push(h2("（一）分析数据与分析流程"));
children.push(
  p("分析基于企业898家前置仓、865个生鲜SKU共90天（2024年3月28日至6月25日）的逐小时运营记录，共450万条（门店×商品×日）样本。字段涵盖小时级销量、24维小时库存状态、折扣率、天气、节假日及门店与商品属性，核心字段如表3-1所示。"),
  caption("表3-1 核心数据字段一览"),
  table(
    ["字段", "含义", "本研究用途"],
    [
      ["sale_amount", "小时销量（真实成交）", "核心观测信号，删失建模对象"],
      ["hours_stock_status", "24维0/1向量，1=该小时缺货", "构建删失掩码，CARS奖励塑形"],
      ["stock_hour6_22_cnt", "6:00—22:00营业时段缺货小时数（0—16，越大越严重）", "缺货严重程度统计口径"],
      ["discount", "小时折扣率（≤1）", "识别清仓甩卖等价格冲击"],
      ["holiday_flag / 天气特征", "节假日标识、温度、降雨", "外生需求冲击控制变量"],
      ["dt", "日期，2024-03-28至06-25共90天", "训练/评估时间切分"]
    ],
    [2600, 3800, 3200]
  ),
  note("注：分析流程为“数据与缺货判定机制说明 → 缺货与删失现状刻画 → 根因诊断 → 对策映射”。")
);

children.push(h2("（二）缺货的判定机制：口径与数据表现"));
children.push(
  p("本研究全部后续分析都建立在“缺货标注”之上，因此本节首先讲清一个基础问题：数据中什么情形被判定为缺货，判定依据从何而来，这一判定在数据中呈现怎样的行为特征。"),
  h3("1. 判定口径：基于仓储作业的库存状态记录，而非销量推断"),
  p("数据中的缺货并非由“无销量”推断得出，而是来自仓储作业侧的客观记录：仓库管理系统按小时记录每个SKU的可售库存状态，可售库存为零时记为缺货，与前端的销量记录是两套独立的记录链路。即：缺货标注回答的是“此刻有没有货可卖”，销量回答的是“此刻卖了多少”，二者数据源独立。hours_stock_status字段即为该逐小时判定结果（1=缺货、0=有货）。"),
  h3("2. 数据表现一：判定与销量记录相互独立"),
  p("若判定如上述口径所述独立于销量，则数据中应存在大量“有货但零销量”的时段（真实无需求），并与“缺货且零销量”的时段（被掩盖的需求）清晰可分。对全部小时级记录的交叉统计（表3-2）正是如此："),
  caption("表3-2 库存状态×销量 交叉统计（小时粒度）"),
  table(
    ["象限", "占比", "业务解读"],
    [
      ["有货且有销量", "21.5%", "正常成交"],
      ["有货但零销量", "53.6%", "真实无需求时段（生鲜SKU长尾、深夜低峰客观存在）"],
      ["缺货且零销量", "24.2%", "被掩盖的真实需求——删失样本，本研究处理对象"],
      ["缺货但有销量", "0.72%", "小时边界处库存清零前的最后成交，占比可忽略"]
    ],
    [2400, 1400, 5800]
  ),
  p("“有货但零销量”占比高达53.6%——判定并不随销量走：没卖出去不等于没货，这一象限清晰刻画了判定与销量两条链路的独立性。"),
  p("对表3-2中“缺货但有销量”的0.72%需作专门说明：这并非标注错误，而是小时粒度聚合的边界效应——若某小时内先成交、后售罄（如17:05成交、17:40售罄），该小时便同时带有销量与缺货标记。数据验证支持这一解释：连续缺货段第一个小时内有销量的比例为5.73%，而缺货段第三个小时及以后仅为2.31%，即“缺货却有销量”显著集中在缺货刚开始的边界小时；扣除边界效应后的残余占比不足全部小时的0.2%，属数据噪声级别，不影响标注可信度。"),
  h3("3. 数据表现二：判定跟踪真实库存事件"),
  p("凌晨0—5点大部分商品自然无成交，但该时段大量商品标注为“有货”；缺货率曲线恰恰在0点处最高（约41%）、早6点补货到仓后骤降至约6%——判定精确地在“补货到仓”这一库存事件时点跳变，而非随销量的有无波动。以油条为例：凌晨无销量时其状态为有货，只有前一日真正售罄的SKU才在凌晨保持缺货标注。判定的这一行为与库存的物理运动完全一致。"),
  h3("4. 数据表现三：零销量样本日的两类成因清晰可分"),
  p("对全日零销量的样本日进一步分解：73.5%的全天零销量样本日伴随全天缺货（销量缺失由缺货解释），仅23.5%是全天有货下的真实零需求。两类成因在数据中被清晰区分，为后续“哪些零销量该被修正”提供了客观依据。"),
  p("综上，缺货的判定机制是清晰的：以仓储系统的逐小时库存状态为准、与销量链路独立、随补货事件跳变。后续分析即以该标注作为删失掩码的依据。同时，考虑到作业层面可能存在标注误差，本文在6.5节设计了多口径敏感性分析，检验结论对判定口径的稳健性。")
);

children.push(h2("（三）根因诊断：现行模式为何失效"));
children.push(
  p("基于上述可信数据，现行“预测→规则”两阶段补货模式的失效可归结为三个层层递进的根因。"),
  h3("根因一：需求信号被删失污染，预测从源头失真"),
  p("预测模型的训练目标是拟合历史销量，但历史销量在缺货时段是真实需求的右删失观测。直接拟合观测销量会系统性低估需求，订货量随之偏低，缺货进一步加剧，形成“缺货→低估→少订→更严重缺货”的自强化恶性循环。数据中4.0%的全天缺货样本日和17.6万个连续缺货段（最长15天）表明，部分商品已陷入该循环的长期陷阱。这是现行模式最底层的数据根因——它决定了“更准的预测”必须首先解决删失，而非简单更换更强的预测模型。"),
  h3("根因二：静态规则决策，无法感知缺货的历史积累与动态后果"),
  p("现行规则决策是“当日独立”的静态映射：预测值×系数=订货量，既不记忆过去连续缺货造成的潜在需求积累，也不考虑当日决策对次日库存起点的动态影响。而生鲜补货天然是序贯决策问题——今日少订不仅损失今日销量，还会改变明日的初始库存与补货空间；连续多日缺货的SKU需要补偿性超订才能恢复服务水平。静态规则对此无能为力，这是决策层面的结构性根因。"),
  h3("根因三：外生冲击响应依赖人工，滞后且不一致"),
  p("促销折扣、节假日、天气突变等外生冲击会显著改变需求分布（数据验证：强促销日销量约为正常日数倍，节假日效应显著）。现行模式下，冲击响应依赖店长临时人工上调订货，调整幅度因人而异、调整时机滞后于冲击发生，导致促销期“缺货与爆仓并存”。这是执行层面的根因。"),
);

children.push(h2("（四）根因—对策映射：对解决方案的能力要求"));
children.push(
  p("三个根因共同界定了理想解决方案必须具备的能力，形成如表3-3所示的根因—对策映射。该映射表明：问题求解需要一个“能修正删失偏差、能做序贯动态决策、能自动响应外生冲击”的决策框架——这正是本研究选择删失感知深度强化学习路线的业务逻辑依据，而非方法偏好。"),
  caption("表3-3 根因—对策—能力映射表"),
  table(
    ["诊断根因", "对策方向", "对应技术能力"],
    [
      ["根因一：需求信号被删失污染", "先修正需求信号，再做决策", "删失感知的需求建模（Tobit/NB似然、补货/平移修正）"],
      ["根因二：静态规则无视序贯依赖", "将补货建模为马尔可夫决策过程", "深度强化学习（PPO），状态含缺货历史，奖励含长期后果"],
      ["根因三：外生冲击响应滞后", "冲击变量进入状态空间，策略端到端自动响应", "状态空间纳入促销/节假日/天气，CARS奖励塑形动态调节风险厌恶度"]
    ],
    [3200, 2800, 3600]
  )
);
children.push(pageBreak());

/* ============ 四、数据集与探索性分析 ============ */
children.push(h1("四、数据集与探索性数据分析"));

children.push(h2("（一）数据集介绍"));
children.push(
  p("本研究数据为叮咚买菜真实运营数据，覆盖898家前置仓、865个生鲜SKU、18个城市、90天（2024年3月28日至6月25日）的逐小时粒度记录，总数据量450万条（门店×商品×日）。每条记录包含小时级销量、24维小时库存状态、折扣率、天气、节假日以及门店类型、商品品类等属性字段。在生鲜零售易腐品库存优化领域，真实企业级、逐小时粒度、带真实缺货标注的大规模数据长期稀缺，本数据集较好地填补了这一研究基础。"),
  h3("1. 覆盖广度"),
  bullet("门店维度：898家前置仓，覆盖一线至三线城市，含多种店型，足以检验方法在不同经营条件下的鲁棒性。"),
  bullet("商品维度：865个SKU覆盖蔬菜、水果、肉禽蛋、水产、速食等核心生鲜品类，保质期1—3天，均为典型的易腐短保商品。"),
  bullet("时间维度：2024年3月28日至6月25日共90天，涵盖工作日、周末、清明、五一、端午等节假日及大促节点，需求场景完整。"),
  bullet("空间维度：18个城市，南北方气候差异与消费偏好差异可支撑泛化性分析。"),
  h3("2. 粒度与规模"),
  p("数据为小时级粒度：每条记录是（门店×商品×日）维度，包含当日24个小时的销售量与库存状态。训练集约450万条样本日记录，评估集约35万条记录，样本量足以支撑深度强化学习模型训练与统计显著性检验。"),
  h3("3. 关键特征字段"),
  p("核心字段见第三章表3-1。其中 stock_hour6_22_cnt 为营业时段（6:00—22:00，共16小时）缺货小时数，取值0—16，数值越大代表缺货越严重，是本文缺货统计的统一口径。"),
  h3("4. 数据的独特价值：真实缺货标注"),
  p("绝大多数零售数据仅记录销量，缺货时段的真实需求不可见且不可知。本数据通过24维小时库存状态字段提供真实逐小时缺货标注，使研究者能够：明确识别哪些销量是被删失的（需求>销量）、为仿真环境提供真实的库存动态参数、量化“缺货导致的需求损失”这一长期难以估计的量。缺货的判定机制已在第三章第（二）节详细说明。"),
  h3("5. 数据时效与扩展：长期数据的泛化验证计划"),
  p("本报告的全部统计分析基于2024年二季度（90天）数据完成。需要说明的是，2026年2月企业进一步开放了覆盖2023年6月至2025年7月（770天、2.3万条店品序列、1057家门店、576个SKU）的长期运营数据，字段口径与本研究数据完全一致（含24维小时缺货标注），缺货结构相近（48.2%的样本日存在缺货、5.5%全天缺货）。该长期数据覆盖两个完整季节周期，弥补了90天数据无法刻画冬季需求与跨季节动态的局限。本文计划在实证阶段以90天数据为主实验基准（保证与CADRE等已发表结果的可比性），并以长期数据开展跨季节、跨年份的泛化验证，检验CARS-PPO策略在季节漂移下的稳健性。")
);

children.push(h2("（二）缺货现状与需求删失的具体表现"));
children.push(
  p("按营业时段缺货小时数对450万条样本日统计（表4-1）：55.7%的样本日全天有货，44.3%存在不同程度缺货，其中重度缺货（≥8小时）占17.5%，全天缺货（16小时）占4.0%。"),
  caption("表4-1 营业时段缺货程度分布（样本日粒度）"),
  table(
    ["缺货程度", "缺货小时数", "样本占比"],
    [
      ["全天有货", "0 小时", "55.7%"],
      ["轻微缺货", "1—3 小时", "9.2%"],
      ["中度缺货", "4—7 小时", "17.5%"],
      ["重度缺货", "8—11 小时", "9.6%"],
      ["极重缺货", "12—15 小时", "3.9%"],
      ["全天缺货", "16 小时", "4.0%"]
    ],
    [2800, 2600, 2200]
  ),
  p("需求删失在数据中的具体表现可归纳为四类："),
  h3("1. 完全删失：整日零销量"),
  p("约4.0%的（门店×商品×日）样本整日零销量且全天缺货。此类样本在训练集中呈现为“销量=0”，若不修正，模型将学到“该商品无需求”的错误信号，后续订货量趋零，彻底锁死该商品。"),
  h3("2. 部分时段删失：销量曲线被人为削平"),
  p("更普遍的情况是部分小时缺货。生鲜电商的晚高峰（18—21点）恰是缺货高发时段，高峰需求被删失意味着模型低估的恰是最有价值时段的需求。对比分析显示：正常日销量曲线呈双峰形态，而重度缺货日的晚高峰峰值被显著削平（见图4-2）。"),
  h3("3. 连续缺货累积：恶性循环的微观基础"),
  p("数据中存在175,740个连续缺货段（同一门店同一SKU连续≥3天缺货），最长连续缺货达15天。连续缺货期间真实需求持续被掩盖，为恶性循环提供了微观证据，也说明删失不是随机噪声而是结构性、持续性现象。"),
  h3("4. 删失率与时变性"),
  p("缺货率呈显著U型分布：凌晨0时缺货率约41.1%（前一日晚间售罄后尚未补货），早间6时补货后降至约6.2%，随后白天缓慢累积上升，至晚间21时再升至约41.0%（见图4-1）。删失不是恒定背景噪声，而是随时间动态变化的过程——这正适合用强化学习的状态转移框架刻画，也是CARS机制需要感知“缺货历史”的直接依据。")
);

children.push(h2("（三）探索性数据分析（EDA）"));
children.push(
  p("为系统刻画缺货规律、识别建模要点，对全部450万条样本日记录（7,560万条小时记录）开展探索性数据分析，共形成六组关键发现。"),
  h3("1. 缺货率的日内U型分布"),
  ...fig("图3-1_逐小时缺货率U型曲线.png", "图4-1 缺货率的日内U型分布与营业时段"),
  p("凌晨0点缺货率约41%，源于前一日晚间售罄、补货尚未到仓；早6点补货后骤降至约6%，随后全天缓慢爬升，晚21点再回到约41%。灰色区间为补货到仓后的营业时段。两点建模启示：其一，缺货高度集中于营业前段与晚间高峰，削平的恰是最高价值的销售时段；其二，缺货率随时段系统性变化，需求删失是时变过程，补货模型必须感知时段与缺货历史，静态统计特征无法刻画。"),
  h3("2. 正常日与重度缺货日的销量曲线对比"),
  ...fig("图3-2_正常日vs缺货日对比.png", "图4-2 正常销售日与重度缺货日的小时销量曲线对比"),
  p("随机选取20个门店×商品组合，分别绘制其正常日（缺货<4小时）与重度缺货日（≥8小时）的小时销量均值曲线，阴影为四分位区间。正常日呈典型双峰形态（午间小高峰+晚间18—21点大高峰），重度缺货日晚高峰峰值被明显削平，部分样本在高峰时段销量直接归零——被削掉的部分即被删失的真实需求。该图直观说明：若直接以观测销量训练预测模型，模型学到的将是“被削平后的需求”，订货决策从源头被系统性带偏，这正是本文要解决的第一根因。"),
  h3("3. 缺货程度分布"),
  ...fig("图3-3_缺货小时数分布.png", "图4-3 营业时段缺货程度分布（样本日粒度）"),
  p("按营业时段缺货小时数将全部样本日分为六档：全天有货占55.7%，轻度至中度缺货（1—7小时）占26.7%，重度及以上缺货（≥8小时）合计达17.5%，其中全天缺货（16小时）占4.0%。口径说明：样本占比=落在该区间的（门店×商品×日）样本数÷全部样本日数（898家×865品×90天=450万）。缺货不是个别门店的偶发问题，而是近半数样本日都面临的常态约束，需求删失处理必须是模型的内生机制、而非事后补丁。"),
  h3("4. 90天销量与缺货的动态演化"),
  ...fig("图3-4_90天时序删失日.png", "图4-4 90天日均销量与缺货程度动态演化（叠加促销与节假日）"),
  p("上图：全部门店×商品的日均销量曲线（7日移动平均），圆点为强促销日（折扣率最高的10%样本），背景浅色区间为法定节假日。下图：同期的样本日平均缺货小时数。两点规律：其一，强促销日销量脉冲明显（约为正常日数倍），凸显促销等外生冲击对需求的强扰动；其二，节假日与促销期过后缺货小时数往往同步抬升，说明现行静态规则在外生冲击下响应不足，需求脉冲后库存被迅速击穿——这为根因三提供了直接的时间序列证据。"),
  h3("5. 缺货与环境因素的关联分析"),
  ...fig("图3-5_外部协变量影响.png", "图4-5 缺货程度与外部环境的关联"),
  p("按星期、气温、降水、节假日分组统计平均缺货小时数。工作日与周末差异有限，但节假日（清明、五一、端午）缺货显著抬升；气温与降水通过影响到店/到家需求结构对缺货产生间接影响。结论：缺货并非纯粹内生随机，而是受促销、节假日、天气等外生变量共同驱动——建模时需将这类协变量纳入状态空间，帮助智能体区分“自身订货失误导致的缺货”与“外部冲击导致的缺货”，这对奖励塑形策略的设计至关重要。"),
  h3("6. 连续缺货的持续性"),
  ...fig("图3-6_连续缺货段分布.png", "图4-6 连续缺货段的长度分布（对数坐标）"),
  p("统计同一门店同一SKU连续≥3天缺货的缺货段：共175,740段，其中3—7天为主，超过14天的长尾段存在，最长达15天。连续缺货是“缺货→低估→少订→更严重缺货”恶性循环的微观证据，也提示仿真环境中智能体需要经历足够长的时序才能学会从持续缺货中恢复，回合长度设计与课程学习策略需据此设定。"),
  p("EDA小结：缺货在本数据中呈高比例（44.3%）、强时变（U型）、可持续（最长15天连续缺货）且受外生变量驱动四大特征。它们共同决定了：补货模型必须内生处理需求删失、必须感知缺货历史、必须自动响应外生冲击——这正是第六章技术路线逐点回应的对象。")
);
children.push(pageBreak());

/* ============ 五、文献综述 ============ */
children.push(h1("五、现有研究综述"));
children.push(p("围绕“删失感知深度强化学习在生鲜零售补货策略中的研究”这一选题，本文系统梳理五条研究主线，并在此基础上指出现有研究空白与本研究的切入点。"));

children.push(h2("（一）需求删失的识别与 unbiased 估计"));
children.push(
  p("需求删失问题最早源于航空收益管理领域。Larson & Robinson（1994）在1,761个航空预订观测上首次实证删失需求对预测的系统影响；McGill（1995）系统比较EM算法、投影删失（Projection Detruncation, PD）与双删失样本（Doubly Censored, DC）三类方法；Queenan等（2007）基于喜来登酒店集团的真实数据发现，约80%的高峰时段存在需求删失，使用非约束需求可显著提升预测精度；Liu等（2008）进一步引入顾客行为因素改进删失修正。"),
  p("零售场景下，Yavuz & Kaya（2021）针对报纸摊型易腐品零售商构建“部分删失需求+固定生命周期+随机提前期”的报童模型，证明平均生命周期越短越需要加大订货；其2023年后续研究将需求删失与提前期不确定性联合建模，再次确认“考虑删失的预测显著优于不考虑删失的预测”。Miguéis等（2022）以鲜鱼零售为场景，用可解释机器学习量化缺货导致的删失需求，发现温度、工作日/周末与缺货持续时间是关键预测因子，并据此提出增订策略——该文与本研究同属生鲜场景，但止步于预测层，未涉及动态补货决策。"),
  p("方法论层面，Afaki & Kremer（2024）系统对比了需求修正与需求平移两大删失建模流派；Lu等（2025）在KDD上提出CensorIDF，用逆数据频率加权缓解爆款商品删失偏差，已在叮咚买菜生产环境部署；与本研究数据同源的最新工作CADRE（Yin等，2026）将Tobit式负二项删失似然嵌入TimeXer深度预测器，并叠加报童分位数决策损失，实现预测层删失修正与成本感知订货的端到端联合训练，在FreshRetailNet-50K上取得WAPE 36.71%、还原偏差−1.3%的当前最优结果，但其决策层仍为“每日独立、按预测需求分位数订货”的静态报童模式，未进入序贯决策框架。总体看，删失建模在预测层已较成熟，但“修正后的需求信号如何服务于动态序贯决策”在生鲜零售场景中仍未被充分回答。这一空白之所以指向强化学习：预测层方法的产物是修正后的需求信号，而信号要转化为决策价值，必须进入一个能消费该信号、并处理“今日决策影响明日库存起点”之序贯依赖的框架——RL的状态空间可直接承接删失掩码与修正需求，其试错机制不要求需求分布具有解析形式，恰是预测层成果的天然下游。")
);
children.push(h2("（二）易腐品库存与动态定价联合优化"));
children.push(
  p("易腐品库存理论以Nahmias（1982）的经典综述为起点，Karasmen等（2011）进一步系统化了易腐品库存模型体系。Herbon & Cai（2022）针对具有真实保质期的易腐品建立“动态定价+订货”联合模型，证明只优化定价或只优化库存均显著次优；Guo等（2021）进一步将生鲜度、价格与营销努力三维联合优化。"),
  p("近期该主线向数据驱动方向演进：Yuan等（2025）将联合动态定价与补货建模为泛化的马尔可夫决策过程，兼顾新鲜度依赖的需求、浪费与顾客等待成本；Zhang等（2024）针对多产品易腐品引入谱风险度量处理需求不确定性。该主线的共识是：生鲜库存决策必须是动态的、联合的，但鲜有工作处理需求删失导致的观测失真。且该主线依赖的解析动态规划要求需求分布已知且观测无偏——观测一旦被删失污染、协变量维度上升（天气、促销、数百SKU），精确求解即遭遇维度灾难；RL的无模型试错恰好绕开这两个前提，这是“动态联合优化”的学术共识最终需要借助RL落地的技术原因。")
);
children.push(h2("（三）需求预测：从统计方法到Transformer与联邦学习"));
children.push(
  p("生鲜需求预测经历了从统计方法到深度学习的完整演进。统计方法（ARIMA、指数平滑）至今仍是最常用基准；LSTM、Transformer等深度模型因能捕捉非线性与长期依赖而成为主流；M5竞赛进一步确立了树模型（LightGBM等）在层级零售预测中的强势地位。"),
  p("近期研究呈现两个值得关注的方向：一是复杂架构的收益递减——Kaya & Yavuz（2023）在真实生鲜数据上系统比较后指出，数据质量与特征工程的边际收益高于模型架构复杂度；二是隐私保护下的跨店协同——Pan等（2025）在AAAI上提出联邦多层级预测框架，在不共享原始销售数据的前提下实现多店协同预测。本研究将同时覆盖统计基准与深度模型作为对照组。需要指出的是，预测回答的只是“需求是多少”，而补货要回答的是“订多少”：两者之间隔着非对称的缺货—腐损成本结构、跨日库存动态与缺货反馈回路，“预测×静态系数”的映射无法消化这些结构，预测更准也不必然带来决策更好。将预测信号作为状态输入、以业务后果为奖励直接优化决策，正是RL相对“预测+规则”范式的结构性优势，也是本研究将预测层定位为决策层输入而非终点的原因。")
);
children.push(h2("（四）补货决策：从静态规则到强化学习"));
children.push(
  p("传统库存决策以报童模型、(s, S)策略及其扩展为代表。此类静态规则无法适应生鲜需求的剧烈波动与多目标权衡。强化学习因天然适配序贯决策，近年在库存补货领域快速发展：DeepStock（阿里巴巴）在天猫超市场景验证了DRL补货的可行性；SCOPE（香港大学+京东）将补货与仓库间调拨联合建模，用深度RL与图网络捕捉SKU间关联；Liu等（2024）用PPO处理易腐品的动态补货；Nomura等（2025）验证了RL在可变提前期下的鲁棒性，并系统比较了缺货成本设定对策略的影响。"),
  p("方法论层面，De Moor等（2022）系统比较了Q-learning、SARSA与PPO，并提出基于启发式的奖励塑形机制；Dehaybe等（2024）提出Selective DQN，用分位数模拟的选择性规划降低策略梯度方差。与本文最相关的是PIC-RL（Liu等，2026）：它在云计算服务场景中将需求删失直接编码进RL的状态与奖励。但PIC-RL面向云资源场景，未处理生鲜的腐损约束与时变删失结构——生鲜零售+删失感知RL仍属空白。需要辨明的是，现有DRL补货的缺陷不在RL路线本身，而在其输入信号将销量等同真实需求；修复该缺陷亦无需抛弃RL——RL的状态与奖励是两个可编程接口，将缺货历史编码进状态、将删失感知的惩罚调节编码进奖励，缺陷即可在框架内部消解。PIC-RL已在云场景验证该接口的可行性，本文所做的是将其引入生鲜腐损与时变删失结构之中。")
);
children.push(h2("（五）大语言模型与智能体方法在供应链的初步探索"));
children.push(
  p("LLM在库存领域的应用刚刚起步。GPT-4o-Inventory（2025）在5个SKU上验证了“GPT-4o+提示工程”的决策潜力，但成本与幻觉风险显著；InvBench评测（2025）指出LLM在稳定环境下可接近最优，但在多级供应链的动态环境中与专业RL差距明显；Yang & Birge（2025）证明LLM agent可通过反思机制在报童问题上收敛到最优；ORPR（京东，2025）则代表另一条路线——将运筹规则与强化学习结合，在SKU泛化上表现优异。"),
  p("该主线的证据反而强化了RL的定位：InvBench等评测显示，LLM在稳定环境下可接近最优，但在多级供应链的动态环境中与专业RL差距明显，说明“通用智能+提示工程”尚不能替代“针对序贯决策结构专门优化”的框架；生鲜补货强时序、强反馈、强成本非对称的特性恰是RL的主场。本研究将LLM与OR混合方法纳入对比基准，以全面验证深度强化学习路线的竞争力，同时追踪这一前沿方向作为未来工作。")
);
children.push(h2("（六）研究空白与本研究定位"));
children.push(
  p("将五条主线交叉分析（表5-1），可清晰识别本研究的研究空白："),
  caption("表5-1 现有研究矩阵分析"),
  table(
    ["研究对象", "预测层删失修正", "决策层删失感知", "生鲜/易腐品", "企业级真实数据验证"],
    [
      ["航空/酒店收益管理", "✓", "✗", "✗", "✓"],
      ["零售预测（含生鲜）", "✓", "✗", "部分", "部分"],
      ["易腐品库存理论", "✗", "✗", "✓", "部分"],
      ["DRL补货（DeepStock/SCOPE等）", "✗", "✗", "部分", "✓"],
      ["PIC-RL（云计算）", "✗", "✓", "✗", "✓"],
      ["CensorIDF（叮咚买菜）", "✓", "✗", "✗", "✓"],
      ["CADRE（Sustainability 2026）", "✓", "✗", "✓", "✓"],
      ["本研究", "✓", "✓", "✓", "✓"]
    ],
    [3100, 1600, 1600, 1400, 1900]
  ),
  p("进一步需要回答的是：这些空白为何恰好能被深度强化学习填补？这取决于RL的三项结构性能力。其一，序贯决策的原生框架：补货决策的本质是马尔可夫决策过程——今日订货改变明日初始库存，缺货后果跨日累积，RL以MDP为母语、策略以状态为输入天然携带历史记忆，而报童规则与静态映射在数学形式上就无法表达依赖历史的策略。其二，无模型的试错学习能力：传统动态规划要求需求分布解析已知且观测无偏，在删失污染与高维协变量下遭遇维度灾难；RL不需显式分布，直接在与仿真环境的交互中优化策略，与预测层的删失似然生成器天然衔接。其三，状态与奖励的可编程接口：删失感知需要两个落点——“让决策看见缺货历史”（状态接口）与“让缺货厌恶随情境动态变化”（奖励接口）；预测模型没有奖励接口，规则方法没有状态接口，唯有RL同时具备两者，这是删失感知能够端到端落地的结构性原因，也是本文CARS机制的设计空间所在。"),
  p("需要强调的是，本研究选择深度强化学习路线，并非预设方法偏好，而是由第三章根因诊断自然导出：根因一要求先修正删失偏差（否则任何决策都在失真信号上空转），根因二要求序贯动态决策框架（静态规则无法感知缺货积累），根因三要求端到端的冲击响应（人工规则滞后且不一致）。现有研究中，预测层删失修正只解决根因一；常规DRL补货只部分解决根因二且无视根因一；PIC-RL处理了删失感知但脱离生鲜腐损场景；与本研究数据同源的CADRE虽在生鲜场景实现了预测层删失修正与成本感知订货的联合（根因一），但其订货决策为静态分位数规则，不感知缺货历史、不做序贯优化（根因二、三未覆盖）。将“生鲜腐损约束+删失感知+序贯决策”统一于一个框架的工作尚属空白——这正是本研究的切入点与创新空间所在。")
);
children.push(pageBreak());

/* ============ 六、选题确定 ============ */
children.push(h1("六、选题确定"));
children.push(h2("（一）论文题目"));
children.push(p("删失感知深度强化学习在生鲜零售补货策略中的研究——基于叮咚买菜真实运营数据的实证"));
children.push(h2("（二）研究目标"));
children.push(
  p("针对第三章诊断出的三个根因，设定三个层层递进的研究目标："),
  h3("1. 学术目标"),
  p("构建“删失感知的需求建模+PPO动态补货决策”的双层框架，实现预测层删失修正与决策层删失感知的闭环，回答“如何利用右删失数据训练出无偏的补货策略”这一兼具理论与实践价值的问题。"),
  h3("2. 方法目标"),
  p("设计CARS（Censoring-Aware Reward Shaping）机制：当智能体处于高缺货历史状态时，动态上调缺货惩罚权重，引导其优先恢复服务连续性；当库存健康时，则侧重腐损控制。预期在废品率、服务水平、WAPE、缺货恢复速度四项核心指标上系统性超越统计基准与无删失感知的PPO消融版本。"),
  h3("3. 实践目标"),
  p("形成可直接对接企业现有WMS数据的完整补货决策方案，输出可操作的门店订货建议，为企业降本增效提供实证支撑。")
);
children.push(h2("（三）研究方案与技术路线"));
children.push(
  h3("1. 总体框架"),
  p("遵循“提出问题（第二章）→分析问题（第三章）→解决问题（本章）”的专硕研究逻辑，构建“预测修正层+决策优化层+仿真验证层”三层架构："),
  h3("2. 预测修正层：删失感知的需求信号恢复"),
  p("需求信号恢复采用与同数据SOTA方法CADRE相同的统计原理——右删失负二项似然（有货小时贡献P(D=销量)，缺货小时贡献生存概率P(D≥售罄点)），保证预测层建立在已被验证的方法家族之上；但角色与实现不同：CADRE中需求恢复是端到端预测的副产品，本文则显式输出修正后的需求信号，既作为决策层的状态输入，也作为仿真环境的需求生成器——从拟合的删失后验分布中滚动采样需求轨迹，供智能体在回测环境中交互学习。此外，CADRE的消融实验表明“先插补、再预测”的两阶段路线（WAPE 38.94%）明显逊于端到端删失似然（36.71%），故本文以删失似然为主轨，传统插补策略（品类均值、时间衰减、上下文感知）仅作为消融基线参与对比。预测层的基线对比将直接复用官方基线套件——恢复类的TimesNet、SAITS、ImputeFormer等与预测类的SSA、TFT、DLinear、Chronos-2零样本基础模型，保证对比口径与官方一致、结果与后续公开工作可比；其中基础模型类基线不含删失概念，预期在缺货时段仍存在系统性低估，构成本文删失感知路线的直接对照。"),
  p("此处必须回答一个关键问题：缺货并不代表顾客此刻一定有需求，那么如何确定缺货损失了多少需求？本文的基本立场是——缺货小时损失的具体数值永远不可观测，本文不追求逐小时的精确还原，而是将每小时真实需求视为随机变量D：从有货时段学习“门店×商品×时段×协变量”条件下的需求分布，再据此推断缺货时段的期望需求；似然构造上，有货小时贡献P(D=销量），缺货小时贡献P(D≥售罄点），删失信息以统计学严格的方式进入参数估计。该框架天然区分两种情形：凌晨时段长尾商品的需求分布集中于零，即使标记缺货，其期望损失也接近于零（不夸大）；晚高峰热销商品的需求分布显著为正，缺货时段的期望损失相应较大（不低估）。换言之，“此刻有没有需求”由时段、星期、促销、天气、历史同期销量等协变量决定，而非由缺货标记本身决定。决策层的CARS机制亦不依赖精确的还原数值——奖励塑形只需删失掩码（哪些小时的销量信号不可信）与估计出的需求分布，故对单点估计误差不敏感；本层的目标是消除“把缺货当零需求”造成的系统性低估，而非追求不可达的逐小时真值。"),
  h3("3. 决策优化层：CARS-PPO补货智能体"),
  bullet("状态空间：当前库存、缺货历史（连续缺货小时数、近7日小时级删失率）、修正后需求预测、价格/折扣、天气、节假日；"),
  bullet("动作空间：离散订货量（0至最大库存容量的若干档位），每日凌晨下单一次、当日清库、日内不改单——与前置仓每日一次补货的业务实际一致；"),
  bullet("奖励函数：销售额-订货成本-废品成本-缺货惩罚×CARS权重（权重随缺货历史动态调整）；"),
  bullet("算法：PPO，GAE优势估计，熵正则化探索。"),
  bullet("时间粒度：需求建模与状态感知为小时级（删失掩码与缺货历史按24小时刻画，保留晚间高峰被削峰的删失结构），订货决策为天级一次；奖励日内按小时滚动结算、日终汇总。"),
  h3("4. 仿真验证层：数字孪生回测环境"),
  p("基于历史数据构建轻量级仿真环境：训练阶段的需求由修正后的删失后验分布逐小时采样生成（仿真器内需求全程已知、可精确结算），库存动态按真实补货提前期与小时级消耗模拟；评估结算则锁定在全天无缺的“干净日”真实需求轨迹上进行，保证成绩单按真实数据而非模型生成数据计算。评估指标涵盖废品率、服务水平、WAPE、缺货恢复速度，与线性回归、指数平滑、无删失感知PPO消融版本在同一环境中全面对比。"),
  h3("5. 与CADRE等前沿方法的对比定位"),
  p("近期基于同一数据来源的CADRE方法（Yin等，Sustainability 2026）代表该数据上的当前最优水平：其以TimeXer为编码骨干，叠加负二项删失似然恢复头与报童分位数决策头，实现WAPE 36.71%、删失还原偏差−1.3%，并在干净日仿真中达到废品率6.4%、服务水平94.7%。需要强调的是，CADRE并非强化学习方法，其决策为“每日独立、按预测需求分位数订货”的静态报童规则。本研究与CADRE的核心区别在决策层：CADRE回答“明日需求分布是多少、按什么分位数订货”，是单日独立的静态决策；本文将补货建模为马尔可夫决策过程，智能体携带缺货历史状态做序贯决策，并通过CARS机制感知连续缺货的累积后果、学习从持续缺货中恢复的策略——真实缺货呈多日持续性（最长连续15天），这正是静态框架未覆盖而序贯框架擅长的问题。此外，CADRE原文亦指出其受控再删失验证床比真实连续缺货更浅、结论有待实地验证，本文的序贯仿真回测可作为互补证据。为保证可比性，本文采用与CADRE一致的评估协议：预测精度仅在无缺货小时上计算WAPE，删失还原能力用受控再删失半合成实验检验，补货效果在仿真环境中结算（见6.5节）。")
);
children.push(h2("（四）研究创新点"));
children.push(
  h3("1. 首次将需求删失显式编码进生鲜零售DRL的状态与奖励设计"),
  p("现有DRL库存管理研究（DeepStock、SCOPE等）均未处理需求删失；PIC-RL虽处理删失但面向云计算场景、未考虑生鲜腐损特性。本文首次在生鲜零售场景中将删失信息同时编码进状态空间（缺货历史特征）与奖励函数（CARS动态惩罚），使智能体能够区分“真无需求”与“被删失的需求”。"),
  h3("2. 提出区别于启发式奖励塑形的删失感知奖励塑形机制CARS"),
  p("De Moor等（2022）的奖励塑形基于领域专家手工设计的启发式规则，本文的CARS完全由数据驱动，基于真实缺货标注动态调节奖励权重，直接由可观测的缺货状态而非人工先验塑形奖励信号。"),
  h3("3. 构建“预测层删失修正+决策层删失感知”的双层闭环框架"),
  p("预测层用Tobit/NB似然修正删失偏差，决策层用CARS感知删失后果，两层通过共享的删失掩码与需求信号形成闭环。现有研究中，Afaki & Kremer（2024）、CensorIDF（2025）等仅在预测层处理删失，PIC-RL仅在决策层处理删失；CADRE（2026）虽将删失似然与决策损失联合训练，但其决策层为静态分位数订货规则——真实缺货小时不进入其决策损失、模型不含缺货历史状态、不做跨日序贯优化。本文决策层是携带缺货记忆、感知连续缺货累积后果的序贯智能体，与上述工作形成本质区别。"),
  h3("4. 基于真实企业数据建立首个生鲜零售删失感知RL基准"),
  p("填补生鲜场景下删失感知RL基准的空白，为后续研究提供可复现的评测基线。")
);
children.push(h2("（五）敏感性分析与循环论证防范设计"));
children.push(
  h3("1. 缺货判定口径的多口径敏感性分析"),
  p("如答辩讨论所指出，缺货判定标准本身是影响结论的重要前提。第三章已说明缺货的判定机制及其数据表现，本文仍将在实证部分进一步设置系统的敏感性分析，检验核心结论对判定口径的稳健性："),
  caption("表6-1 缺货判定口径敏感性分析设计"),
  table(
    ["口径", "判定规则", "设置目的"],
    [
      ["口径A（基准）", "WMS作业标注：hours_stock_status=1 即为缺货", "现行口径，判定机制见第三章"],
      ["口径B（宽松）", "有货但连续≥3小时零销量，追加记为疑似缺货", "捕捉标注可能遗漏的“有货无销售”异常（如商品损坏未上架）"],
      ["口径C（严格）", "当日缺货≥4小时才计为缺货日，短时段不计", "检验短时标注噪声对结论的影响"],
      ["口径D（纯销量推断）", "完全不使用库存标注，仅用销量序列推断缺货（模拟无标注数据集）", "量化真实标注的信息价值，证明数据优势"]
    ],
    [1800, 4400, 3400]
  ),
  p("在四种口径下分别训练CARS-PPO并对比四项核心指标。预期结论：口径A下模型性能最优（真实标注信息量最大），口径B/D的性能衰减幅度可量化“真实缺货标注”的价值；若各口径下核心结论（CARS优于无删失感知消融）方向一致，则表明本文方法对判定口径具有稳健性。"),
  h3("2. 删失需求估计的有效性验证（半合成实验）"),
  p("此外，删失需求估计本身的有效性将通过半合成实验严格验证：选取全天有货的样本日（真实需求完整可见），人为掩盖部分销量以模拟缺货，再用6.3节的估计方法还原被掩盖的需求，将还原值与真实值直接对比并报告还原误差。需要区分“掩码协议”与“恢复方法”：掩码协议是考题，恢复方法是答题人，二者独立。掩码设计分两层：第一层为单小时部分深度掩码（与CADRE的受控再删失协议一致，随机削减高需求小时销量的20%—60%并打上缺货标记），保证本文结果可与CADRE直接对比；第二层为连续时段整段掩码（模拟“某小时售罄后连续数小时缺货”的真实缺货形态，并按真实缺货率的时段分布加权），检验恢复方法在真实缺货场景下的能力——CADRE原文亦承认其单小时浅层掩码与真实连续缺货存在差距，第二层设计正是对这一评估空白的补充。由于真实值已知，估计准确与否可被客观检验——这是删失需求研究文献中的标准验证范式，也是从方法论上回应“如何知道缺了多少”的最终依据。"),
  h3("3. 循环论证的防范：三道防线"),
  p("针对“修正后的需求并非真实需求、仿真验证是否存在循环论证”的方法论质疑，本文设置三道防线：其一，训练与结算分离——删失后验生成的需求仅用于智能体训练期的探索环境，所有方法的最终指标均在干净日真实需求轨迹上结算，模型不用自己的产物给自己打分；其二，相对比较原则——所有方法在同一仿真环境中训练与结算，环境偏差对所有方法一致，本文只主张“CARS优于无删失感知消融”的相对增益，不主张运营级绝对数字；其三，需求生成器误设鲁棒性检验——分别以负二项、Tweedie与经验bootstrap三种生成器重训重测，仅当CARS增益跨生成器一致时，才将其归因于方法本身而非某个生成器的假设。加之不修正的替代方案（直接以删失销量训练）已被证实存在约8%的系统性低估，修正路线在证据意义上严格优于不修正。")
);
children.push(h2("（六）预期成果与评估指标"));
children.push(
  p("三种评估口径构成“无害—有效—有用”的递进证据链，各自回答不同问题、缺一不可：①无缺货小时WAPE——回答“删失感知机制在干净数据上是否有副作用”：该口径下删失似然退化为普通似然，预期与骨干模型打平即合格，因训练信号去偏带来的间接增益则属于额外红利；仅用这一口径无法展示本文价值，但缺少这一口径则无法排除机制本身的损害。②受控再删失还原——回答“被掩盖的需求能否找回”，是本文方法的主场。③干净日仿真结算——回答“补货决策是否真的更好”，是废品率、服务水平等四项业务指标的最终裁决依据，CARS决策层的价值仅在此口径下体现。"),
  caption("表6-2 预期成果指标体系"),
  table(
    ["维度", "指标", "定义", "目标"],
    [
      ["成本控制", "废品率", "废弃库存 / 总订货量", "≤ 6.4%（对齐CADRE水平）"],
      ["服务质量", "服务水平", "1 - 缺货时长 / 营业时长", "≥ 94.7%（对齐CADRE水平）"],
      ["预测精度", "WAPE", "加权绝对百分比误差", "≤ 36.71%（对齐CADRE水平）"],
      ["响应能力", "缺货恢复速度", "从连续缺货到恢复正常库存的平均天数", "相比基准策略显著缩短"],
      ["对比基准", "统计基准", "线性回归、指数平滑等经典方法", "全面超越"],
      ["消融验证", "无CARS的PPO", "相同架构、去除删失感知组件", "证明CARS组件的独立增益"],
      ["前沿对比", "CADRE等SOTA方法", "同数据来源下最新方法", "对比分析各自适用边界"]
    ],
    [1800, 1900, 3400, 2500]
  )
);
children.push(h2("（七）研究计划与预期贡献"));
children.push(
  caption("表6-3 研究计划（2026年3月—2027年3月）"),
  table(
    ["阶段", "时间", "工作内容", "交付物"],
    [
      ["文献与理论", "2026.03—04", "删失建模、易腐品库存、DRL文献深化", "文献综述、理论框架"],
      ["数据与建模", "2026.04—05", "EDA深化、需求修正模型、仿真环境搭建", "EDA报告、仿真环境"],
      ["模型开发", "2026.05—07", "CARS-PPO实现与调优、消融实验", "训练代码、中间结果"],
      ["实证分析", "2026.07—09", "全量回测、口径敏感性分析、与SOTA对比", "实证结果"],
      ["论文撰写", "2026.09—12", "初稿撰写、内部评审、修改", "论文初稿"],
      ["修改答辩", "2027.01—03", "预答辩、盲审修改、正式答辩", "学位论文"]
    ],
    [1900, 2200, 3400, 2100]
  ),
  h3("预期贡献"),
  p("理论贡献：为“需求删失环境下的易腐品动态补货”这一理论空白提供新的建模范式与求解方法。方法贡献：CARS机制可推广至网约车、酒店、航空等所有存在需求删失的动态定价/库存场景。实践贡献：直接对接企业WMS系统，输出可操作的补货决策建议。")
);
children.push(h2("（八）论文结构安排"));
children.push(
  table(
    ["章节", "内容"],
    [
      ["第1章 绪论", "研究背景与问题提出、研究意义、技术路线、论文结构"],
      ["第2章 理论基础与文献综述", "需求删失理论、易腐品库存理论、深度强化学习综述"],
      ["第3章 企业背景与问题分析", "企业现行补货模式诊断、缺货口径验证、根因分析（对应本报告二、三章）"],
      ["第4章 模型构建", "删失感知需求建模 + CARS-PPO补货智能体 + 仿真环境设计"],
      ["第5章 实证分析", "基准对比、消融实验、口径敏感性分析、与CADRE等SOTA方法对比"],
      ["第6章 结论与展望", "研究结论、管理启示、局限与未来方向"]
    ],
    [3000, 6600]
  )
);
children.push(h2("（九）风险分析与应对"));
children.push(
  caption("表6-4 研究风险与应对策略"),
  table(
    ["风险", "应对策略"],
    [
      ["PPO训练稳定性风险：奖励尺度不当或探索不足导致策略震荡", "采用奖励归一化与梯度裁剪，加入课程学习从简单场景起步，监控KL散度控制策略更新幅度"],
      ["基准对比压力：CADRE等方法已具较强性能", "强调CARS在缺货恢复速度和服务水平上的差异化增益，用消融实验证明删失感知的独立价值"],
      ["缺货口径争议风险：标注误差或口径敏感性影响结论可信度", "按6.5节设计执行四种口径的敏感性分析，以多口径一致结论支撑稳健性主张"]
    ],
    [4400, 5200]
  )
);
children.push(pageBreak());

/* ============ 七、参考文献 ============ */
children.push(h1("七、参考文献"));
const refs = [
  "Yin, C., Zheng, Y., & Kong, Z. (2026). Demand hidden by stockouts: Censored demand recovery and green waste-reduction forecasting for sustainable fresh-food consumption under emerging-market urbanization. Sustainability, 18(15), 7642.",
  "Larson, P. D., & Robinson, L. W. (1994). Demand censoring and its effect on inventory forecasting: An empirical study of airline bookings. Journal of the Transportation Research Forum, 33(2), 77-84.",
  "McGill, J. I. (1995). Censored regression analysis of multiclass passenger demand data subject to joint capacity constraints. Annals of Operations Research, 60(1), 209-240.",
  "Queenan, C. C., Ferguson, M. E., Higbie, J., & Kapoor, R. (2007). A comparison of unconstraining methods to improve revenue management systems. Production and Operations Management, 16(6), 729-746.",
  "Liu, P. H., Smith, D. R., & Orkin, S. B. (2008). Care to estimate the unconstrained hotel demand? Journal of Revenue and Pricing Management, 7(4), 315-331.",
  "Yavuz, M., & Kaya, O. (2021). Inventory management for perishable goods with fixed lifetimes, random leadtimes, and censored demand. European Journal of Industrial Engineering, 15(4), 469-488.",
  "Kaya, O., & Yavuz, M. (2023). Retail demand forecasting: A comparative study of deep learning and ensemble methods. International Journal of Production Research, 62(15), 5413-5433.",
  "Yavuz, M., & Kaya, O. (2023). Demand forecasting with partially censored data for inventory management of perishables. Central European Journal of Operations Research, 32, 287-306.",
  "Miguéis, V. L., Camanho, A. S., & Borges, J. F. (2022). Predicting demand for fresh fish retailing using interpretable machine learning. Omega, 110, 102638.",
  "Afaki, B., & Kremer, M. (2024). Improving demand forecasts in censored environments: A comparison of demand correction and demand shifting. SSRN Working Paper.",
  "Lu, Z., Zhang, H., Chen, X., Yu, X., Xie, X., Kong, X., & Yang, M. (2025). Aiming at the target: An inverse data frequency weighted loss for mitigating demand censoring bias. KDD 2025.",
  "Nahmias, S. (1982). Perishable inventory theory: A review. Operations Research, 30(4), 680-708.",
  "Karasmen, I. Z., Scheller-Wolf, A., & Deniz, B. (2011). Managing perishable and aging inventories: Review and future research directions. International Series in Operations Research & Management Science, 151, 393-436.",
  "Herbon, A., & Cai, X. (2022). Modeling of dynamic pricing, ordering and preservation investment for perishable products with a real lifetime. Engineering Optimization, 55(9), 1442-1461.",
  "Guo, Y., Wang, X., & Fan, Z. P. (2021). Joint optimization of dynamic pricing and replenishment for perishables with freshness-, price- and marketing effort-dependent demand. Kybernetes, 51(4), 1546-1567.",
  "Yuan, Y., Pang, Z., & Shen, Z. (2025). Joint dynamic pricing and inventory management of perishable products with freshness-dependent demand. Naval Research Logistics, 72(7), 1063-1084.",
  "Zhang, J., Li, S., & Ke, G. Y. (2024). Managing perishable inventory systems with multiple products: A spectral risk measure approach. INFOR: Information Systems and Operational Research, 62(3), 614-631.",
  "Pan, J., Li, J., Wang, X., & Zhang, Q. (2025). Federated multi-level forecasting for privacy-preserving multi-store retail demand prediction. AAAI 2025.",
  "Liu, Z., Zhang, Z., Wang, B., Zhou, Y., Yao, Z., Huang, H., & Li, W. (2024). Practical inventory management for perishable goods with deep reinforcement learning. Neurocomputing, 587, 127689.",
  "Nomura, Y., Takahashi, S., & Sato, K. (2025). Inventory management of perishable products using deep reinforcement learning under uncertainty. Algorithms, 18(7), 398.",
  "De Moor, B. J., Gijsbrechts, J., & Boute, R. N. (2022). Reward shaping to improve the performance of deep reinforcement learning in perishable inventory management. European Journal of Operational Research, 301(2), 535-545.",
  "Dehaybe, D., Catanzaro, D., Chevalier, P., & Lusby, R. M. (2024). A new deep reinforcement learning approach for the perishable inventory management problem. Transportation Research Part E, 192, 103778.",
  "Liu, Z., et al. (2026). PIC-RL: Censoring-aware deep reinforcement learning for cloud service provisioning. IEEE Transactions on Services Computing, in press.",
  "Vanroose, P., Boute, R., & Verguts, T. (2025). GPT-4o-Inventory: Assessing LLM performance on inventory decision-making. arXiv:2505.16891.",
  "van der Veen, A., et al. (2025). InvBench: Can LLMs learn to manage inventory? A benchmark for sequential decision-making. arXiv:2506.03608.",
  "Yang, Z., & Birge, J. R. (2025). Inventory management with LLM-based agents. arXiv:2509.14236.",
  "ORPR: Operations rules and deep reinforcement learning for inventory optimization. JD.com Technical Report, 2025.",
  "Ding, X., Zhang, Y., Bin, T., Wei, X., & Li, Z. (2024). DeepStock: A deep reinforcement learning approach to inventory management. CIKM 2024.",
  "Wu, J., Li, Q., Chen, L., & Zhang, D. (2025). SCOPE: Supply chain optimization with graph-based deep reinforcement learning. KDD 2025 Applied Data Science Track."
];
refs.forEach((r, i) => {
  children.push(new Paragraph({ spacing: { line: 360, after: 80 }, indent: { left: 480, hanging: 480 },
    children: [new TextRun({ text: `[${i + 1}] ${r}`, font: { ascii: "Times New Roman", eastAsia: SONG }, size: 21 })] }));
});

const doc = new Document({
  styles: { default: { document: { run: { font: { ascii: "Times New Roman", eastAsia: SONG }, size: 24 } } } },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: { ascii: "Times New Roman", eastAsia: SONG } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1800, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "硕士专业学位论文选题报告", font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18, color: "666666" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: { ascii: "Times New Roman", eastAsia: SONG }, size: 18 })] })] }) },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`✅ 已生成: ${OUT}`);
  console.log(`   大小: ${(buf.length / 1024).toFixed(0)} KB, 图片数: ${(buf.toString('binary').match(/image\/png/g) || []).length}`);
});
