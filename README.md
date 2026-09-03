# 硕士论文选题报告工作存档

**论文题目**：删失感知深度强化学习在生鲜零售补货策略中的研究——基于叮咚买菜真实运营数据的实证

## 目录结构

| 目录 | 内容 |
|---|---|
| `对话记录/` | 与AI助手的完整研究讨论记录（含数据验证、方法论设计、答辩问题推演全过程） |
| `报告文档/` | 选题报告各版本（原始版 → 修订版 → 完整版 → 答辩修订版/最终版） |
| `生成脚本/` | 报告与图表的生成脚本（docx由 `gen_thesis_v2.js` 生成，图由 `gen_figures.py` 等生成） |
| `论文图表/` | EDA图表（图4-1~4-6，300 DPI） |

## 数据集

- **FreshRetailNet-50K**：https://huggingface.co/datasets/Dingdong-Inc/FreshRetailNet-50K （主实验数据，90天）
- **FreshRetailNet-LT**：https://huggingface.co/datasets/Dingdong-Inc/FreshRetailNet-LT （长期数据，770天，用于泛化验证）
- 数据集未包含在本仓库（体积大），请从 HuggingFace 下载，许可 CC-BY-4.0

## 核心方法

- **预测层**：右删失负二项似然（Tobit思想），有货小时贡献 P(D=销量)，缺货小时贡献 P(D≥销量)
- **决策层**：CARS-PPO（删失感知奖励塑形的近端策略优化），状态含小时级缺货历史
- **验证协议**：①无缺货小时WAPE ②受控再删失半合成验证（两层掩码）③干净日真实轨迹仿真结算 + 生成器鲁棒性（规矩四）+ 缺货口径A-D敏感性分析

## 脚本运行

```bash
# 生成报告 docx（需全局安装 docx: npm i -g docx）
NODE_PATH=$(npm root -g) node 生成脚本/gen_thesis_v2.js

# 生成图表（需 pandas/pyarrow/matplotlib）
python3 生成脚本/gen_figures.py
```
