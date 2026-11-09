# 成绩报表统计分析（score-analysis-general）

把一个或两个「联考 / 期末成绩报表」变成可交付的统计结果：
单期做 13 张横断面表，两期做 8 张纵向增值表。零依赖数据，纯 pandas + openpyxl。

## 它解决什么问题

学校发的成绩报表通常是**两行合并表头**（第 1 行分组名、第 2 行「原始分/赋分/班名/校名」），
直接读会错位。本技能负责：识别表头 → 扁平化 → 12 层分析 → 输出多工作表 xlsx + 校验锚点。

另一个常见坑是**把原始分变化当进步**。本技能对两期数据同时给「原始分口径」和「标准分口径」，
并给出一个快速判别式：原始分上升率远离 50%、标准分上升率接近 50% → 变化来自试卷难度。

## 目录

```
SKILL.md                技能说明（给 agent 读）
references/SOP.md       方法手册：12 层分析、校验清单、20 条踩坑
scripts/cross_sectional.py   横断面分析（单期）
scripts/longitudinal.py      纵向增值分析（两期）
examples/               600 人合成数据 + 预期输出（可直接跑通）
```

English readers: see [README.en.md](README.en.md).

## 安装

把整个目录放进你的技能目录（例如 `~/.dsh/skills/`），或让 agent 直接读取 `SKILL.md`。

## 用法

```bash
# 单期：13 张表
python scripts/cross_sectional.py 成绩报表.xlsx [输出.xlsx]

# 两期：8 张表（原始分 + 标准分两套口径）
python scripts/longitudinal.py A期.xlsx B期.xlsx [输出.xlsx]
```

依赖：`pandas`、`numpy`、`openpyxl`（读 .xls 另需 `xlrd`）。

## 支持的报表形态

- 两行合并表头（最常见）
- 单行表头
- 三行表头（多一行标题）
- 县区列名兼容：县 / 县区 / 区域 / 区县

## 校验锚点（脚本会打印，必须人工核对）

| 校验 | 应满足 |
|---|---|
| 分数段合计 vs 校核 | 每科相等 |
| 贡献率合计 | 100.0% |
| 分差分解 | 四科之和 = 小计差 |
| 纵向恒等式 | 误差 < 0.5 的比例 > 99% |
| 配对率 | 姓名交集 / 较多一方 > 85% |

任何一项不过，先查数据，不要往下写报告。

## 先跑一遍示例

`examples/` 里有一份 600 人的合成数据（全部虚构）和一份预期输出：

```bash
python examples/make_example_data.py
python scripts/cross_sectional.py examples/样例_A期.xlsx
python scripts/longitudinal.py examples/样例_A期.xlsx examples/样例_B期.xlsx
```

预期看到「分数段合计不一致：无、贡献率 100.00%」，以及语文原始分 +8 分但标准分几乎为 0
——这就是难度识别的那个例子。详见 [examples/README.md](examples/README.md)。

## 数据

**本仓库不含任何真实成绩数据。** 见 [DATA-POLICY.md](DATA-POLICY.md)。

⚠️ 注意：**脚本生成的分析结果本身含个人信息**（「个人明细」有姓名/学校，「样本-学校」有校名）。
这些输出文件只应在本机使用，外发（发给家长、同事、群）前请自行脱敏或删除姓名列。

## 许可

MIT，见 [LICENSE](LICENSE)。
