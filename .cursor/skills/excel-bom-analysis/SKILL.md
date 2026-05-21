---
name: excel-bom-analysis
description: >-
  Runs Polars-based BOM and raw-material cost analysis on Excel exports (产品编码,
  组件编码, MENGE, 组件单价). Use when the user works in this project, places .xlsx
  in data/, says 开始执行 or 【开始执行】, or asks about BOM / 成本驱动 / 用量占比 /
  bom_analysis output, Top3 组件 sheet, column meanings / 列说明.
---

# Excel BOM 成本分析（本项目）

## 项目布局

- **网页版（独立包）**：`web/` 目录自包含——`app.py`、`process_excel.py`、`templates/`、`requirements.txt`、`start_web.bat`；上传/结果在 `web/uploads/`、`web/results/`。不依赖项目根下的 `scripts/`。可选命令行：在 `web/` 下 `python process_excel.py`，读 `web/data/`、写 `web/output/`。
- `data/`（项目根）：命令行批量分析时，把待分析的 `.xlsx` 放在**项目根目录下的 `data` 文件夹**（与 `scripts`、`output` 同级）。用户若在对话中上传 Excel 且走命令行流程，助手应将其写入该 `data/` 后再跑脚本。
- 可多文件；Windows 下同名大小写重复只读一次。**`data/` 内有 2 个及以上 Excel 时**，助手在跑脚本前须**先问**用户：要 **合并一份**（`python scripts/process_excel.py`）还是 **每个 Excel 单独一份报告**（`python scripts/process_excel.py --per-file`）；仅 1 个文件时不必问。用户若已明说「合并」或「分开」，直接按其意执行。
- `scripts/process_excel.py`：主脚本（Polars + pandas 读表）；`--per-file` 时每源文件输出 `bom_analysis_<源文件名净化>_<时间戳>.xlsx`
- `output/`：`bom_analysis_*.xlsx`（默认合并为一份带时间戳；`--per-file` 则多个文件）
- `requirements.txt`：`pandas`, `openpyxl`, `polars`, `pyarrow`

## 用户说「开始执行」或「【开始执行】」时

在项目根目录执行：

1. `python -m pip install -r requirements.txt`（或 `py -m pip ...`）
2. 若 `data/` 下 **≥2 个** Excel：先让用户选合并或 `--per-file`，再运行对应命令；**1 个**文件则直接 `python scripts/process_excel.py`。
3. `python scripts/process_excel.py`（合并）或 `python scripts/process_excel.py --per-file`（分文件）

完成后用中文说明：`output` 下本次生成的 `bom_analysis_*.xlsx` 路径（**per-file 时列全**）；并**按规则**给出「五张工作表速览」（每张：用途 + 3～5 个核心列名；多份报告时结构相同，速览写一次即可）。产品/组件编码在 xlsx 中为文本。细则见 `.cursor/rules/excel-bom-assistant.mdc`「脚本成功后的汇报」。

## 输出表含义（给业务解读时）

- **产品总成本排名**：产品原材料总成本、核心成本组件及占该产品成本比例
- **BOM明细_占比与成本**：产品-组件行级用量占比与行成本
- **组件全局成本排名**：全数据集内组件总成本贡献
- **跨产品用量波动**：变异系数 CV 高表示该组件在不同产品间配比差异大
- **各产品Top3成本组件**：每产品成本前三组件（仅 xlsx 工作表）

## 原始表与结果表「列」说明（摘要）

用户问「哪些列有意义」时，除下表外**完整逐列释义**见 `.cursor/rules/excel-bom-assistant.mdc` 中的 **「表格列说明」**。

**原始表（必需）**：`产品编码`（成品）、`组件编码`（子件）、`MENGE`（用量）、`组件单价`（与用量单位一致时的单价）。**常用可选**：`MEINS`（单位）、`MAKTX`（组件描述/名称）。**常见忽略列**：如 `ZBJNO`、`ZSNO`、`JGLX`、`生产日期` 等导出字段，脚本不读。

**结果表共性**：`用量占比%` = 该组件用量占**该产品**总用量比例；`行原材料成本` = `MENGE合计` × `组件单价`；`核心成本组件` = 该产品下行成本最高的组件。

## 与 Rules 的关系

`.cursor/rules/excel-bom-assistant.mdc`（`alwaysApply: true`）含**开场白、执行流程、原始/结果表完整列说明**。**Cursor 只认 `.mdc` 规则文件**。本 Skill 补充技术路径与触发词。
