---
name: excel-bom-analysis
description: >-
  Web-only Polars BOM cost analysis. Use when the user works in this project,
  starts the Flask app, uploads Excel on the web UI, or asks about BOM / 成本驱动 /
  用量占比 / BI 看板 / web/results output / column meanings / 列说明.
---

# Excel BOM 成本分析（Web 仅）

## 项目布局

- **唯一产品入口**：[`web/`](web/) — `app.py`、`process_excel.py`、`templates/`、`requirements.txt`
- **上传**：浏览器 → `web/uploads/`（分析成功后一般删除）
- **结果**：`web/results/{session_id}.xlsx` + BI `/bi/<session_id>`；**默认保留 7 天**后自动清理
- **启动**：根目录 `start_web.bat`，或 `cd web && python -m pip install -r requirements.txt && python app.py`
- **依赖**：根目录 `requirements.txt` 指向 `-r web/requirements.txt`

**已无** 根目录 `scripts/`、`data/`、`output/` 命令行批量流程。

## 助手引导用户时

1. 说明这是 **Web BOM 分析**，数据在本机。
2. 请用户 **启动 Web**（`start_web.bat` 或上述命令），在 <http://127.0.0.1:5000> **上传 Excel**。
3. 多文件：页面上选 **合并** 或 **分开** 出报告（对应 API `mode=merge` / `per_file`）。
4. 用户在对话里附加 Excel 时：可协助 **启动服务** 或说明需在网页上传；**不要**再写入 `data/` 或跑 `scripts/process_excel.py`。

## 分析成功后的汇报（必须）

若用户已完成网页分析或你协助确认了 `web/results/` 下的文件，回复须包含：

**（一）结果位置** — `web/results/<session_id>.xlsx` 或 BI 链接；多文件 `per_file` 时列出各 token/下载项。

**（二）六张工作表速览**（每张：用途 + 3～5 个核心列名）：

1. **产品总成本排名** — `产品编码`、`原材料总成本`、`核心成本组件编码`、`核心成本组件占产品成本%`
2. **BOM明细_占比与成本** — `产品编码`、`组件编码`、`MENGE合计`、`用量占比%`、`行原材料成本`
3. **组件全局成本排名** — `组件编码`、`全局总成本贡献`、`全局成本占比%`、`涉及产品数`
4. **跨产品用量波动** — `组件编码`、`变异系数CV`、`涉及产品数`
5. **各产品Top3成本组件** — `产品编码`、`成本排名`、`组件编码`、`行原材料成本`
6. **价格历史明细** — `产品编码`、`组件编码`、`组件单价`、`CREATEDATE`（供 BI 价格波动）

列释义全文见 `.cursor/rules/excel-bom-assistant.mdc` **「表格列说明」**。

## 与 Rules 的关系

`.cursor/rules/excel-bom-assistant.mdc`（`alwaysApply: true`）含开场白、Web 流程、完整列说明。本 Skill 补充路径与触发词。
