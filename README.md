# data-analysis

Excel BOM / 原材料成本自动化分析：命令行批量处理 + Web 上传分析。

## 项目结构

| 目录 | 说明 |
|------|------|
| `data/` | 待分析的 Excel（`.xlsx`），放好后运行脚本 |
| `scripts/` | 主分析脚本 `process_excel.py` 等 |
| `output/` | 分析结果 `bom_analysis_*.xlsx` |
| `web/` | Flask 网页版（上传、成本模拟等） |

## 快速开始（命令行）

```bash
python -m pip install -r requirements.txt
python scripts/process_excel.py
```

将导出表放入 `data/` 后执行。多张表时可选：

- 合并一份：`python scripts/process_excel.py`
- 每文件单独报告：`python scripts/process_excel.py --per-file`

## 原始表必需列

`产品编码`、`组件编码`、`MENGE`、`组件单价`（可选：`MEINS`、`MAKTX`）

## Web 版

```bash
cd web
python -m pip install -r requirements.txt
python app.py
```

或双击 `start_web.bat`（项目根目录）。

## 输出说明

结果 Excel 含五张工作表：产品总成本排名、BOM明细_占比与成本、组件全局成本排名、跨产品用量波动、各产品 Top3 成本组件。
