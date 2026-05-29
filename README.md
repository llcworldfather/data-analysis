# Excel BOM 成本分析（Web 版）

在浏览器中上传 BOM Excel，自动生成成本排名、占比、波动等报表，并支持 BI 看板与成本模拟。**数据在本机处理，不上传云端。**

## 项目结构

| 路径 | 说明 |
|------|------|
| [`web/`](web/) | Flask 应用：`app.py`、`process_excel.py`、页面模板 |
| `web/uploads/` | 上传临时文件（分析完成后通常已删除；异常残留会定期清理） |
| `web/results/` | 分析结果 `{session_id}.xlsx`（**默认保留 7 天**，过期自动删除） |
| [`bom_date_key.py`](bom_date_key.py) | 日期键规范化（Web 共用） |

## 快速开始

**方式一（推荐）**：双击项目根目录 [`start_web.bat`](start_web.bat)

**方式二**：

```bash
cd web
python -m pip install -r requirements.txt
python app.py
```

浏览器访问：<http://127.0.0.1:5000>（约 1 秒后自动打开；设 `BOM_WEB_NO_BROWSER=1` 可关闭）

从项目根安装依赖亦可：

```bash
python -m pip install -r requirements.txt
cd web && python app.py
```

## 使用流程

1. 打开首页，下载模板或上传 SAP/内控导出的 Excel（`.xlsx` / `.xls`）
2. 可选上传「产品价格历史清单」
3. 多个 BOM 文件时选择 **合并一份** 或 **每个文件单独报告**
4. 分析完成后进入 **BI 看板** / **成本模拟**，或下载 Excel

## 原始表必需列

`产品编码`（或 ZMATNR）、`组件编码`（或 IDNRK）、`MENGE`、`组件单价`（或 ZDJ）

可选：`MEINS`、`MAKTX`、`CREATEDATE`（生价日期）、`分类` 等

## 输出 Excel（6 张工作表）

1. 产品总成本排名  
2. BOM明细_占比与成本  
3. 组件全局成本排名  
4. 跨产品用量波动  
5. 各产品 Top3 成本组件  
6. 价格历史明细（供 BI 价格波动查询）

## 文件保留

- 默认保留 **7 天**（按文件修改时间）；可通过环境变量 `BOM_RETENTION_DAYS` 调整  
- 服务**启动时**及之后**每 24 小时**自动清理 `web/uploads/`、`web/results/` 中的过期文件

## 可选配置

| 环境变量 | 默认 | 说明 |
|----------|------|------|
| `BOM_RETENTION_DAYS` | `7` | `web/results/` 与侧车文件保留天数 |
| `BOM_BI_CACHE_MAX` | `8` | 内存中 BI 会话 LRU 上限（大会话占用内存较多） |
| `BOM_WRITE_EXCEL` | `1` | 设为 `0` 可跳过后台六表 xlsx，仅 BI 侧车 + 看板 |
| `BOM_WEB_NO_BROWSER` | — | 设为 `1` 启动时不自动打开浏览器 |
| `FLASK_DEBUG` | — | 开发调试（勿在生产开启） |

分析完成后会写入 `web/results/{session_id}_bi.json`（及可选 parquet），**重启服务后 BI 看板可从侧车恢复**，无需等待 Excel。

见 [`docs/predict_api_keys.md`](docs/predict_api_keys.md)（成本预测、DeepSeek 采购建议等 API Key）

### 开发测试

```bash
python -m pip install -r requirements-dev.txt
cd web && python -m pytest tests/test_bi_cache.py tests/test_bi_api.py tests/test_retention.py -q
```
