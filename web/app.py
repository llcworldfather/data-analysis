"""
BOM 成本分析 Web 应用
======================
自包含：与本目录下 process_excel.py、templates/、requirements.txt 即可运行。

启动（请在 web 目录下，或 python 路径指向本目录下的 app.py）:
  双击 start_web.bat
  或: python app.py

访问: http://127.0.0.1:5000（约 1 秒后尝试自动打开浏览器）
不自动打开: 环境变量 BOM_WEB_NO_BROWSER=1
"""
from __future__ import annotations

import io
import json
import math
import os
import re
import threading
import time
import uuid
import traceback
import webbrowser
import zipfile
import html
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from datetime import datetime
import sys

from flask import Flask, request, jsonify, send_file, render_template, Response, stream_with_context
import pandas as pd
import polars as pl

from process_excel import (
    prepare,
    aggregate_lines,
    enrich,
    sheet_product_summary,
    sheet_bom_detail,
    sheet_component_global_rank,
    sheet_cross_product_volatility,
    top3_per_product,
    sheet_price_history,
    normalize_columns,
    norm_material_code,
    REQUIRED_COLS,
    COL_PRODUCT,
    COL_COMPONENT,
    COL_UNIT_PRICE,
    COL_CREATEDATE,
    COL_CATEGORY,
    _write_sheet,
)

app = Flask(__name__, template_folder="templates")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

WEB_DIR    = Path(__file__).parent
UPLOAD_DIR = WEB_DIR / "uploads"
RESULT_DIR = WEB_DIR / "results"
UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)

# BI 内存缓存（淘汰 + 写入须在同一锁内，避免并发竞态）
_BI_CACHE: dict = {}
_BI_CACHE_LOCK = threading.Lock()
_BI_CACHE_MAX = 8


def _bi_cache_put(token: str, entry: dict) -> None:
    """原子写入：容量淘汰与赋值在同一把锁内完成。"""
    with _BI_CACHE_LOCK:
        if len(_BI_CACHE) >= _BI_CACHE_MAX and token not in _BI_CACHE:
            oldest = min(_BI_CACHE, key=lambda k: _BI_CACHE[k]["ts"])
            del _BI_CACHE[oldest]
        _BI_CACHE[token] = entry


def _load_env_file(path: Path) -> None:
    """加载本地 .env，已存在的系统环境变量优先级更高。"""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_env_file(WEB_DIR / ".env")

_ROOT_PROJ = Path(__file__).resolve().parent.parent
if str(_ROOT_PROJ) not in sys.path:
    sys.path.insert(0, str(_ROOT_PROJ))
from bom_date_key import calendar_date_key_yyyymmdd as _calendar_date_key_yyyymmdd

from cost_sim_predict import (
    build_product_category_index,
    build_product_cost_history,
    build_product_price_timeline,
    expand_product_cost_history_dict,
    map_legacy_predict_en_to_zh,
    map_regression_analysis_en_to_zh,
    map_sensitivity_grid_en_to_zh,
    map_sensitivity_item_en_to_zh,
    normalize_predict_dataframes,
    normalize_price_list_cost_fields,
    predict_product_price,
    _predict_product_price_legacy,
    _safe_float,
    quote_month_key_from_series,
)

_UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$"
)


# ---------------------------------------------------------------------------
# Global JSON error handlers – 确保 API 永远返回 JSON，不返回 HTML 调试页
# ---------------------------------------------------------------------------

@app.errorhandler(404)
def err_404(e):
    return jsonify({"error": "接口不存在", "detail": str(e)}), 404

@app.errorhandler(405)
def err_405(e):
    return jsonify({"error": "请求方法不允许", "detail": str(e)}), 405

@app.errorhandler(413)
def err_413(e):
    return jsonify({"error": "文件过大：最大支持 100 MB"}), 413

@app.errorhandler(Exception)
def err_any(e):
    tb = traceback.format_exc().strip().split("\n")
    return jsonify({
        "error":  f"服务器内部错误：{e}",
        "detail": tb[-1] if tb else None,
    }), 500


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# API – Template download
# ---------------------------------------------------------------------------

@app.route("/api/template")
def download_template():
    """生成并返回标准 BOM Excel 模板（列格式与 SAP 导出一致）。"""
    df = pd.DataFrame({
        "ZBJNO":      ["2601070001", "2601070001", "2601070001", "2601070002", "2601070002"],
        "ZSNO":       ["2516001",    "2516001",    "2516001",    "2516002",    "2516002"],
        "IDNRK":      ["1010010004", "1010010293", "1010020025", "1010010004", "1010030011"],
        "MEINS":      ["KG",         "KG",         "KG",         "KG",         "PC"],
        "MENGE":      [600.0,        400.0,        20.0,         500.0,        30.0],
        "JGLX":       [3,            3,            1,            3,            1],
        "ZDJ":        [7.00,         5.55,         2.45,         7.00,         12.80],
        "ZMATNR":     ["8040120649", "8040120649", "8040120649", "8040120650", "8040120650"],
        "MAKTX":      ["原材料A-PE管材", "外购PE粒子", "色母粒", "原材料A-PE管材", "配件B"],
        "CREATEDATE": ["20260107",   "20260107",   "20260107",   "20260107",   "20260107"],
        "分类":        ["管材类",     "管材类",     "管材类",     "管材类",     "配件类"],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BOM数据")
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="BOM分析模板.xlsx",
    )


# ---------------------------------------------------------------------------
# API – Upload & Preview
# ---------------------------------------------------------------------------

def _xlsx_row_count(path: Path) -> int:
    """从 xlsx 快速获取数据行数（不读全部单元格）。
    策略1：读工作表 XML 前 4096 字节，解析 <dimension ref="A1:Z15000"> 元素（< 2ms）。
    策略2：若未找到则用 openpyxl read_only 的 max_row 属性（仍比全量读快很多）。
    """
    # ── 策略1：从 XML 元数据读取 ──────────────────────────────────────────
    try:
        with zipfile.ZipFile(str(path)) as zf:
            sheet = next(
                (n for n in zf.namelist()
                 if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
                None,
            )
            if sheet:
                with zf.open(sheet) as ws:
                    head = ws.read(4096).decode("utf-8", errors="ignore")
                # 匹配 ref="A1:BZ15000"，只需结束行号
                m = re.search(r'ref="[A-Za-z$]+\d+:[A-Za-z$]+(\d+)"', head)
                if m:
                    return max(0, int(m.group(1)) - 1)   # 减去表头行
    except Exception:
        pass

    # ── 策略2：openpyxl read_only（读 XML 元数据，不加载单元格）──────────
    try:
        from openpyxl import load_workbook
        wb = load_workbook(str(path), read_only=True, data_only=True)
        ws = wb.active
        nrows = (ws.max_row or 1) - 1
        wb.close()
        return max(0, nrows)
    except Exception:
        pass

    return 0


def _is_excel_filename(name: str) -> bool:
    n = (name or "").lower()
    return n.endswith(".xlsx") or (n.endswith(".xls") and not n.endswith(".xlsx"))


def _upload_save_ext(filename: str) -> str:
    """根据上传文件名决定落盘扩展名（支持模板 .xls / .xlsx）。"""
    n = (filename or "").lower()
    if n.endswith(".xlsx"):
        return ".xlsx"
    if n.endswith(".xls"):
        return ".xls"
    return ""


def _upload_bom_path(session_id: str, idx: int) -> Path | None:
    for ext in (".xlsx", ".xls"):
        p = UPLOAD_DIR / f"{session_id}_{idx}{ext}"
        if p.exists():
            return p
    return None


def _upload_price_path(session_id: str) -> Path | None:
    for ext in (".xlsx", ".xls"):
        p = UPLOAD_DIR / f"{session_id}_price{ext}"
        if p.exists():
            return p
    return None


def _read_xlsx_as_polars(path: Path) -> pl.DataFrame:
    """读取 Excel 第一张工作表（.xlsx / .xls）。
    .xlsx 优先 polars+fastexcel；.xls 或失败时回退 pandas（需 xlrd）。
    """
    if path.suffix.lower() == ".xls":
        return pl.from_pandas(pd.read_excel(path, sheet_name=0))
    try:
        return pl.read_excel(path)
    except Exception:
        return pl.from_pandas(pd.read_excel(path, sheet_name=0))


def _norm_product_code_series(s: pd.Series) -> pd.Series:
    """统一成品物料号为纯数字字符串，避免 Excel 浮点列变成 '8010040004.0' 导致与 BOM 对不齐。"""
    return norm_material_code(s)


# 产品价格历史清单模板：除「总成本→产品价格」外，下列列会一并并入 BOM 行（按产品/日期对齐）
_PRICE_EXTRA_COLS = ("材料成本", "工费", "重量", "报价月份", "标杆工厂")


def _attach_product_prices_pl(raw_pl: pl.DataFrame, price_path: Path) -> pl.DataFrame:
    """将「产品价格历史清单」左并入 BOM 明细行。

    业务语义（每一行）：某 **产品** 在某一 **报价月份**、某一 **重量(kg)** 下，
    **材料成本 / 工费 / 总成本均为 元/公斤**（写入列「产品价格」= 总成本每公斤）。

    对齐优先级：**产品+报价月份+重量** → **产品+报价月份** → **产品+日期** → **仅产品**（末条补全）。
    """
    raw_norm = normalize_columns(raw_pl)
    bom = raw_norm.to_pandas()
    drop_cols = ["产品价格", *_PRICE_EXTRA_COLS]
    bom = bom.drop(columns=[c for c in drop_cols if c in bom.columns], errors="ignore")

    price = pd.read_excel(price_path, sheet_name=0)
    if COL_PRODUCT not in price.columns:
        if "ZMATNR" in price.columns:
            price = price.rename(columns={"ZMATNR": COL_PRODUCT})
        elif "MATNR" in price.columns:
            price = price.rename(columns={"MATNR": COL_PRODUCT})
    if COL_PRODUCT not in price.columns:
        raise ValueError("产品价格表缺少列：产品编码（或 ZMATNR / MATNR）")

    val_col = None
    # 「总成本」优先：与产品价格历史清单模板一致
    for c in (
        "总成本",
        "产品价格",
        "标价",
        "出厂价",
        "销售价",
        "DMBTR",
        "NETPR",
        "KBETR",
        "NETWR",
    ):
        if c in price.columns:
            val_col = c
            break
    if val_col is None:
        raise ValueError(
            "产品价格表缺少金额列：总成本（或 产品价格 / 标价 / DMBTR 等）"
        )
    price = price.rename(columns={val_col: "产品价格"})

    merge_value_cols = ["产品价格"] + [c for c in _PRICE_EXTRA_COLS if c in price.columns]
    date_candidates = ("生价日期", "CREATEDATE", "生产日期", "创建日期", "日期")
    price_date_col = next((c for c in date_candidates if c in price.columns), None)

    bom[COL_PRODUCT] = _norm_product_code_series(bom[COL_PRODUCT])
    price[COL_PRODUCT] = _norm_product_code_series(price[COL_PRODUCT])
    for c in merge_value_cols:
        if c in price.columns and c not in ("报价月份", "标杆工厂"):
            price[c] = pd.to_numeric(price[c], errors="coerce")

    if "报价月份" in price.columns:
        price["报价月份"] = quote_month_key_from_series(price["报价月份"])
        price["_qm"] = price["报价月份"]
    elif price_date_col:
        price["_qm"] = quote_month_key_from_series(
            price[price_date_col].map(_calendar_date_key_yyyymmdd)
        )
    if "重量" in price.columns:
        price["重量"] = pd.to_numeric(price["重量"], errors="coerce")

    if "报价月份" in bom.columns:
        bom["报价月份"] = quote_month_key_from_series(bom["报价月份"])
        bom["_qm"] = bom["报价月份"]
    bom_date_col = COL_CREATEDATE if COL_CREATEDATE in bom.columns else None
    if "_qm" not in bom.columns and bom_date_col:
        bom["_qm"] = quote_month_key_from_series(
            bom[bom_date_col].map(_calendar_date_key_yyyymmdd)
        )
    if bom_date_col:
        bom["_dk"] = bom[bom_date_col].map(_calendar_date_key_yyyymmdd)
    if price_date_col:
        price["_dk"] = price[price_date_col].map(_calendar_date_key_yyyymmdd)
    if "重量" in bom.columns:
        bom["重量"] = pd.to_numeric(bom["重量"], errors="coerce")

    def _deduped_price(keys: list[str]) -> pd.DataFrame:
        keys = [k for k in keys if k in price.columns]
        if COL_PRODUCT not in keys:
            keys = [COL_PRODUCT, *keys]
        sort_keys = keys if keys else [COL_PRODUCT]
        out = price.sort_values(sort_keys, na_position="last").drop_duplicates(
            keys, keep="last"
        )
        cols = keys + [c for c in merge_value_cols if c in out.columns]
        return out[cols].dropna(subset=["产品价格"], how="all")

    def _coalesce_merge(base: pd.DataFrame, on_keys: list[str]) -> pd.DataFrame:
        on_keys = [k for k in on_keys if k in base.columns and k in price.columns]
        if COL_PRODUCT not in on_keys:
            return base
        use = _deduped_price(on_keys)
        if use.empty:
            return base
        m = base.merge(use, on=on_keys, how="left", suffixes=("", "_px"))
        for c in merge_value_cols:
            px = f"{c}_px"
            if c not in m.columns and px in m.columns:
                m[c] = m[px]
            elif c in m.columns and px in m.columns:
                m[c] = m[c].fillna(m[px])
            m = m.drop(columns=[px], errors="ignore")
        return m

    merged = bom.copy()
    merge_chain: list[list[str]] = []
    if "_qm" in merged.columns and "_qm" in price.columns:
        if "重量" in merged.columns and "重量" in price.columns:
            merge_chain.append([COL_PRODUCT, "_qm", "重量"])
        merge_chain.append([COL_PRODUCT, "_qm"])
    if "_dk" in merged.columns and "_dk" in price.columns:
        merge_chain.append([COL_PRODUCT, "_dk"])
    merge_chain.append([COL_PRODUCT])

    seen: set[tuple[str, ...]] = set()
    for keys in merge_chain:
        t = tuple(keys)
        if t in seen:
            continue
        seen.add(t)
        merged = _coalesce_merge(merged, keys)

    merged = merged.drop(columns=[c for c in ("_qm", "_dk") if c in merged.columns], errors="ignore")
    return pl.from_pandas(merged)


def _product_price_snapshot_from_bom_rows(rows: list[dict]) -> dict:
    """从 BOM 明细首行提取产品价格历史清单合并字段（供成本模拟页）；金额为 元/KG。"""
    if not rows:
        return {}
    r0 = rows[0]
    out: dict = {"价格口径": "元/KG"}
    if "产品价格" in r0:
        out["总成本"] = _optional_rounded_float(r0.get("产品价格"))
    for c in _PRICE_EXTRA_COLS:
        if c in r0:
            out[c] = _optional_rounded_float(r0.get(c))
    return out


@app.route("/api/upload", methods=["POST"])
def upload():
    """接收单个或多个 Excel 上传，校验列名，返回前 5 行预览数据。"""
    # 兼容多文件（files）和单文件（file）两种字段名
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        single = request.files.get("file")
        if single and single.filename:
            files = [single]
        else:
            return jsonify({"error": "未找到文件字段"}), 400

    # 校验所有文件格式（支持模板 .xls / .xlsx）
    for f in files:
        if not _is_excel_filename(f.filename or ""):
            return jsonify({
                "error": f"不支持的格式：{f.filename}",
                "hint":  "请上传 .xlsx 或 .xls 格式的 Excel（如 报价BOM历史清单）",
            }), 400

    session_id = str(uuid.uuid4())
    saved_paths: list[Path] = []
    price_path: Path | None = None

    try:
        for idx, f in enumerate(files):
            ext = _upload_save_ext(f.filename or "")
            p = UPLOAD_DIR / f"{session_id}_{idx}{ext}"
            f.save(str(p))
            saved_paths.append(p)

        price_file = request.files.get("price_file")
        has_price = False
        price_fname: str | None = None
        if price_file and price_file.filename and price_file.filename.strip():
            if not _is_excel_filename(price_file.filename):
                for p in saved_paths:
                    p.unlink(missing_ok=True)
                return jsonify({
                    "error": f"价格表格式不支持：{price_file.filename}",
                    "hint":  "产品价格表请上传 .xlsx 或 .xls（如 产品价格历史清单）",
                }), 400
            pext = _upload_save_ext(price_file.filename)
            price_path = UPLOAD_DIR / f"{session_id}_price{pext}"
            price_file.save(str(price_path))
            has_price = True
            price_fname = price_file.filename

        # 用第一个文件做列校验与预览
        # 只读 5 行数据（+表头），比读全文件快几十倍
        first_path = saved_paths[0]
        try:
            df = pd.read_excel(first_path, sheet_name=0, nrows=5)
        except Exception as exc:
            for p in saved_paths:
                p.unlink(missing_ok=True)
            if price_path is not None:
                price_path.unlink(missing_ok=True)
            return jsonify({"error": f"Excel 解析失败（第 1 个文件）：{exc}"}), 422

        df_pl   = normalize_columns(pl.from_pandas(df))
        df_norm = df_pl.to_pandas()
        missing = [c for c in REQUIRED_COLS if c not in df_norm.columns]
        if missing:
            for p in saved_paths:
                p.unlink(missing_ok=True)
            if price_path is not None:
                price_path.unlink(missing_ok=True)
            return jsonify({
                "error":  f"缺少必要列：{missing}",
                "detail": f"当前列名：{list(df.columns)}",
                "hint":   (
                    "必须包含（或其 SAP 导出别名）：\n"
                    "  产品编码（或 ZMATNR）、组件编码（或 IDNRK）、"
                    "MENGE、组件单价（或 ZDJ）"
                ),
            }), 422

        # 行数：从 xlsx XML 元数据读取，几乎不耗时
        total_rows = _xlsx_row_count(first_path)

        # 保存元数据
        filenames = [f.filename for f in files]
        meta = {
            "filenames":              filenames,
            "count":                  len(files),
            "has_price":              has_price,
            "price_filename":         price_fname,
        }
        (UPLOAD_DIR / f"{session_id}_meta.json").write_text(
            json.dumps(meta, ensure_ascii=False), encoding="utf-8"
        )

        preview_rows = df.fillna("").astype(str).values.tolist()
        return jsonify({
            "session_id":              session_id,
            "filenames":               filenames,
            "file_count":              len(files),
            "has_price":               has_price,
            "price_filename":          price_fname,
            "total_rows":              total_rows,
            "columns":                 list(df.columns),
            "rows":                    preview_rows,
        })

    except Exception as exc:
        for p in saved_paths:
            p.unlink(missing_ok=True)
        if price_path and price_path.exists():
            price_path.unlink(missing_ok=True)
        for ext in (".xlsx", ".xls"):
            (UPLOAD_DIR / f"{session_id}_material_trend{ext}").unlink(missing_ok=True)
        return jsonify({"error": f"上传失败：{exc}"}), 500


@app.route("/api/session/attach_price", methods=["POST"])
def attach_session_price():
    """在已有 BOM 会话上补传第二张「产品价格表」，不重新上传 BOM。"""
    sid = (request.form.get("session_id") or "").strip()
    if not _UUID_RE.match(sid):
        return jsonify({"error": "无效的 session_id"}), 400
    meta_path = UPLOAD_DIR / f"{sid}_meta.json"
    if not meta_path.exists():
        return jsonify({
            "error": "会话不存在或已失效",
            "hint":  "请重新上传 BOM 文件",
        }), 404
    bom0 = next(UPLOAD_DIR.glob(f"{sid}_0.xlsx"), None) or next(
        UPLOAD_DIR.glob(f"{sid}_0.xls"), None
    )
    if bom0 is None or not bom0.exists():
        return jsonify({
            "error": "BOM 文件缺失",
            "hint":  "请重新上传",
        }), 404

    price_file = request.files.get("price_file")
    if not price_file or not price_file.filename or not price_file.filename.strip():
        return jsonify({"error": "未收到价格表文件"}), 400
    if not _is_excel_filename(price_file.filename):
        return jsonify({
            "error":  f"价格表格式不支持：{price_file.filename}",
            "hint":   "产品价格表请上传 .xlsx 或 .xls（如 产品价格历史清单）",
        }), 400

    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return jsonify({"error": "会话元数据损坏，请重新上传"}), 500

    price_fp = UPLOAD_DIR / f"{sid}_price{_upload_save_ext(price_file.filename)}"
    try:
        price_file.save(str(price_fp))
    except Exception as exc:
        return jsonify({"error": f"保存价格表失败：{exc}"}), 500

    meta["has_price"]      = True
    meta["price_filename"] = price_file.filename
    meta_path.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")

    return jsonify({
        "success":          True,
        "has_price":        True,
        "price_filename":   price_file.filename,
    })


@app.route("/api/session/attach_material_trend", methods=["POST"])
def attach_session_material_trend():
    """材料趋势表已停用：保留旧接口，避免旧页面请求触发 404。"""
    return jsonify({
        "success": False,
        "error": "报价材料价格趋势已停用，不再参与成本模拟分析",
    }), 410


def _sidecars_for_session(session_id: str, price_fp: Path | None) -> dict:
    ppt: dict = {}
    if price_fp is not None and price_fp.exists():
        try:
            ppt = build_product_price_timeline(price_fp)
        except Exception as exc:
            print(f"产品价格时间线构建跳过: {exc}", flush=True)
    return ppt


# ---------------------------------------------------------------------------
# API – Run analysis pipeline
# ---------------------------------------------------------------------------

# Excel 后台写入状态：token -> threading.Event（set = 写完或写失败）
_EXCEL_EVENTS: dict[str, threading.Event] = {}
_EXCEL_EVENTS_LOCK = threading.Lock()


def _bg_write_excel(
    out_path: Path,
    token: str,
    sheets: list[tuple[str, pl.DataFrame]],
) -> None:
    """后台线程：把 5 张表写入 Excel 文件，完成后 set Event 通知下载接口。
    Excel 写入是纯 IO 密集型，放后台后用户无需等待即可进入 BI 看板。
    """
    try:
        with pd.ExcelWriter(str(out_path), engine="xlsxwriter") as writer:
            for name, df in sheets:
                _write_sheet(writer, df, name)
    except Exception:
        out_path.unlink(missing_ok=True)   # 写失败则删掉残文件
    finally:
        with _EXCEL_EVENTS_LOCK:
            ev = _EXCEL_EVENTS.pop(token, None)
        if ev:
            ev.set()


def _run_pipeline(raw_df: pl.DataFrame, out_path: Path):
    """执行分析管道，返回 (lines, summary_pl, detail_pl, price_history_pl, pch, bom_pd)，Excel 写入在后台进行。

    整个流程：
      1. polars 管道计算（< 1s）
      2. 填充 BI 缓存（< 0.1s）→ 调用方完成后立即返回给前端
      3. 后台线程写 Excel（10-25s）→ 用户浏览 BI 期间在后台完成
    """
    lines           = enrich(aggregate_lines(prepare(raw_df)))
    summary_pl      = sheet_component_global_rank(lines)
    detail_pl       = sheet_bom_detail(lines)
    price_history_pl = sheet_price_history(raw_df)
    product_cost_history = build_product_cost_history(raw_df)

    # 保留规范化后的原始 BOM 数据（pandas），供新预测算法使用
    try:
        bom_pd: pd.DataFrame | None = normalize_columns(raw_df).to_pandas()
    except Exception:
        bom_pd = None

    # 提前计算所有工作表（纯内存操作，极快），存为列表传给后台线程
    sheets: list[tuple[str, pl.DataFrame]] = [
        ("产品总成本排名",     sheet_product_summary(lines)),
        ("BOM明细_占比与成本", detail_pl),
        ("组件全局成本排名",   summary_pl),
        ("跨产品用量波动",     sheet_cross_product_volatility(lines)),
        ("各产品Top3成本组件", top3_per_product(lines)),
        ("价格历史明细",       price_history_pl),
    ]

    token = out_path.stem   # 文件名即 UUID token
    ev    = threading.Event()
    with _EXCEL_EVENTS_LOCK:
        _EXCEL_EVENTS[token] = ev

    threading.Thread(
        target=_bg_write_excel,
        args=(out_path, token, sheets),
        daemon=True,
    ).start()

    return lines, summary_pl, detail_pl, price_history_pl, product_cost_history, bom_pd


def _read_price_df(price_fp: "Path | None") -> "pd.DataFrame | None":
    """从价格文件读取 pandas DataFrame，规范化产品编码列，供新预测算法使用。"""
    if price_fp is None or not price_fp.exists():
        return None
    try:
        df = pd.read_excel(price_fp, sheet_name=0)
        # 规范化产品编码列名
        if COL_PRODUCT not in df.columns:
            for alias in ("ZMATNR", "MATNR", "所属产品"):
                if alias in df.columns:
                    df = df.rename(columns={alias: COL_PRODUCT})
                    break
        # 规范化 总成本 列（兜底别名）
        if "总成本" not in df.columns:
            for alias in ("产品价格", "标价", "出厂价", "DMBTR"):
                if alias in df.columns:
                    df = df.rename(columns={alias: "总成本"})
                    break
        _, df_norm = normalize_predict_dataframes(None, df)
        return df_norm
    except Exception:
        return None


def _fill_bi_cache(
    token: str,
    summary_pl: pl.DataFrame,
    detail_pl: pl.DataFrame,
    price_history_pl=None,
    product_cost_history: dict | None = None,
    product_price_timeline: dict | None = None,
    bom_pd: "pd.DataFrame | None" = None,
    price_pd: "pd.DataFrame | None" = None,
    product_categories: dict | None = None,
) -> None:
    """把分析管道已算好的 BI 表直接写入内存缓存，供 BI 接口秒返回。"""
    summary_pd = summary_pl.to_pandas()
    summary_records = summary_pd.where(pd.notna(summary_pd), None).to_dict(orient="records")

    detail_pd = detail_pl.to_pandas()
    detail_pd["组件编码"] = detail_pd["组件编码"].astype(str).str.strip()
    detail_pd[COL_PRODUCT] = detail_pd[COL_PRODUCT].astype(str).str.strip()
    if "MENGE合计" in detail_pd.columns:
        detail_pd = detail_pd.sort_values("MENGE合计", ascending=False)

    detail_index: dict[str, list] = {}
    for comp, grp in detail_pd.groupby("组件编码", sort=False):
        detail_index[str(comp)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    product_bom: dict[str, list] = {}
    for pid, grp in detail_pd.groupby(COL_PRODUCT, sort=False):
        product_bom[str(pid)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    # 价格历史明细（供日期区间价格波动查询）
    price_history_records: list = []
    if price_history_pl is not None:
        try:
            ph_pd = price_history_pl.to_pandas()
            if COL_CREATEDATE in ph_pd.columns:
                ph_pd[COL_CREATEDATE] = ph_pd[COL_CREATEDATE].map(_calendar_date_key_yyyymmdd)
            price_history_records = ph_pd.where(pd.notna(ph_pd), None).to_dict(orient="records")
        except Exception:
            pass
    pch = expand_product_cost_history_dict(
        product_cost_history or {},
        product_bom,
        product_price_timeline,
    )

    entry = {
        "summary":              summary_records,
        "detail":               detail_index,
        "product_bom":          product_bom,
        "price_history":        price_history_records,
        "product_cost_history": pch,
        "bom_pd":               bom_pd,
        "price_pd":             price_pd,
        "product_categories":   product_categories or {},
        "ts":                   time.time(),
    }
    _bi_cache_put(token, entry)


@app.route("/api/process", methods=["POST"])
def process():
    """执行 BOM 成本分析管道，将结果写为 Excel 并返回下载令牌。
    
    mode='merge'   （默认）：将所有上传文件纵向合并后生成一份报告。
    mode='per_file'         ：每个上传文件各生成一份独立报告。
    """
    data       = request.get_json(silent=True) or {}
    session_id = data.get("session_id", "")
    mode       = data.get("mode", "merge")   # "merge" | "per_file"

    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的 session_id"}), 400

    meta_path = UPLOAD_DIR / f"{session_id}_meta.json"
    if not meta_path.exists():
        return jsonify({"error": "上传文件已过期，请重新上传"}), 404

    try:
        meta      = json.loads(meta_path.read_text(encoding="utf-8"))
        filenames = meta["filenames"]
        count     = meta["count"]
        has_price = bool(meta.get("has_price"))
        price_fp  = _upload_price_path(session_id) if has_price else None
        if has_price and price_fp is None:
            return jsonify({"error": "元数据标明有产品价格表，但文件已丢失，请重新上传"}), 404

        # ── 读取所有上传文件 ──────────────────────────────────────────────
        raw_list: list[pl.DataFrame] = []
        for i in range(count):
            p = _upload_bom_path(session_id, i)
            if p is None:
                return jsonify({"error": f"上传文件 [{filenames[i]}] 已过期，请重新上传"}), 404
            raw_list.append(_read_xlsx_as_polars(p))

        ppt = _sidecars_for_session(session_id, price_fp)

        def _cleanup_inputs() -> None:
            for i in range(count):
                for ext in (".xlsx", ".xls"):
                    (UPLOAD_DIR / f"{session_id}_{i}{ext}").unlink(missing_ok=True)
            if price_fp is not None:
                price_fp.unlink(missing_ok=True)
            for ext in (".xlsx", ".xls"):
                (UPLOAD_DIR / f"{session_id}_material_trend{ext}").unlink(missing_ok=True)
            meta_path.unlink(missing_ok=True)

        # ── 准备新预测算法所需 DataFrames ────────────────────────────────
        raw_combined_pre = pl.concat(raw_list) if len(raw_list) > 1 else raw_list[0]
        try:
            bom_pd_raw: pd.DataFrame | None = normalize_columns(raw_combined_pre).to_pandas()
        except Exception:
            bom_pd_raw = None
        price_pd_raw = _read_price_df(price_fp)
        bom_pd_raw, price_pd_raw = normalize_predict_dataframes(bom_pd_raw, price_pd_raw)
        product_categories = (
            build_product_category_index(bom_pd_raw) if bom_pd_raw is not None else {}
        )

        # ── 合并模式 ──────────────────────────────────────────────────────
        if mode == "merge":
            combined = raw_combined_pre
            if has_price and price_fp is not None:
                combined = _attach_product_prices_pl(combined, price_fp)
            out_path = RESULT_DIR / f"{session_id}.xlsx"
            lines, summary_pl, detail_pl, price_history_pl, pch, bom_pd_pipe = _run_pipeline(combined, out_path)
            _fill_bi_cache(
                session_id,
                summary_pl,
                detail_pl,
                price_history_pl,
                pch,
                ppt,
                bom_pd=bom_pd_raw,
                price_pd=price_pd_raw,
                product_categories=product_categories,
            )

            _cleanup_inputs()

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            return jsonify({
                "success":        True,
                "mode":           "merge",
                "download_token": session_id,
                "filename":       f"bom_analysis_{ts}.xlsx",
                "stats": {
                    "products":   lines[COL_PRODUCT].n_unique(),
                    "components": lines[COL_COMPONENT].n_unique(),
                    "total_rows": lines.height,
                },
            })

        # ── 分别分析模式 ──────────────────────────────────────────────────
        else:
            tokens: list[dict] = []
            total_rows = 0

            for i, (raw_df_item, fname) in enumerate(zip(raw_list, filenames)):
                token    = str(uuid.uuid4())
                out_path = RESULT_DIR / f"{token}.xlsx"
                try:
                    bom_pd_item: pd.DataFrame | None = normalize_columns(raw_df_item).to_pandas()
                except Exception:
                    bom_pd_item = None
                bom_pd_item, _ = normalize_predict_dataframes(bom_pd_item, None)
                product_categories_item = (
                    build_product_category_index(bom_pd_item) if bom_pd_item is not None else {}
                )
                if has_price and price_fp is not None:
                    raw_df_item = _attach_product_prices_pl(raw_df_item, price_fp)
                lines, summary_pl, detail_pl, price_history_pl, pch, _ = _run_pipeline(raw_df_item, out_path)
                _fill_bi_cache(
                    token,
                    summary_pl,
                    detail_pl,
                    price_history_pl,
                    pch,
                    ppt,
                    bom_pd=bom_pd_item,
                    price_pd=price_pd_raw,
                    product_categories=product_categories_item,
                )
                total_rows += lines.height
                tokens.append({
                    "token":    token,
                    "filename": fname,
                    "stats": {
                        "products":   lines[COL_PRODUCT].n_unique(),
                        "components": lines[COL_COMPONENT].n_unique(),
                        "total_rows": lines.height,
                    },
                })

            _cleanup_inputs()

            return jsonify({
                "success":        True,
                "mode":           "per_file",
                "download_tokens": tokens,
                "stats": {
                    "files":      len(tokens),
                    "total_rows": total_rows,
                },
            })

    except Exception as exc:
        tb_lines = traceback.format_exc().strip().split("\n")
        detail   = tb_lines[-1] if tb_lines else None
        return jsonify({"error": f"分析失败：{exc}", "detail": detail}), 500


# ---------------------------------------------------------------------------
# BI Analysis – 内存缓存（避免每次请求重复读 Excel）
# ---------------------------------------------------------------------------

def _load_bi_cache(session_id: str, file_path: Path) -> dict:
    """
    读取 Excel 中 BI 用表，构建索引后存入缓存并返回。
    summary : list[dict]          —— 组件全局成本排名
    detail  : dict[str, list]     —— 组件编码 → 各产品明细行
    product_bom: dict[str, list]  —— 产品编码 → 该产品 BOM 明细行（成本模拟用）
    """
    # 1. 读取两张工作表
    summary_df = pd.read_excel(file_path, sheet_name="组件全局成本排名")
    detail_df  = pd.read_excel(file_path, sheet_name="BOM明细_占比与成本")

    # 2. 预处理 detail：组件编码、产品编码统一为字符串，按 MENGE合计 降序
    detail_df["组件编码"] = detail_df["组件编码"].astype(str).str.strip()
    detail_df[COL_PRODUCT] = detail_df[COL_PRODUCT].astype(str).str.strip()
    if "MENGE合计" in detail_df.columns:
        detail_df = detail_df.sort_values("MENGE合计", ascending=False)

    # 3. 将 detail 按组件编码建立字典索引，查询时 O(1)
    detail_index: dict[str, list] = {}
    for comp, grp in detail_df.groupby("组件编码", sort=False):
        detail_index[str(comp)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    product_bom: dict[str, list] = {}
    for pid, grp in detail_df.groupby(COL_PRODUCT, sort=False):
        product_bom[str(pid)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    summary_records = summary_df.where(pd.notna(summary_df), None).to_dict(orient="records")

    # 尝试读取价格历史明细（旧版 Excel 可能不含此表，容错处理）
    price_history_records: list = []
    try:
        ph_df = pd.read_excel(file_path, sheet_name="价格历史明细")
        if COL_CREATEDATE in ph_df.columns:
            ph_df[COL_CREATEDATE] = ph_df[COL_CREATEDATE].map(_calendar_date_key_yyyymmdd)
        price_history_records = ph_df.where(pd.notna(ph_df), None).to_dict(orient="records")
    except Exception:
        pass

    entry = {
        "summary":              summary_records,
        "detail":               detail_index,
        "product_bom":          product_bom,
        "price_history":        price_history_records,
        "product_cost_history": {},
        "ts":                   time.time(),
    }

    _bi_cache_put(session_id, entry)
    return entry


def _get_bi_cache(session_id: str, file_path: Path) -> dict:
    """命中则更新时间戳后返回；未命中则从磁盘加载。"""
    with _BI_CACHE_LOCK:
        if session_id in _BI_CACHE:
            c = _BI_CACHE[session_id]
            if "product_bom" not in c:
                del _BI_CACHE[session_id]
            else:
                _BI_CACHE[session_id]["ts"] = time.time()
                return _BI_CACHE[session_id]
    # 缓存未命中，在锁外读 Excel（避免长时间持锁）
    return _load_bi_cache(session_id, file_path)


# ---------------------------------------------------------------------------
# BI Analysis – Page & API
# ---------------------------------------------------------------------------

@app.route("/bi/<session_id>")
def bi_view(session_id: str):
    """BI 分析看板页面。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return "分析结果不存在或已过期，请返回首页重新上传文件并执行分析。", 404
    return render_template("bi.html", token=session_id)


@app.route("/sim/<session_id>")
def cost_sim_view(session_id: str):
    """按产品改组件单价 → 预测原材料总成本（BOM 精确重算）。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return "分析结果不存在或已过期，请返回首页重新上传文件并执行分析。", 404
    return render_template("cost_sim.html", token=session_id)


def _bi_cache_or_404(session_id: str) -> tuple[dict | None, Path | None, tuple | None]:
    if not _UUID_RE.match(session_id):
        return None, None, (jsonify({"error": "无效的令牌"}), 400)
    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return None, None, (jsonify({"error": "文件不存在或已过期，请重新分析"}), 404)
    try:
        return _get_bi_cache(session_id, file_path), file_path, None
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return None, None, (jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}), 500)


def _float_safe(x, default: float = 0.0) -> float:
    """转 float；空值/非数字/NaN/Inf 返回 default（避免 JSON 出现非法 NaN）。"""
    try:
        if x is None or (isinstance(x, str) and not str(x).strip()):
            return default
        f = float(x)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except (TypeError, ValueError):
        return default


def _optional_rounded_float(x, *, ndigits: int = 6) -> float | None:
    """用于 API 的可选数值：缺失或 NaN/Inf 返回 None，否则 round 后返回 Python float。"""
    if x is None or (isinstance(x, str) and not str(x).strip()):
        return None
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    if math.isnan(f) or math.isinf(f):
        return None
    return round(f, ndigits)


def _json_clean_scalar(v):
    """标准 JSON 不支持 NaN；将 NaN/Inf、pd.NA 等转为 None。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        return v
    if isinstance(v, int) and not isinstance(v, bool):
        return v
    if isinstance(v, float):
        return None if (math.isnan(v) or math.isinf(v)) else v
    try:
        if pd.api.types.is_scalar(v) and pd.isna(v):
            return None
    except TypeError:
        pass
    try:
        fv = float(v)
        if math.isnan(fv) or math.isinf(fv):
            return None
        return fv
    except (TypeError, ValueError):
        return v


def _json_clean_row(row: dict) -> dict:
    return {k: _json_clean_scalar(v) for k, v in row.items()}


@app.route("/api/bi/products/<session_id>")
def bi_products(session_id: str):
    """返回本分析会话中的产品编码（供成本模拟页选择）。

    产品数 ≤ ``FULL_LIST_THRESHOLD`` 时一次性返回全部 id，前端本地筛选；
    超过阈值时不返回完整列表，须带查询参数 ``q`` 做子串匹配（限量），避免超大 JSON 与 DOM。
    """
    FULL_LIST_THRESHOLD = 2000
    try:
        limit = int(request.args.get("limit", 100))
    except (TypeError, ValueError):
        limit = 100
    limit = max(1, min(limit, 500))

    cache, _, err = _bi_cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None
    pb = cache.get("product_bom") or {}
    all_ids = sorted(pb.keys(), key=lambda s: (len(str(s)), str(s)))
    total = len(all_ids)
    q = (request.args.get("q") or "").strip()

    min_q_len = 3 if total > 200_000 else 2

    if total <= FULL_LIST_THRESHOLD:
        if not q:
            items = all_ids
        else:
            items = []
            for p in all_ids:
                if q in str(p):
                    items.append(p)
                    if len(items) >= limit:
                        break
        return jsonify(
            {
                "count":           total,
                "total":           total,
                "search_required": False,
                "min_q_len":       min_q_len,
                "products":        [{"id": str(p), "label": str(p)} for p in (items if q else all_ids)],
            }
        )

    # 大会话：禁止一次性下发全部产品编码
    if not q:
        return jsonify(
            {
                "count":           total,
                "total":           total,
                "search_required": True,
                "min_q_len":       min_q_len,
                "products":        [],
                "hint":            f"共 {total} 个产品，请在输入框中输入至少 {min_q_len} 位编码片段进行搜索",
            }
        )

    if len(q) < min_q_len:
        return jsonify(
            {
                "count":           total,
                "total":           total,
                "search_required": True,
                "min_q_len":       min_q_len,
                "products":        [],
                "hint":            f"请至少输入 {min_q_len} 位再搜索（当前 {len(q)} 位）",
            }
        )

    matched: list[str] = []
    for p in all_ids:
        if q in str(p):
            matched.append(str(p))
            if len(matched) >= limit:
                break

    return jsonify(
        {
            "count":           total,
            "total":           total,
            "search_required": True,
            "min_q_len":       min_q_len,
            "q":               q,
            "match_count":     len(matched),
            "truncated":       len(matched) >= limit,
            "products":        [{"id": x, "label": x} for x in matched],
        }
    )


def _baseline_prices_from_snap(
    price_snap: dict,
    weight_kg: float | None,
) -> dict[str, Any]:
    """从价格清单快照解析基准整件价(元/件)与基准价(元/kg)，避免把元/件误乘重量。"""
    wt = _safe_float(weight_kg)
    if wt is None:
        wt = _safe_float(price_snap.get("重量"))
    norm = normalize_price_list_cost_fields(
        _safe_float(price_snap.get("材料成本")),
        _safe_float(price_snap.get("工费")),
        _safe_float(price_snap.get("总成本")),
        wt,
    )
    piece = norm.get("total_piece")
    per_kg = norm.get("total_per_kg")
    raw_total = _safe_float(price_snap.get("总成本"))
    if piece is None and raw_total is not None:
        if norm.get("costs_are_per_piece"):
            piece = raw_total
            per_kg = raw_total / wt if wt and wt > 0 else None
        elif wt and wt > 0:
            per_kg = raw_total
            piece = raw_total * wt
        else:
            piece = raw_total
            per_kg = raw_total
    list_unit = "元/件" if norm.get("costs_are_per_piece") else "元/KG"
    return {
        "基准产品价格": round(piece, 4) if piece is not None else None,
        "基准产品价格_每公斤": round(per_kg, 4) if per_kg is not None else None,
        "清单材料工费口径": list_unit,
    }


def _map_new_predict_result(
    pred: dict,
    price_snap: dict,
    *,
    is_bom_load: bool = False,
) -> dict:
    """将新 predict_product_price 的英文键结果映射为前端兼容的中文键格式。"""
    if "error" in pred:
        bl = _baseline_prices_from_snap(price_snap, None)
        return {
            "预测产品价格": None,
            "预测产品价格_每公斤": None,
            "基准产品价格": bl.get("基准产品价格") or price_snap.get("总成本"),
            "基准产品价格_每公斤": bl.get("基准产品价格_每公斤") or price_snap.get("总成本"),
            "预测方法": "—",
            "预测可信度": 0,
            "可信度等级": "低",
            "可信度说明": pred["error"],
            "敏感性分析": [],
            "敏感性网格": {"可用": False},
            "模型历史误差": {"可用": False, "说明": pred["error"]},
            "预测警告": [pred["error"]],
            "预测详情": None,
            "产品重量_kg": None,
            "价格口径": "元/KG",
        }

    conf = int(pred.get("confidence_score") or 0)
    if conf >= 90:
        level = "高"
    elif conf >= 65:
        level = "中等"
    elif conf >= 40:
        level = "较低，建议参考"
    else:
        level = "低，仅供参考"

    point_est = pred.get("point_estimate")   # 元/件
    point_kg = pred.get("point_per_kg")      # 元/kg
    weight = pred.get("detail", {}).get("base_weight")
    try:
        weight_f = float(weight) if weight is not None else None
    except (TypeError, ValueError):
        weight_f = None

    bl = _baseline_prices_from_snap(price_snap, weight_f)
    base_product = bl.get("基准产品价格")
    base_total = bl.get("基准产品价格_每公斤")

    user_adjusted = False if is_bom_load else bool(pred.get("user_adjusted_prices"))
    if is_bom_load:
        point_est = base_product
        point_kg = base_total
    diff_pred = None
    diff_pred_kg = None
    if user_adjusted:
        if point_est is not None and base_product is not None:
            try:
                diff_pred = round(float(point_est) - float(base_product), 6)
            except (TypeError, ValueError):
                pass
        if point_kg is not None and base_total is not None:
            try:
                diff_pred_kg = round(float(point_kg) - float(base_total), 6)
            except (TypeError, ValueError):
                pass

    warnings_list = list(dict.fromkeys(pred.get("warnings") or []))
    cred_parts: list[str] = []
    method = pred.get("method") or ""
    if method == "cost_structure_ratio":
        cred_parts.append("成本结构公式：材料成本随 BOM 比例传导，工费不变")
    elif method.startswith("conduction_coeff"):
        cred_parts.append("传导系数模型：由历史 BOM 与材料成本变动估计")
    detail = pred.get("detail") or {}
    n_months = detail.get("n_price_months")
    if n_months:
        cred_parts.append(f"价格历史 {n_months} 个月")
    if warnings_list:
        cred_parts.append(f"风险提示 {len(warnings_list)} 条（见上方列表）")
    return {
        "预测产品价格": point_est,
        "预测产品价格_每公斤": point_kg,
        "基准产品价格": base_product,
        "基准产品价格_每公斤": base_total,
        "清单材料工费口径": bl.get("清单材料工费口径"),
        "预测方法": pred.get("method"),
        "预测可信度": conf,
        "可信度等级": level,
        "可信度说明": "；".join(cred_parts) if cred_parts else "无明显风险",
        "用户已调价": user_adjusted,
        "差额_预测产品价": diff_pred,
        "差额_预测产品价_每公斤": diff_pred_kg,
        "基准参照说明": (
            f"价格清单基准：整件 {base_product} 元/件，{base_total} 元/kg"
            if base_product is not None and base_total is not None
            else None
        ),
        "敏感性分析": [
            map_sensitivity_item_en_to_zh(x) for x in (pred.get("sensitivity") or [])
        ],
        "敏感性网格": map_sensitivity_grid_en_to_zh(pred.get("sensitivity_grid") or {}),
        "回归分析": map_regression_analysis_en_to_zh(pred.get("regression_analysis") or {})
        if pred.get("regression_analysis")
        else None,
        "模型历史误差": pred.get("model_error") or {"可用": False},
        "预测警告": warnings_list,
        "预测详情": pred.get("detail"),
        "产品重量_kg": weight,
        "价格口径": "元/KG",
    }


@app.route("/api/bi/product_bom/<session_id>")
def bi_product_bom(session_id: str):
    """某产品的 BOM 明细行 + 当前原材料总成本；含「单价每涨 1」对总成本的边际（= MENGE合计）。"""
    product = (request.args.get("product") or "").strip()
    if not product:
        return jsonify({"error": "缺少参数 product（产品编码）"}), 400
    cache, _, err = _bi_cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None
    rows = (cache.get("product_bom") or {}).get(product)
    if not rows:
        return jsonify({"error": f"未找到产品：{product}"}), 404

    lines: list[dict] = []
    baseline = 0.0
    for r in rows:
        row = dict(r)
        m = _float_safe(row.get("MENGE合计"))
        p0 = _float_safe(row.get("组件单价"))
        line_cost = _float_safe(row.get("行原材料成本"), m * p0)
        baseline += line_cost
        row["行原材料成本"] = line_cost
        row["单价每变动1对总成本的边际"] = m
        lines.append(row)
    # 重新计算成本占比（需最终 baseline）
    if baseline > 0:
        for row in lines:
            lc = _float_safe(row.get("行原材料成本"))
            row["成本占产品%"] = round(100.0 * lc / baseline, 4)
    lines.sort(key=lambda x: _float_safe(x.get("行原材料成本")), reverse=True)

    price_snap = _product_price_snapshot_from_bom_rows(rows)
    product_price = price_snap.get("总成本")
    prices_old = {
        str(r.get("组件编码") or "").strip(): _float_safe(r.get("组件单价"))
        for r in rows
    }

    # 优先使用新预测算法（需要 bom_pd 和 price_pd）
    bom_pd = cache.get("bom_pd")
    price_pd = cache.get("price_pd")
    if bom_pd is not None and price_pd is not None:
        pred = predict_product_price(
            product,
            {},
            bom_pd,
            price_pd,
            reference_prices=prices_old,
        )
        pred = _map_new_predict_result(pred, price_snap, is_bom_load=True)
    else:
        pred = map_legacy_predict_en_to_zh(
            _predict_product_price_legacy(
                product=product,
                simulated_material=baseline,
                baseline_material=baseline,
                price_snap=price_snap,
                product_cost_history=cache.get("product_cost_history"),
                price_history=cache.get("price_history"),
                bom_rows=rows,
                prices_new=prices_old,
                prices_old=prices_old,
                product_categories=cache.get("product_categories"),
            )
        )

    lines_out = [_json_clean_row(row) for row in lines]

    return jsonify(
        {
            "product": product,
            "baseline原材料总成本": round(baseline, 6),
            "产品价格": product_price,
            "总成本": product_price,
            **{k: v for k, v in price_snap.items() if k != "总成本"},
            **pred,
            "说明": (
                "主结果「预测产品价格」由历史报价关系与成本结构（材料成本+工费）推算，不是简单 BOM 行相加。"
                "「模拟原材料总成本」为各组件数量×模拟单价之和，供对照。"
            ),
            "lines": lines_out,
        }
    )


def _simulate_product_cost(
    cache: dict,
    product: str,
    prices_in: dict,
) -> dict:
    """组件单价 → BOM 材料合计 + 预测产品总价 + 可信度。"""
    rows = (cache.get("product_bom") or {}).get(product)
    if not rows:
        raise ValueError(f"未找到产品：{product}")

    prices = {str(k).strip(): _float_safe(v) for k, v in prices_in.items()}
    prices_old: dict[str, float] = {}

    baseline = 0.0
    simulated = 0.0
    line_out: list[dict] = []
    for r in rows:
        comp = str(r.get("组件编码") or "").strip()
        m = _float_safe(r.get("MENGE合计"))
        old_p = _float_safe(r.get("组件单价"))
        prices_old[comp] = old_p
        old_line = _float_safe(r.get("行原材料成本"), m * old_p)
        baseline += old_line
        new_p = prices[comp] if comp in prices else old_p
        new_line = m * new_p
        simulated += new_line
        line_out.append(
            {
                "组件编码": comp,
                "组件名称": r.get("组件名称"),
                "MENGE合计": m,
                "原单价": old_p,
                "模拟单价": new_p,
                "原行成本": round(old_line, 6),
                "模拟行成本": round(new_line, 6),
            }
        )

    sim_round = round(simulated, 6)
    for row in line_out:
        nl = _float_safe(row.get("模拟行成本"))
        row["成本占产品%"] = (
            round(100.0 * nl / simulated, 4) if simulated > 1e-12 else 0.0
        )

    price_snap = _product_price_snapshot_from_bom_rows(rows)

    # 优先使用新预测算法
    bom_pd = cache.get("bom_pd")
    price_pd = cache.get("price_pd")
    if bom_pd is not None and price_pd is not None:
        all_prices = {c: prices.get(c, prices_old[c]) for c in prices_old}
        pred = predict_product_price(
            product,
            all_prices,
            bom_pd,
            price_pd,
            reference_prices=prices_old,
        )
        pred = _map_new_predict_result(pred, price_snap)
    else:
        pred = map_legacy_predict_en_to_zh(
            _predict_product_price_legacy(
                product=product,
                simulated_material=simulated,
                baseline_material=baseline,
                price_snap=price_snap,
                product_cost_history=cache.get("product_cost_history"),
                price_history=cache.get("price_history"),
                bom_rows=rows,
                prices_new={c: prices.get(c, prices_old[c]) for c in prices_old},
                prices_old=prices_old,
                product_categories=cache.get("product_categories"),
            )
        )

    ref_total = price_snap.get("总成本")

    payload: dict = {
        "product": product,
        "baseline原材料总成本": round(baseline, 6),
        "模拟原材料总成本": sim_round,
        "差额": round(simulated - baseline, 6),
        "总成本": ref_total,
        **{k: v for k, v in price_snap.items() if k != "总成本"},
        "差额_相对总成本": (
            round(sim_round - ref_total, 6) if ref_total is not None else None
        ),
        "lines": [_json_clean_row(x) for x in line_out],
        **pred,
    }
    return payload


@app.route("/api/bi/product_cost_simulate/<session_id>", methods=["POST"])
def bi_product_cost_simulate(session_id: str):
    """根据组件单价预测产品总价（含可信度），并返回 BOM 材料合计明细。"""
    cache, _, err = _bi_cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None
    data = request.get_json(silent=True) or {}
    product = str(data.get("product") or "").strip()
    prices_in = data.get("prices") or {}
    if not product:
        return jsonify({"error": "缺少 product"}), 400
    try:
        payload = _simulate_product_cost(cache, product, prices_in)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    return jsonify(payload)


@app.route("/api/bi/summary/<session_id>")
def bi_summary(session_id: str):
    """返回「组件全局成本排名」数据，并同步将「BOM明细」预加载进缓存。
    第一次调用会读整个 Excel（稍慢），之后所有 detail 请求均走内存。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "文件不存在或已过期，请重新分析"}), 404
    try:
        cache = _get_bi_cache(session_id, file_path)
        return jsonify({"data": cache["summary"]})
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}), 500


@app.route("/api/bi/detail/<session_id>")
def bi_detail(session_id: str):
    """返回某组件在各产品中的用量明细，走内存缓存、O(1) 查找，无磁盘 IO。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    component = request.args.get("component", "").strip()
    if not component:
        return jsonify({"error": "缺少参数 component"}), 400
    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "文件不存在或已过期，请重新分析"}), 404
    try:
        cache   = _get_bi_cache(session_id, file_path)
        records = cache["detail"].get(component, [])
        return jsonify({
            "data":      records,
            "component": component,
            "count":     len(records),
        })
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}), 500


# ---------------------------------------------------------------------------
# BI Analysis – Price Fluctuation
# ---------------------------------------------------------------------------

@app.route("/api/bi/price_fluctuation/<session_id>")
def bi_price_fluctuation(session_id: str):
    """按日期区间查询组件价格波动，按价格差值绝对值降序返回，并给出分类影响汇总。
    参数：start_date(YYYYMMDD)、end_date(YYYYMMDD)、min_change_pct(可选)
    """
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400

    start_date = (request.args.get("start_date") or "").strip().replace("-", "")
    end_date   = (request.args.get("end_date")   or "").strip().replace("-", "")
    min_pct    = abs(float(request.args.get("min_change_pct", 0) or 0))

    if not start_date or not end_date:
        return jsonify({"error": "请提供 start_date 和 end_date（格式：YYYYMMDD 或 YYYY-MM-DD）"}), 400

    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "文件不存在或已过期，请重新分析"}), 404

    try:
        cache         = _get_bi_cache(session_id, file_path)
        price_history = cache.get("price_history", [])

        if not price_history:
            return jsonify({
                "error": "无价格历史数据（可能是旧版分析结果），请重新上传并执行分析",
                "data":  [], "category_impact": [], "has_category": False,
            }), 200

        df = pd.DataFrame(price_history)

        # 统一 CREATEDATE / 生价日期 为 8 位 YYYYMMDD（与 bom_date_key 一致）
        if COL_CREATEDATE in df.columns:
            df[COL_CREATEDATE] = df[COL_CREATEDATE].map(_calendar_date_key_yyyymmdd)
            df_range = df[(df[COL_CREATEDATE] >= start_date) & (df[COL_CREATEDATE] <= end_date)].copy()
        else:
            df_range = df.copy()   # 无日期列则不过滤，返回全部

        if df_range.empty:
            return jsonify({
                "data": [], "category_impact": [], "has_category": False,
                "message": f"所选日期范围（{start_date}–{end_date}）内无数据",
                "total": 0,
            })

        comp_col   = "组件编码"
        price_col  = "组件单价"
        has_cat    = bool(COL_CATEGORY in df_range.columns and df_range[COL_CATEGORY].notna().any())
        has_name   = bool("MAKTX" in df_range.columns or "组件名称" in df_range.columns)
        name_col   = "组件名称" if "组件名称" in df_range.columns else ("MAKTX" if "MAKTX" in df_range.columns else None)

        results = []
        for comp, grp in df_range.groupby(comp_col, sort=False):
            if COL_CREATEDATE in grp.columns:
                grp = grp.sort_values(COL_CREATEDATE)
            price_vals  = grp[price_col].dropna().astype(float)
            if price_vals.empty:
                continue
            price_start = float(price_vals.iloc[0])
            price_end   = float(price_vals.iloc[-1])
            price_min   = float(price_vals.min())
            price_max   = float(price_vals.max())
            price_diff  = price_end - price_start
            price_pct   = round((price_diff / price_start * 100), 2) if price_start != 0 else 0.0

            row = {
                "组件编码":   str(comp),
                "期初价格":   round(price_start, 4),
                "期末价格":   round(price_end, 4),
                "期间最低价": round(price_min, 4),
                "期间最高价": round(price_max, 4),
                "价格差值":   round(price_diff, 4),
                "价格变动%":  price_pct,
                "记录数":     int(len(grp)),
                "涉及产品数": int(grp["产品编码"].nunique()) if "产品编码" in grp.columns else 0,
            }
            if name_col:
                row["组件名称"] = str(grp[name_col].dropna().iloc[0]) if grp[name_col].notna().any() else ""
            if has_cat:
                row[COL_CATEGORY] = str(grp[COL_CATEGORY].dropna().iloc[0]) if grp[COL_CATEGORY].notna().any() else ""
            results.append(row)

        # 按变动百分比绝对值过滤
        if min_pct > 0:
            results = [r for r in results if abs(r["价格变动%"]) >= min_pct]

        # 默认按价格差值降序排列（最大涨幅在前）
        results.sort(key=lambda r: r["价格差值"], reverse=True)

        # 分类影响汇总
        category_impact = []
        if has_cat and results:
            cat_df = pd.DataFrame(results)
            cat_df = cat_df[cat_df[COL_CATEGORY].notna() & (cat_df[COL_CATEGORY] != "")]
            if not cat_df.empty:
                for cat, cgrp in cat_df.groupby(COL_CATEGORY, sort=False):
                    pcts = cgrp["价格变动%"]
                    abs_max = float(pcts.abs().max())
                    level   = "高" if abs_max >= 10 else ("中" if abs_max >= 3 else "低")
                    category_impact.append({
                        "分类":        str(cat),
                        "平均变动%":   round(float(pcts.mean()), 2),
                        "最大涨幅%":   round(float(pcts.max()),  2),
                        "最大跌幅%":   round(float(pcts.min()),  2),
                        "最大绝对变动%": round(abs_max, 2),
                        "涉及组件数":  int(len(cgrp)),
                        "上涨组件数":  int((pcts > 0).sum()),
                        "下跌组件数":  int((pcts < 0).sum()),
                        "影响等级":    level,
                    })
                category_impact.sort(key=lambda r: r["最大绝对变动%"], reverse=True)

        return jsonify({
            "data":           results,
            "category_impact": category_impact,
            "date_range":     {"start": start_date, "end": end_date},
            "total":          len(results),
            "has_category":   has_cat,
            "has_name":       bool(name_col),
        })

    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"价格波动分析失败：{exc}", "detail": tb[-1]}), 500


def _parse_excel_dates(series: pd.Series) -> pd.Series:
    """把 Excel/SAP 常见日期格式统一成 pandas datetime。"""
    text = (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .str[:8]
    )
    return pd.to_datetime(text, format="%Y%m%d", errors="coerce")


def _first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def _strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def _normalize_search_url(url: str) -> str:
    """DuckDuckGo 结果常包一层跳转链接，这里还原真实地址。"""
    url = html.unescape(url or "")
    if "uddg=" in url:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        if qs.get("uddg"):
            return qs["uddg"][0]
    if url.startswith("//"):
        return "https:" + url
    return url


def _fetch_page_summary(url: str) -> dict:
    """抓取搜索结果页的标题和 meta 描述，补强搜索摘要。"""
    if not url or not url.startswith(("http://", "https://")):
        return {"title": "", "description": ""}
    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                )
            },
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            page = resp.read(180000).decode("utf-8", errors="ignore")
    except Exception:
        return {"title": "", "description": ""}

    title = ""
    m_title = re.search(r"<title[^>]*>(.*?)</title>", page, re.S | re.I)
    if m_title:
        title = _strip_html(m_title.group(1))

    description = ""
    m_desc = re.search(
        r'<meta[^>]+(?:name|property)=["\'](?:description|og:description)["\'][^>]+content=["\']([^"\']+)["\']',
        page,
        re.S | re.I,
    )
    if not m_desc:
        m_desc = re.search(
            r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+(?:name|property)=["\'](?:description|og:description)["\']',
            page,
            re.S | re.I,
        )
    if m_desc:
        description = _strip_html(m_desc.group(1))

    return {"title": title[:160], "description": description[:260]}


def _baidu_ai_search(query: str, *, max_results: int = 5) -> list[dict]:
    """调用百度千帆 AI 搜索 API，返回网页标题、摘要、链接和日期。"""
    api_key = os.environ.get("BAIDU_SEARCH_API_KEY", "").strip()
    if not api_key:
        return []

    short_query = query.strip()
    # 文档限制 content 72 字符以内；中文按更保守的长度截断。
    if len(short_query) > 60:
        short_query = short_query[:60]

    payload = {
        "messages": [{"role": "user", "content": short_query}],
        "search_source": "baidu_search_v2",
        "resource_type_filter": [{"type": "web", "top_k": max(1, min(max_results, 10))}],
        "search_filter": {
            "range": {
                "page_time": {
                    "gte": "now-3M/d",
                    "lte": "now/d",
                }
            }
        },
        "sort": {"priority": "auto"},
    }
    headers = {
        "Content-Type": "application/json",
        # 文档示例使用 X-Appbuilder-Authorization，用户提供的 bce-v3 key 属于该类 API Key。
        "X-Appbuilder-Authorization": f"Bearer {api_key}",
    }
    req = urllib.request.Request(
        "https://qianfan.baidubce.com/v2/ai_search/web_search",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=25) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        return [{
            "title": "百度千帆搜索失败",
            "snippet": f"HTTP {exc.code}: {detail[:240]}",
            "url": "",
            "source": "baidu_qianfan",
        }]
    except Exception as exc:
        return [{
            "title": "百度千帆搜索失败",
            "snippet": str(exc),
            "url": "",
            "source": "baidu_qianfan",
        }]

    if data.get("code") or data.get("message"):
        return [{
            "title": "百度千帆搜索失败",
            "snippet": f"{data.get('code', '')} {data.get('message', '')}".strip(),
            "url": "",
            "source": "baidu_qianfan",
        }]

    results: list[dict] = []
    for item in data.get("references", [])[:max_results]:
        title = str(item.get("title") or item.get("web_anchor") or "").strip()
        snippet = str(item.get("snippet") or item.get("content") or "").strip()
        url = str(item.get("url") or "").strip()
        date = str(item.get("date") or "").strip()
        if title or snippet or url:
            results.append({
                "title": title,
                "snippet": snippet,
                "url": url,
                "date": date,
                "source": "baidu_qianfan",
            })
    return results


def _web_search_material_price(material_name: str, *, max_results: int = 3) -> list[dict]:
    """用公开搜索结果摘要收集物料近期价格/行情证据。"""
    query = f"{material_name} 2026 价格 走势 行情 报价 原材料"
    baidu_api_results = _baidu_ai_search(query, max_results=max_results)
    if baidu_api_results and not all(item.get("title") == "百度千帆搜索失败" for item in baidu_api_results):
        return baidu_api_results

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        )
    }

    url = "https://duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query})
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            page = resp.read().decode("utf-8", errors="ignore")
    except Exception as exc:
        page = ""

    blocks = re.split(r'<div[^>]+class="[^"]*result[^"]*"[^>]*>', page)
    results: list[dict] = []
    for block in blocks:
        title_match = re.search(r'class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', block, re.S)
        if not title_match:
            continue
        snippet_match = re.search(r'class="result__snippet"[^>]*>(.*?)</(?:a|div)>', block, re.S)
        title = _strip_html(title_match.group(2))
        link = _normalize_search_url(title_match.group(1))
        snippet = _strip_html(snippet_match.group(1)) if snippet_match else ""
        page_summary = _fetch_page_summary(link) if len(results) < 2 else {"title": "", "description": ""}
        if not snippet:
            snippet = page_summary.get("description", "")
        if page_summary.get("title") and page_summary["title"] not in title:
            title = f"{title} | {page_summary['title']}"
        if title:
            results.append({"title": title, "snippet": snippet, "url": link})
        if len(results) >= max_results:
            break
    if results:
        return results[:max_results]

    # 中文行业行情词优先回退到百度，通常比英文搜索引擎更容易命中中文报价/行情页。
    baidu_url = "https://www.baidu.com/s?" + urllib.parse.urlencode({"wd": query})
    try:
        with urllib.request.urlopen(urllib.request.Request(baidu_url, headers=headers), timeout=15) as resp:
            page = resp.read().decode("utf-8", errors="ignore")
    except Exception:
        page = ""

    for block in re.findall(r'<div[^>]+(?:class="[^"]*result[^"]*"|tpl="[^"]+")[^>]*>(.*?)</div>\s*</div>', page, re.S | re.I):
        title_match = re.search(r"<h3[^>]*>.*?<a[^>]+href=\"([^\"]+)\"[^>]*>(.*?)</a>.*?</h3>", block, re.S | re.I)
        if not title_match:
            continue
        snippet_match = re.search(r'<span[^>]+class="[^"]*(?:content-right|c-color-text)[^"]*"[^>]*>(.*?)</span>', block, re.S | re.I)
        if not snippet_match:
            snippet_match = re.search(r'<div[^>]+class="[^"]*(?:c-abstract|c-span-last)[^"]*"[^>]*>(.*?)</div>', block, re.S | re.I)
        link = _normalize_search_url(title_match.group(1))
        title = _strip_html(title_match.group(2))
        snippet = _strip_html(snippet_match.group(1)) if snippet_match else ""
        if title:
            results.append({"title": title, "snippet": snippet, "url": link})
        if len(results) >= max_results:
            break
    if results:
        return results

    # 百度也没结果时，再自动回退到 Bing。
    bing_url = "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query})
    try:
        with urllib.request.urlopen(urllib.request.Request(bing_url, headers=headers), timeout=15) as resp:
            page = resp.read().decode("utf-8", errors="ignore")
    except Exception as exc:
        return [{"title": "搜索失败", "snippet": str(exc), "url": ""}]

    for block in re.findall(r'<li[^>]+class="b_algo"[^>]*>(.*?)</li>', page, re.S | re.I):
        title_match = re.search(r"<h2[^>]*>.*?<a[^>]+href=\"([^\"]+)\"[^>]*>(.*?)</a>.*?</h2>", block, re.S | re.I)
        if not title_match:
            continue
        snippet_match = re.search(r'<p[^>]*>(.*?)</p>', block, re.S | re.I)
        link = _normalize_search_url(title_match.group(1))
        title = _strip_html(title_match.group(2))
        snippet = _strip_html(snippet_match.group(1)) if snippet_match else ""
        page_summary = _fetch_page_summary(link) if len(results) < 2 else {"title": "", "description": ""}
        if not snippet:
            snippet = page_summary.get("description", "")
        if page_summary.get("title") and page_summary["title"] not in title:
            title = f"{title} | {page_summary['title']}"
        if title:
            results.append({"title": title, "snippet": snippet, "url": link})
        if len(results) >= max_results:
            break
    if results:
        return results

    # 搜索引擎不可用时，给出行业行情入口作为方向性证据，避免最终报告完全失去依据。
    lower = material_name.lower()
    fallback = [
        {
            "title": f"{material_name} 1688 报价/批发价格入口",
            "snippet": "用于观察同类物料现货报价区间、供应商报价密度和价格离散度，适合作为采购比价参考。",
            "url": "https://www.1688.com/",
        }
    ]
    if any(k in lower for k in ["ppr", "pp-r", "pp", "管件", "管材", "给水管"]):
        fallback.extend([
            {
                "title": "PP / PPR 管材上游原料行情入口",
                "snippet": "PPR 管件和管材主要受 PP 原料、丙烯、原油、装置检修、下游开工和房地产/基建需求影响。",
                "url": "https://plas.oilchem.net/",
            },
            {
                "title": "生意社 PP 商品指数与价格行情入口",
                "snippet": "可用于跟踪 PP 原料价格趋势、上游成本支撑和短期市场情绪。",
                "url": "https://www.100ppi.com/",
            },
        ])
    if any(k in lower for k in ["pe", "聚乙烯", "包装", "膜", "opp"]):
        fallback.extend([
            {
                "title": "PE / 包装膜上游原料行情入口",
                "snippet": "PE、OPP、包装膜价格通常受原油、乙烯/丙烯、聚烯烃供应、薄膜需求和库存影响。",
                "url": "https://plas.oilchem.net/",
            },
            {
                "title": "生意社聚乙烯/聚丙烯行情入口",
                "snippet": "可用于观察聚烯烃原料价格趋势和成本传导压力。",
                "url": "https://www.100ppi.com/",
            },
        ])
    if any(k in lower for k in ["钛白", "钛白粉"]):
        fallback.append({
            "title": "钛白粉行情与上游钛矿/硫酸成本入口",
            "snippet": "钛白粉价格通常受钛矿、硫酸、出口订单、行业开工和库存变化影响。",
            "url": "https://www.100ppi.com/",
        })
    if any(k in lower for k in ["碳酸钙", "钙粉"]):
        fallback.append({
            "title": "碳酸钙/填充母料行情与矿石能源成本入口",
            "snippet": "碳酸钙价格通常受矿石、能源、运输、环保限产和下游塑料填充需求影响。",
            "url": "https://www.100ppi.com/",
        })
    return fallback[:max_results]


def _deepseek_model() -> str:
    return os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-pro").strip() or "deepseek-v4-pro"


def _deepseek_headers() -> dict:
    api_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "未配置 DEEPSEEK_API_KEY。请先在 web/.env 中设置 DEEPSEEK_API_KEY。"
        )
    return {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }


def _deepseek_chat_completion(payload: dict, *, timeout: int = 60) -> dict:
    req = urllib.request.Request(
        "https://api.deepseek.com/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=_deepseek_headers(),
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"DeepSeek API 调用失败（HTTP {exc.code}）：{detail[:300]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"无法连接 DeepSeek API：{exc}") from exc


_MARKET_SEARCH_TOOL = {
    "type": "function",
    "function": {
        "name": "search_market_price",
        "description": "搜索某个物料或上游原材料最近三个月的市场价格、行情、供需、检修、成本变化等公开网页信息。",
        "parameters": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "搜索关键词，例如：PPR 管件 最近三个月 PP 原料 价格 走势 检修 开工率",
                },
                "material_name": {
                    "type": "string",
                    "description": "该搜索对应的物料名称或上游原料名称。",
                },
            },
            "required": ["query", "material_name"],
        },
    },
}


def _run_market_search_tool(query: str, material_name: str) -> dict:
    results = _web_search_material_price(query, max_results=4)
    evidence_strength = "较弱"
    if any(item.get("snippet") for item in results):
        evidence_strength = "中等"
    if any(re.search(r"价格|行情|报价|走势|涨|跌|检修|开工|成本|原油|PP|PE|PPR|钛白粉|碳酸钙", item.get("title", "") + item.get("snippet", ""), re.I) for item in results):
        evidence_strength = "可用于方向判断"
    return {
        "query": query,
        "material_name": material_name,
        "evidence_strength": evidence_strength,
        "results": results,
    }


def _prepare_market_search_messages(
    materials: list[dict],
    start_date: str,
    end_date: str,
) -> tuple[list[dict], list[dict]]:
    """让 DeepSeek 先决定搜索什么，再执行工具并返回可用于最终回答的消息。"""
    first_payload = _deepseek_purchase_advice_payload(
        materials,
        start_date,
        end_date,
        stream=False,
        use_tools=True,
    )
    messages = list(first_payload["messages"])
    first_resp = _deepseek_chat_completion(first_payload, timeout=90)
    message = first_resp["choices"][0]["message"]
    tool_calls = message.get("tool_calls") or []

    if not tool_calls:
        messages.append({
            "role": "assistant",
            "content": message.get("content") or "未生成搜索工具调用。",
        })
        messages.append({
            "role": "user",
            "content": "你没有调用搜索工具。请明确说明实时行情证据不足，并基于产业链通用逻辑给出保守采购建议。",
        })
        return messages, []

    messages.append({
        "role": "assistant",
        "content": message.get("content"),
        "tool_calls": tool_calls,
    })

    evidence: list[dict] = []
    for tool_call in tool_calls[:8]:
        fn = tool_call.get("function", {})
        if fn.get("name") != "search_market_price":
            continue
        try:
            args = json.loads(fn.get("arguments") or "{}")
        except json.JSONDecodeError:
            args = {}
        query = str(args.get("query") or "").strip()
        material_name = str(args.get("material_name") or query or "未知物料").strip()
        if not query:
            query = f"{material_name} 最近三个月 价格 行情 走势"

        result = _run_market_search_tool(query, material_name)
        evidence.append(result)
        messages.append({
            "role": "tool",
            "tool_call_id": tool_call["id"],
            "content": json.dumps(result, ensure_ascii=False),
        })

    messages.append({
        "role": "user",
        "content": (
            "请基于以上工具搜索结果输出最终 Markdown 采购行情分析。"
            "必须先总结市场因素和短期展望，再给推荐购入、涨幅需注意和采购操作建议。"
            "如果搜索结果包含价格、报价、行情、走势、上游原料等关键词，就可以作为方向性判断依据。"
            "不要反复写“实时行情证据不足”；只在完全没有相关结果的物料上标注证据不足。"
            "如果没有精确涨跌幅，就写“未获得精确涨跌幅，按方向性行情判断”。"
        ),
    })
    return messages, evidence


def _call_deepseek_for_purchase_advice(materials: list[dict], start_date: str, end_date: str) -> str:
    messages, _ = _prepare_market_search_messages(materials, start_date, end_date)
    payload = _deepseek_purchase_advice_payload(
        materials,
        start_date,
        end_date,
        stream=False,
        use_tools=False,
        thinking_enabled=True,
        messages=messages,
    )
    data = _deepseek_chat_completion(payload, timeout=90)
    try:
        return data["choices"][0]["message"]["content"].strip()
    except Exception as exc:
        raise RuntimeError(f"DeepSeek 返回格式异常：{data}") from exc


def _build_purchase_advice_materials(session_id: str, file_path: Path) -> dict:
    cache = _get_bi_cache(session_id, file_path)
    price_history = cache.get("price_history", [])
    if not price_history:
        raise ValueError("无物料明细数据，请重新上传并执行分析")

    df = pd.DataFrame(price_history)
    name_col = "组件名称" if "组件名称" in df.columns else ("MAKTX" if "MAKTX" in df.columns else None)
    required = [COL_COMPONENT]
    missing = [c for c in required if c not in df.columns]
    if missing or not name_col:
        raise ValueError(
            "缺少 AI 联网分析所需列：需要 组件编码、MAKTX（或组件名称）；"
            f"当前缺少：{missing + ([] if name_col else ['MAKTX'])}"
        )

    df[name_col] = df[name_col].fillna("").astype(str).str.strip()
    df[COL_COMPONENT] = df[COL_COMPONENT].fillna("").astype(str).str.strip()
    df = df[(df[name_col] != "") & (df[COL_COMPONENT] != "")]
    if df.empty:
        raise ValueError("MAKTX / 组件编码有效数据为空，无法生成 AI 建议")

    rows: list[dict] = []
    for (comp, name), grp in df.groupby([COL_COMPONENT, name_col], sort=False):
        row = {
            "组件编码": str(comp),
            "物料名称": str(name),
            "记录数": int(len(grp)),
            "涉及产品数": int(grp[COL_PRODUCT].nunique()) if COL_PRODUCT in grp.columns else 0,
        }
        rows.append(row)

    if not rows:
        raise ValueError("未找到可分析的 MAKTX 物料记录")

    # 先选使用面更广、记录更多的物料交给 DeepSeek，后端不再做联网搜索。
    rows.sort(key=lambda r: (r["涉及产品数"], r["记录数"]), reverse=True)
    selected = rows[:12]

    today = datetime.now()
    start_dt = today - pd.DateOffset(months=3)
    return {
        "materials": selected,
        "date_range": {"start": start_dt.strftime("%Y-%m-%d"), "end": today.strftime("%Y-%m-%d")},
        "total": len(rows),
    }


def _deepseek_purchase_advice_payload(
    materials: list[dict],
    start_date: str,
    end_date: str,
    *,
    stream: bool,
    use_tools: bool = False,
    thinking_enabled: bool = False,
    messages: list[dict] | None = None,
) -> dict:
    prompt = (
        "你是制造业采购和原材料成本分析顾问。Excel 只提供待分析物料清单，不提供价格结论。"
        "请你基于最近三个月的公开市场行情、网上价格趋势、原材料供需信息，"
        "判断这些物料是否适合购入或需要注意涨价风险。"
        "输出风格要像专业采购行情简报：先讲市场逻辑，再给动作建议，不能只罗列物料。\n"
        "要求：\n"
        "1. 用 Markdown 输出。\n"
        "2. 必须按以下结构输出：\n"
        "   - `## 近期影响市场价格的核心因素`\n"
        "   - `## 短期市场展望`\n"
        "   - `## 推荐购入`\n"
        "   - `## 涨幅较大需注意`\n"
        "   - `## 采购操作建议`\n"
        "3. `近期影响市场价格的核心因素` 至少写 3 条，围绕供应、成本、需求、替代品、进口/出口、季节性、政策或检修等因素展开。\n"
        "4. `短期市场展望` 要给出明确判断，例如高位震荡、偏强运行、弱势回调、低位企稳，并说明为什么。\n"
        "5. `推荐购入` 和 `涨幅较大需注意` 每条包含物料名称、组件编码、判断方向、依据、建议动作。\n"
        "6. 每条建议必须有“有理有据”的原因：例如上游原料、供需变化、成本支撑、需求淡旺季、库存或行业开工率等。\n"
        "7. 如果工具搜索结果没有精确涨跌幅，不要反复写“实时行情证据不足”；应改写为“未获得精确涨跌幅，以下为方向性判断”。\n"
        "8. 只在某个物料完全没有价格/行情/报价/走势相关搜索结果时，才对该物料标注“证据不足”。\n"
        "9. 可以把搜索结果中的价格页、报价页、行情页、原料价格走势页作为方向性证据，但不要编造具体涨跌幅、具体开工率或具体价格。\n"
        "10. 不要按 Excel 记录数排序，也不要假装 Excel 里有市场涨跌幅。\n"
        "11. 优先结合物料关键词判断，例如 PE、PP-R、管材、管件、包装膜、钛白粉、碳酸钙等相关大宗原料趋势。\n"
        "12. 语言要像采购经理能直接看的行情简报：简短、有判断、有依据、有动作。\n\n"
        "工具使用要求：\n"
        "- 在最终分析前，优先调用 `search_market_price` 工具搜索 3-8 个关键行情问题。\n"
        "- 搜索词由你根据物料清单自行设计，既要搜具体物料，也要搜上游原料，如 PP、PE、PPR、钛白粉、碳酸钙等。\n"
        "- 最终结论要引用工具结果里的标题/摘要关键词作为依据。\n\n"
        "参考表达风格：\n"
        "- 供应收缩：若上游装置检修、开工率下降或供应偏紧，价格下方支撑增强。\n"
        "- 高成本支撑：若原油、丙烯、丙烷、甲醇、钛矿等上游成本处于高位，下游物料存在被动跟涨压力。\n"
        "- 需求端压力：若处于传统淡季或终端开工不足，价格上行空间会受压。\n"
        "- 操作建议：不要只写“推荐购买”，要写按需补库、锁价、分批采购、观望、供应商比价或替代料评估。\n\n"
        f"日期范围：{start_date} 至 {end_date}\n"
        f"待分析物料清单 JSON：{json.dumps(materials, ensure_ascii=False)}"
    )
    payload = {
        "model": _deepseek_model(),
        "thinking": {"type": "enabled" if thinking_enabled else "disabled"},
        "temperature": 0.2,
        "stream": stream,
        "messages": messages or [
            {"role": "system", "content": "你只基于用户提供的数据给采购建议，避免空泛表达。"},
            {"role": "user", "content": prompt},
        ],
    }
    if use_tools:
        payload["tools"] = [_MARKET_SEARCH_TOOL]
        payload["tool_choice"] = "auto"
    return payload


def _stream_deepseek_for_purchase_advice(materials: list[dict], start_date: str, end_date: str):
    messages, _ = _prepare_market_search_messages(materials, start_date, end_date)

    req = urllib.request.Request(
        "https://api.deepseek.com/chat/completions",
        data=json.dumps(
            _deepseek_purchase_advice_payload(
                materials,
                start_date,
                end_date,
                stream=True,
                use_tools=False,
                thinking_enabled=True,
                messages=messages,
            ),
            ensure_ascii=False,
        ).encode("utf-8"),
        headers=_deepseek_headers(),
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            for raw_line in resp:
                line = raw_line.decode("utf-8", errors="ignore").strip()
                if not line.startswith("data:"):
                    continue
                data = line[5:].strip()
                if data == "[DONE]":
                    break
                try:
                    chunk = json.loads(data)
                    delta = chunk["choices"][0].get("delta", {}).get("content", "")
                except Exception:
                    delta = ""
                if delta:
                    yield delta
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"DeepSeek API 调用失败（HTTP {exc.code}）：{detail[:300]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"无法连接 DeepSeek API：{exc}") from exc


@app.route("/api/bi/ai_purchase_advice/<session_id>")
def bi_ai_purchase_advice(session_id: str):
    """读取 MAKTX 物料名称，计算最近三个月价格波动，并调用 DeepSeek 生成采购建议。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400

    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "文件不存在或已过期，请重新分析"}), 404

    try:
        info = _build_purchase_advice_materials(session_id, file_path)
        materials_for_ai = info["materials"]
        start_s = info["date_range"]["start"]
        end_s = info["date_range"]["end"]
        advice = _call_deepseek_for_purchase_advice(materials_for_ai, start_s, end_s)

        return jsonify({
            "advice": advice,
            "materials": materials_for_ai,
            "date_range": info["date_range"],
            "total": info["total"],
        })

    except ValueError as exc:
        return jsonify({"error": str(exc)}), 422
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"AI 采购建议生成失败：{exc}", "detail": tb[-1]}), 500


@app.route("/api/bi/ai_purchase_advice_stream/<session_id>")
def bi_ai_purchase_advice_stream(session_id: str):
    """流式返回 Markdown 格式 AI 采购建议。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的令牌"}), 400

    file_path = RESULT_DIR / f"{session_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "文件不存在或已过期，请重新分析"}), 404

    def sse(event: str, data) -> str:
        return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

    @stream_with_context
    def generate():
        try:
            info = _build_purchase_advice_materials(session_id, file_path)
            yield sse("meta", info)
            for delta in _stream_deepseek_for_purchase_advice(
                info["materials"],
                info["date_range"]["start"],
                info["date_range"]["end"],
            ):
                yield sse("delta", {"text": delta})
            yield sse("done", {"ok": True})
        except Exception as exc:
            yield sse("error", {"error": f"AI 采购建议生成失败：{exc}"})

    return Response(generate(), mimetype="text/event-stream")


# ---------------------------------------------------------------------------
# API – Download result
# ---------------------------------------------------------------------------

@app.route("/api/download/<session_id>")
def download(session_id: str):
    """凭 session_id（下载令牌）下载分析结果 Excel。
    若 Excel 仍在后台写入，最多等待 120 秒；写完即返回文件。
    """
    if not _UUID_RE.match(session_id):
        return jsonify({"error": "无效的下载令牌"}), 400

    file_path = RESULT_DIR / f"{session_id}.xlsx"

    # 如果后台写入尚未完成，等 Event 信号（最多 120s）
    with _EXCEL_EVENTS_LOCK:
        ev = _EXCEL_EVENTS.get(session_id)
    if ev is not None:
        ev.wait(timeout=120)

    if not file_path.exists():
        return jsonify({"error": "文件生成失败或已过期，请重新分析"}), 404

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        file_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"bom_analysis_{ts}.xlsx",
    )


# ---------------------------------------------------------------------------
# API – Excel ready status（供前端判断导出按钮是否可用）
# ---------------------------------------------------------------------------

@app.route("/api/excel_ready/<session_id>")
def excel_ready(session_id: str):
    """返回 Excel 是否已写好。前端可轮询此接口来更新「导出」按钮状态。"""
    if not _UUID_RE.match(session_id):
        return jsonify({"ready": False}), 400
    with _EXCEL_EVENTS_LOCK:
        ev = _EXCEL_EVENTS.get(session_id)
    if ev is None:
        # 不在写入队列里：要么已写完（文件存在），要么根本没有
        ready = (RESULT_DIR / f"{session_id}.xlsx").exists()
    else:
        ready = ev.is_set()
    return jsonify({"ready": ready})


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def _open_browser_when_ready(url: str = "http://127.0.0.1:5000/", delay_sec: float = 1.0) -> None:
    """服务启动后自动打开默认浏览器。设环境变量 BOM_WEB_NO_BROWSER=1 可关闭。"""
    flag = os.environ.get("BOM_WEB_NO_BROWSER", "").strip().lower()
    if flag in ("1", "true", "yes", "on"):
        return
    time.sleep(delay_sec)
    webbrowser.open(url)


def _should_auto_open_browser(*, debug: bool, use_reloader: bool) -> bool:
    """是否在「当前进程」里调度打开浏览器（避免 debug 重载时父、子进程各开一次）。"""
    if os.environ.get("BOM_WEB_NO_BROWSER", "").strip().lower() in ("1", "true", "yes", "on"):
        return False
    # 在 app.run() 之前 app.debug 仍为默认 False，不能用 app.debug 判断。
    if debug and use_reloader:
        return os.environ.get("WERKZEUG_RUN_MAIN") == "true"
    return True


if __name__ == "__main__":
    _DEBUG = True
    _USE_RELOADER = _DEBUG

    print()
    print("=" * 52)
    print("   BOM 成本分析 Web 应用")
    print("   访问地址 → http://127.0.0.1:5000")
    print("=" * 52)
    print()
    if _should_auto_open_browser(debug=_DEBUG, use_reloader=_USE_RELOADER):
        threading.Thread(
            target=_open_browser_when_ready,
            kwargs={"url": "http://127.0.0.1:5000/", "delay_sec": 1.0},
            daemon=True,
        ).start()
    app.run(debug=_DEBUG, use_reloader=_USE_RELOADER, port=5000, host="0.0.0.0")
