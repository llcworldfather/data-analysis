# -*- coding: utf-8 -*-
"""
BOM 成本分析（Polars）— 网页版随包副本
--------------------------------------
本文件与 `web/app.py` 同目录，供 Flask 直接 import，不依赖项目其他目录。

可选命令行（在 web 目录下）：
  python process_excel.py          # 读取本目录下 data/ 内全部 xlsx → output/
  python process_excel.py --per-file

依赖见同目录 requirements.txt
"""
from __future__ import annotations

import argparse
import re
import sys
from typing import Any
from datetime import datetime
from pathlib import Path

import pandas as pd
import polars as pl

ROOT       = Path(__file__).resolve().parent
DATA_DIR   = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

COL_PRODUCT    = "产品编码"
COL_COMPONENT  = "组件编码"
COL_QTY        = "MENGE"
COL_UNIT       = "MEINS"
COL_UNIT_PRICE = "组件单价"
COL_MAKTX      = "MAKTX"
COL_CREATEDATE = "CREATEDATE"
COL_CATEGORY   = "分类"

REQUIRED_COLS = [COL_PRODUCT, COL_COMPONENT, COL_QTY, COL_UNIT_PRICE]

# SAP / 内控清单列名 → 脚本标准列名（自动映射，多种导出格式均可直接使用）
_COL_ALIASES: dict[str, str] = {
    # SAP 导出
    "ZMATNR": COL_PRODUCT,
    "IDNRK":  COL_COMPONENT,
    "ZDJ":    COL_UNIT_PRICE,
    # 报价 BOM 历史清单（模板：报价BOM历史清单.xls）
    "所属产品": COL_PRODUCT,
    "材料编码": COL_COMPONENT,
    "组件数量": COL_QTY,
    "物料用量": COL_QTY,
    "材料单价": COL_UNIT_PRICE,
    "物料单价": COL_UNIT_PRICE,
    "材料型号": COL_MAKTX,
    "基本单位": COL_UNIT,
    "生价日期": COL_CREATEDATE,
    # 通用日期列
    "生产日期": COL_CREATEDATE,
    "创建日期": COL_CREATEDATE,
    "日期":    COL_CREATEDATE,
}


def normalize_columns(df: pl.DataFrame) -> pl.DataFrame:
    """将 SAP 导出列名（ZMATNR/IDNRK/ZDJ）自动重命名为脚本标准列名。
    若标准列名已存在则不覆盖，兼容两种命名格式。"""
    renames = {
        old: new
        for old, new in _COL_ALIASES.items()
        if old in df.columns and new not in df.columns
    }
    return df.rename(renames) if renames else df


def norm_material_code_scalar(code: Any) -> str:
    """单值物料编码规范化（与 norm_material_code 列处理规则一致）。"""
    if code is None:
        return ""
    s = str(code).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    if s in ("nan", "None", "<NA>"):
        return ""
    return s


def norm_material_code(series: pd.Series) -> pd.Series:
    """统一物料编码为字符串：去空白、去掉 Excel 浮点尾缀 .0。"""
    return series.map(norm_material_code_scalar)


# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------

def _excel_paths_in_data() -> list[Path]:
    """收集 data/ 下所有 Excel 工作簿（扩展名不区分大小写：.xlsx / .XLSX）。"""
    if not DATA_DIR.is_dir():
        return []
    seen: set[str] = set()
    paths: list[Path] = []
    for p in DATA_DIR.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() != ".xlsx":
            continue
        key = p.name.lower()
        if key not in seen:
            seen.add(key)
            paths.append(p)
    return sorted(paths, key=lambda x: x.name.lower())


def _print_no_excel_error() -> None:
    listing = ""
    if DATA_DIR.is_dir():
        names = sorted(p.name for p in DATA_DIR.iterdir())
        listing = f"\n当前 data 目录内文件：{names if names else '（空）'}"
    else:
        listing = f"\n（不存在 data 目录，期望路径：{DATA_DIR.resolve()}）"
    print(
        f"错误：在 {DATA_DIR.resolve()} 中未找到 Excel（扩展名须为 .xlsx，大小写不限）。"
        f"{listing}\n"
        f"请将 .xlsx 放在本目录下的 data 文件夹中。"
    )


def _read_one_excel(path: Path) -> pl.DataFrame:
    print(f"  读取：{path.name}")
    return (
        pl.from_pandas(pd.read_excel(path, sheet_name=0))
        .with_columns(pl.lit(path.name).alias("_源文件"))
    )


def _validate_required_columns(raw: pl.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        print(f"错误：缺少必要列 {missing}\n当前列名：{raw.columns}")
        sys.exit(1)


def load_all_excels() -> pl.DataFrame:
    paths = _excel_paths_in_data()
    if not paths:
        _print_no_excel_error()
        sys.exit(1)
    frames: list[pl.DataFrame] = []
    for p in paths:
        frames.append(_read_one_excel(p))
    raw = pl.concat(frames, how="diagonal")
    _validate_required_columns(raw)
    return raw


def _safe_report_stem(stem: str, max_len: int = 60) -> str:
    """用于输出文件名：去掉 Windows 非法字符，避免空串。"""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem)
    s = s.strip(" .") or "report"
    return s[:max_len]


def _unique_output_paths_for_per_file(paths: list[Path], ts: str) -> list[tuple[Path, Path]]:
    """返回 (源路径, 输出 xlsx 路径) 列表；同名净化 stem 时自动加后缀区分。"""
    stem_count: dict[str, int] = {}
    result: list[tuple[Path, Path]] = []
    for p in paths:
        base = _safe_report_stem(p.stem)
        n = stem_count.get(base, 0) + 1
        stem_count[base] = n
        suffix = "" if n == 1 else f"_{n}"
        out_name = f"bom_analysis_{base}{suffix}_{ts}.xlsx"
        result.append((p, OUTPUT_DIR / out_name))
    return result


# ---------------------------------------------------------------------------
# 清洗与聚合
# ---------------------------------------------------------------------------

def prepare(raw: pl.DataFrame) -> pl.DataFrame:
    raw = normalize_columns(raw)
    return (
        raw
        .with_columns([
            pl.col(COL_PRODUCT).cast(pl.String).str.strip_chars(),
            pl.col(COL_COMPONENT).cast(pl.String).str.strip_chars(),
            pl.col(COL_QTY).cast(pl.Float64, strict=False).fill_null(0.0),
            pl.col(COL_UNIT_PRICE).cast(pl.Float64, strict=False).fill_null(0.0),
        ])
        .filter(pl.col(COL_QTY) > 0)
    )


def aggregate_lines(d: pl.DataFrame) -> pl.DataFrame:
    """按 产品编码+组件编码 分组，汇总用量、取均价、取计量单位与组件名称。"""
    gcols = [COL_PRODUCT, COL_COMPONENT]
    agg_exprs: list[pl.Expr] = [
        pl.col(COL_QTY).sum().alias("MENGE合计"),
        pl.col(COL_UNIT_PRICE).mean().alias("组件单价"),
    ]
    if COL_UNIT in d.columns:
        agg_exprs.append(pl.col(COL_UNIT).first().alias("计量单位"))
    if COL_MAKTX in d.columns:
        agg_exprs.append(pl.col(COL_MAKTX).first().alias("组件名称"))
    if "产品价格" in d.columns:
        agg_exprs.append(pl.col("产品价格").first().alias("产品价格"))
    for extra in ("材料成本", "工费", "重量", "报价月份", "标杆工厂"):
        if extra in d.columns:
            agg_exprs.append(pl.col(extra).first().alias(extra))
    return d.group_by(gcols).agg(agg_exprs)


def enrich(lines: pl.DataFrame) -> pl.DataFrame:
    """补充每行的用量占比（%）与行原材料成本。"""
    product_total = (
        lines.group_by(COL_PRODUCT)
        .agg(pl.col("MENGE合计").sum().alias("_产品总MENGE"))
    )
    return (
        lines
        .join(product_total, on=COL_PRODUCT, how="left")
        .with_columns([
            (pl.col("MENGE合计") / pl.col("_产品总MENGE") * 100).round(4).alias("用量占比%"),
            (pl.col("MENGE合计") * pl.col("组件单价")).alias("行原材料成本"),
        ])
        .drop("_产品总MENGE")
    )


# ---------------------------------------------------------------------------
# 分析模块
# ---------------------------------------------------------------------------

def sheet_product_summary(lines: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「产品总成本排名」
    每个产品一行，含：组件种类数、产品总用量、原材料总成本，
    以及成本最高的那个组件（核心成本组件）的编码、名称、行成本、用量占比、成本占比。
    按原材料总成本从高到低排列。
    """
    base = lines.group_by(COL_PRODUCT).agg([
        pl.col(COL_COMPONENT).n_unique().alias("组件种类数"),
        pl.col("MENGE合计").sum().alias("产品总用量"),
        pl.col("行原材料成本").sum().alias("原材料总成本"),
    ])
    driver_exprs: list[pl.Expr] = [
        pl.col(COL_COMPONENT)
          .sort_by("行原材料成本", descending=True).first()
          .alias("核心成本组件编码"),
        pl.col("行原材料成本")
          .sort_by("行原材料成本", descending=True).first()
          .alias("核心成本组件行成本"),
        pl.col("用量占比%")
          .sort_by("行原材料成本", descending=True).first()
          .alias("核心成本组件用量占比%"),
    ]
    if "组件名称" in lines.columns:
        driver_exprs.append(
            pl.col("组件名称")
              .sort_by("行原材料成本", descending=True).first()
              .alias("核心成本组件名称")
        )
    drivers = lines.group_by(COL_PRODUCT).agg(driver_exprs)
    out = base.join(drivers, on=COL_PRODUCT, how="left")
    if "产品价格" in lines.columns:
        pp = lines.group_by(COL_PRODUCT).agg(pl.col("产品价格").first().alias("产品价格"))
        out = out.join(pp, on=COL_PRODUCT, how="left").with_columns(
            pl.when(
                (pl.col("产品价格").is_not_null())
                & (pl.col("产品价格") > 1e-12)
                & (pl.col("原材料总成本").is_not_null())
            )
            .then((pl.col("原材料总成本") / pl.col("产品价格") * 100).round(2))
            .otherwise(None)
            .alias("原材料成本占产品价格%")
        )
    return (
        out.with_columns(
            (pl.col("核心成本组件行成本") / pl.col("原材料总成本") * 100)
            .round(2)
            .alias("核心成本组件占产品成本%")
        )
        .sort("原材料总成本", descending=True)
    )


def sheet_bom_detail(lines: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「BOM明细_占比与成本」
    每个产品-组件对一行，列名均为业务可读的中文，
    含：产品编码、组件编码、组件名称、计量单位、MENGE合计、用量占比%、组件单价、行原材料成本。
    """
    col_order = [COL_PRODUCT, COL_COMPONENT]
    if "组件名称" in lines.columns:
        col_order.append("组件名称")
    if "计量单位" in lines.columns:
        col_order.append("计量单位")
    col_order += ["MENGE合计", "用量占比%", "组件单价", "行原材料成本"]
    for extra in ("产品价格", "材料成本", "工费", "重量", "报价月份"):
        if extra in lines.columns and extra not in col_order:
            col_order.append(extra)
    return lines.select([c for c in col_order if c in lines.columns])


def sheet_component_global_rank(lines: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「组件全局成本排名」
    每个组件一行，按"该组件在所有产品中的总成本贡献"从高到低排列。
    含：全局总成本贡献、全局成本占比%（该组件占所有产品原材料总成本的份额）、
        涉及产品数、全局总用量、平均用量占比%。
    这张表直接回答「哪个组件是全厂最大的成本驱动因素」。
    """
    total = lines.select(pl.col("行原材料成本").sum()).item()
    return (
        lines.group_by(COL_COMPONENT).agg([
            pl.col("行原材料成本").sum().alias("全局总成本贡献"),
            pl.col(COL_PRODUCT).n_unique().alias("涉及产品数"),
            pl.col("MENGE合计").sum().alias("全局总用量"),
            pl.col("用量占比%").mean().round(4).alias("平均用量占比%（在各产品中）"),
        ])
        .with_columns(
            (pl.col("全局总成本贡献") / total * 100).round(4).alias("全局成本占比%")
        )
        .sort("全局总成本贡献", descending=True)
    )


def sheet_cross_product_volatility(lines: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「跨产品用量波动」
    每个组件一行，按变异系数（CV = 标准差 / 均值）从高到低排列。
    CV 越高 → 该组件在不同产品中的配比差异越大（不标准化）；
    CV 趋近 0 → 无论用在哪个产品，占比几乎一致（高度标准化组件）。
    只统计出现在 ≥2 个产品中的组件（单一产品专用件 CV 无意义）。
    """
    return (
        lines.group_by(COL_COMPONENT).agg([
            pl.col("用量占比%").mean().round(4).alias("平均用量占比%"),
            pl.col("用量占比%").std().round(4).alias("用量占比标准差"),
            pl.col(COL_PRODUCT).n_unique().alias("涉及产品数"),
        ])
        .filter(pl.col("涉及产品数") >= 2)
        .with_columns(
            (pl.col("用量占比标准差") / pl.col("平均用量占比%")).round(4).alias("变异系数CV")
        )
        .sort("变异系数CV", descending=True, nulls_last=True)
    )


def top3_per_product(lines: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「各产品Top3成本组件」用数据：每个产品最多 3 行（按行原材料成本排名）。
    列：产品编码、成本排名(1/2/3)、组件编码、组件名称、MENGE合计、用量占比%、组件单价、行原材料成本。
    """
    col_order = [COL_PRODUCT, "成本排名", COL_COMPONENT]
    if "组件名称" in lines.columns:
        col_order.append("组件名称")
    if "计量单位" in lines.columns:
        col_order.append("计量单位")
    col_order += ["MENGE合计", "用量占比%", "组件单价", "行原材料成本"]

    ranked = (
        lines
        .with_columns(
            pl.col("行原材料成本")
            .rank(method="ordinal", descending=True)
            .over(COL_PRODUCT)
            .cast(pl.Int32)
            .alias("成本排名")
        )
        .filter(pl.col("成本排名") <= 3)
    )
    return (
        ranked
        .select([c for c in col_order if c in ranked.columns])
        .sort([COL_PRODUCT, "成本排名"])
    )


def sheet_price_history(raw: pl.DataFrame) -> pl.DataFrame:
    """
    工作表「价格历史明细」
    保留原始逐行价格与日期，供 BI 层按时间区间查询组件价格波动。
    列：产品编码、组件编码、组件名称（可选）、分类（可选）、组件单价、CREATEDATE（可选）
    """
    raw = normalize_columns(raw)
    cols = [COL_PRODUCT, COL_COMPONENT]
    for optional in (COL_MAKTX, COL_CATEGORY, COL_UNIT_PRICE, COL_CREATEDATE):
        if optional in raw.columns:
            cols.append(optional)

    df = raw.select([c for c in cols if c in raw.columns]).with_columns([
        pl.col(COL_PRODUCT).cast(pl.String).str.strip_chars(),
        pl.col(COL_COMPONENT).cast(pl.String).str.strip_chars(),
        pl.col(COL_UNIT_PRICE).cast(pl.Float64, strict=False).fill_null(0.0),
    ])
    if COL_CREATEDATE in df.columns:
        df = df.with_columns(pl.col(COL_CREATEDATE).cast(pl.String).str.strip_chars())
    return df.filter(pl.col(COL_UNIT_PRICE) > 0)


# ---------------------------------------------------------------------------
# 写出
# ---------------------------------------------------------------------------

def _write_sheet(writer: pd.ExcelWriter, df: pl.DataFrame, name: str) -> None:
    """将 Polars DataFrame 写入 Excel 工作表，并把产品/组件编码列强制保持为文本，
    防止 Excel 将长数字编码自动转为科学计数法。"""
    pdf = df.to_pandas()
    for col in pdf.columns:
        if col in (COL_PRODUCT, COL_COMPONENT, "核心成本组件编码"):
            pdf[col] = pdf[col].astype(str)
    pdf.to_excel(writer, sheet_name=name, index=False)


def run_pipeline(raw: pl.DataFrame, out_xlsx: Path, *, label: str = "") -> None:
    """假定 raw 已加载；日志从 [2/6] 起，与主流程中 [1/6] 加载衔接。"""
    prefix = f"{label} " if label else ""
    print(f"{prefix}      原始数据：{raw.height:,} 行 × {raw.width} 列")

    print(f"{prefix}[2/6] 清洗与聚合...")
    lines = enrich(aggregate_lines(prepare(raw)))
    print(f"      聚合后：{lines.height:,} 个产品-组件组合")

    print(f"{prefix}[3/6] 产品总成本排名...")
    tbl_summary = sheet_product_summary(lines)

    print(f"{prefix}[4/6] 组件全局成本排名...")
    tbl_comp_rank = sheet_component_global_rank(lines)

    print(f"{prefix}[5/6] 跨产品用量波动...")
    tbl_volatility = sheet_cross_product_volatility(lines)

    print(f"{prefix}[6/6] 各产品 Top3 成本组件...")
    tbl_top3 = top3_per_product(lines)

    print(f"{prefix}写入 Excel：{out_xlsx.name} …")
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        _write_sheet(writer, tbl_summary,             "产品总成本排名")
        _write_sheet(writer, sheet_bom_detail(lines), "BOM明细_占比与成本")
        _write_sheet(writer, tbl_comp_rank,           "组件全局成本排名")
        _write_sheet(writer, tbl_volatility,          "跨产品用量波动")
        _write_sheet(writer, tbl_top3,                "各产品Top3成本组件")

    print(f"{prefix}  已写出：{out_xlsx}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="读取 web/data/ 下 Excel，生成 BOM 成本分析报表到 web/output/。"
    )
    parser.add_argument(
        "--per-file",
        action="store_true",
        help="每个 Excel 单独生成一份 bom_analysis_<文件名>_<时间戳>.xlsx（默认：全部合并为一份）",
    )
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if args.per_file:
        paths = _excel_paths_in_data()
        if not paths:
            _print_no_excel_error()
            sys.exit(1)
        pairs = _unique_output_paths_for_per_file(paths, ts)
        print(f"模式：每文件一份报告（共 {len(pairs)} 个 Excel）\n")
        for i, (src, out_xlsx) in enumerate(pairs, start=1):
            tag = f"[{i}/{len(pairs)} {src.name}]"
            print(f"{tag} [1/6] 加载数据...")
            raw = _read_one_excel(src)
            _validate_required_columns(raw)
            run_pipeline(raw, out_xlsx, label=tag)
        print(f"\n完成！共生成 {len(pairs)} 个 Excel 报告（时间戳 {ts}）。")
        print("  工作表：产品总成本排名 | BOM明细_占比与成本 | 组件全局成本排名 | 跨产品用量波动 | 各产品Top3成本组件")
        return

    print("模式：合并 data/ 下全部 Excel 为一份报告\n")
    print("[1/6] 加载数据...")
    raw = load_all_excels()
    out_xlsx = OUTPUT_DIR / f"bom_analysis_{ts}.xlsx"
    run_pipeline(raw, out_xlsx)
    print(f"\n完成！")
    print(f"  Excel 报告：{out_xlsx}")
    print(f"  工作表：产品总成本排名 | BOM明细_占比与成本 | 组件全局成本排名 | 跨产品用量波动 | 各产品Top3成本组件")


if __name__ == "__main__":
    main()
