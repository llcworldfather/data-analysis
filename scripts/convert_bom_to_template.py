# -*- coding: utf-8 -*-
"""
将 SAP 宽表 BOM 导出转为「报价BOM历史清单」模板列格式。

用法:
  python scripts/convert_bom_to_template.py
  python scripts/convert_bom_to_template.py "源.xlsx" "输出.xlsx"
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

DATASET = Path(r"C:\Users\00109151\Desktop\数据集")
DEFAULT_SOURCE = DATASET / "25年至今的BOM清单.xlsx"
TEMPLATE_COLS = [
    "生价日期",
    "所属产品",
    "材料编码",
    "基本单位",
    "材料型号",
    "组件数量",
    "材料单价",
    "报价号",
    "报价流水号",
]
DEFAULT_OUT = DATASET / "报价BOM历史清单_由25年BOM导出.xlsx"


def _norm_code(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    s = re.sub(r"\.0+$", "", s)
    return s.replace("nan", "").replace("None", "")


def _norm_date(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)) and not pd.isna(val):
        return str(int(val))
    s = str(val).strip().replace("-", "").replace("/", "")[:8]
    digits = re.sub(r"\D", "", s)
    return digits[:8] if digits else ""


def _detect_product_component_cols(df: pd.DataFrame) -> tuple[str, str]:
    """
    判断成品列与子件列。
    - 若 MATNR/ZMATNR 几乎唯一、IDNRK 多种 → 成品在 MATNR，子件在 IDNRK
    - 若 IDNRK 几乎唯一、MATNR 多种 → 成品在 IDNRK，子件在 MATNR（如 1020110079 那类导出）
    """
    comp_fallback = "MATNR" if "MATNR" in df.columns else "ZMATNR"
    if comp_fallback not in df.columns and "IDNRK" not in df.columns:
        raise ValueError(f"无法识别 BOM 列，现有: {list(df.columns)}")

    idn = df["IDNRK"].map(_norm_code) if "IDNRK" in df.columns else pd.Series(dtype=str)
    mat = (
        df[comp_fallback].map(_norm_code)
        if comp_fallback in df.columns
        else pd.Series(dtype=str)
    )
    idn_n = idn.replace("", pd.NA).nunique()
    mat_n = mat.replace("", pd.NA).nunique()

    if mat_n <= max(3, idn_n // 2) and idn_n > mat_n:
        return comp_fallback, "IDNRK"
    if idn_n <= max(3, mat_n // 2) and mat_n > idn_n:
        return "IDNRK", comp_fallback
    if idn_n == 1 and mat_n > 1:
        return "IDNRK", comp_fallback
    if mat_n == 1 and idn_n > 1:
        return comp_fallback, "IDNRK"
    return comp_fallback, "IDNRK"


def convert_bom_df(df: pd.DataFrame) -> pd.DataFrame:
    """SAP/导出宽表 → 报价BOM历史清单列。"""
    product_col, component_col = _detect_product_component_cols(df)
    if product_col not in df.columns:
        raise ValueError(f"源表缺少成品列 {product_col}")

    name_col = "MAKTX" if "MAKTX" in df.columns else "ZMAKTX"
    out = pd.DataFrame(
        {
            "生价日期": df["CREATEDATE"].map(_norm_date) if "CREATEDATE" in df.columns else "",
            "所属产品": df[product_col].map(_norm_code),
            "材料编码": df[component_col].map(_norm_code) if component_col in df.columns else "",
            "基本单位": df["MEINS"].astype(str).str.strip() if "MEINS" in df.columns else "",
            "材料型号": df[name_col].fillna(df.get("ZMAKTX", "")).astype(str).str.strip()
            if name_col in df.columns
            else (df["ZMAKTX"].astype(str).str.strip() if "ZMAKTX" in df.columns else ""),
            "组件数量": pd.to_numeric(df["MENGE"], errors="coerce")
            if "MENGE" in df.columns
            else 0,
            "材料单价": pd.to_numeric(df["ZDJ"], errors="coerce")
            if "ZDJ" in df.columns
            else 0,
            "报价号": df["ZBJNO"].map(_norm_code) if "ZBJNO" in df.columns else "",
            "报价流水号": df["ZSNO"].map(_norm_code) if "ZSNO" in df.columns else "",
        }
    )
    out = out[TEMPLATE_COLS]
    out = out[(out["所属产品"] != "") & (out["材料编码"] != "")].copy()
    return out.sort_values(
        ["生价日期", "所属产品", "材料编码", "报价流水号"]
    ).reset_index(drop=True)


def main(source: Path | None = None, out_path: Path | None = None) -> Path:
    src = source or DEFAULT_SOURCE
    if not src.exists():
        raise FileNotFoundError(f"未找到源文件: {src}")
    raw = pd.read_excel(src, sheet_name=0)
    out_df = convert_bom_df(raw)
    prod_col, comp_col = _detect_product_component_cols(raw)
    dest = out_path or DEFAULT_OUT
    dest.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_excel(dest, index=False, sheet_name="Sheet1")
    print(f"源文件: {src.name}")
    print(f"列映射: 所属产品←{prod_col}  材料编码←{comp_col}")
    print(f"源行数: {len(raw)} → 导出行数: {len(out_df)}")
    print(f"所属产品: {out_df['所属产品'].nunique()} 个 → {sorted(out_df['所属产品'].unique())}")
    print(f"材料编码: {out_df['材料编码'].nunique()} 个")
    print(f"生价日期: {out_df['生价日期'].nunique()} 个")
    print(f"已写入: {dest}")
    return dest


if __name__ == "__main__":
    src_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    out_arg = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    if src_arg is None:
        main()
    else:
        main(src_arg, out_arg)
