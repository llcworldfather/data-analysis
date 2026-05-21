# -*- coding: utf-8 -*-
"""求两个 Excel 中「同时出现」的产品编码交集。

默认路径（可改）：
  - 历史清单：Desktop/25年至今产品历史清单.xlsx  工作表「数据」列 MATNR
  - BOM 导出：Desktop/EXPORT20260323.XLSX         第一张表列 产品编码

用法：
  python scripts/intersect_product_codes_two_excels.py
  python scripts/intersect_product_codes_two_excels.py --hist "D:/a.xlsx" --bom "D:/b.xlsx"
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


def _norm_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def main() -> None:
    desktop = Path.home() / "Desktop"
    ap = argparse.ArgumentParser(description="两表产品编码交集（历史 MATNR ∩ BOM 产品编码）")
    ap.add_argument(
        "--hist",
        type=Path,
        default=desktop / "25年至今产品历史清单.xlsx",
        help="历史清单 xlsx",
    )
    ap.add_argument(
        "--bom",
        type=Path,
        default=desktop / "EXPORT20260323.XLSX",
        help="BOM 导出 xlsx",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "output",
        help="输出目录",
    )
    args = ap.parse_args()

    h = pd.read_excel(args.hist, sheet_name="数据", usecols=["MATNR"])
    b = pd.read_excel(args.bom, sheet_name=0, usecols=["产品编码"])

    set_h = set(_norm_series(h["MATNR"]).dropna())
    set_b = set(_norm_series(b["产品编码"]).dropna())
    for s in (set_h, set_b):
        s.discard("")
        s.discard("nan")

    inter = sorted(set_h & set_b, key=lambda x: (len(x), x))

    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    txt = out_dir / "产品编码_两表交集.txt"
    xlsx = out_dir / "产品编码_两表交集.xlsx"
    txt.write_text("\n".join(inter), encoding="utf-8")
    pd.DataFrame({"产品编码": inter}).to_excel(xlsx, index=False)

    print("历史清单 唯一 MATNR:", len(set_h))
    print("BOM 唯一 产品编码:", len(set_b))
    print("交集个数:", len(inter))
    print("仅历史清单:", len(set_h - set_b))
    print("仅 BOM:", len(set_b - set_h))
    print("TXT :", txt)
    print("XLSX:", xlsx)


if __name__ == "__main__":
    main()
