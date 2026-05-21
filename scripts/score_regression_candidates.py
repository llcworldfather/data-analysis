# -*- coding: utf-8 -*-
"""在三表 + 可选新 BOM 中，为每个产品打「回归适合度」分并排序。"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import polars as pl

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"
sys.path.insert(0, str(WEB))
sys.path.insert(0, str(ROOT))

from process_excel import norm_material_code, normalize_columns  # noqa: E402
from cost_sim_predict import (  # noqa: E402
    build_product_cost_history,
    build_product_price_timeline,
    expand_product_cost_history_dict,
    history_per_kg_pair,
    _linear_fit_predict,
    _regression_fit_grade,
)

DATA = Path(r"C:\Users\00109151\Desktop\数据集")
OUT = ROOT / "output" / "regression_best_candidates.csv"


def _norm_codes(s: pd.Series) -> pd.Series:
    return norm_material_code(s)


def _load_bom_combined() -> pd.DataFrame:
    """合并原始 BOM 清单 + 25年导出清单。"""
    paths = [
        DATA / "报价BOM历史清单.xls",
        DATA / "报价BOM历史清单_由25年BOM导出.xlsx",
    ]
    frames = []
    for p in paths:
        if p.exists():
            df = pd.read_excel(p)
            df["_来源文件"] = p.name
            frames.append(df)
    if not frames:
        raise FileNotFoundError("未找到任何 BOM 文件")
    bom = pd.concat(frames, ignore_index=True)
    if "所属产品" in bom.columns and "产品编码" not in bom.columns:
        pass
    elif "产品编码" in bom.columns and "所属产品" not in bom.columns:
        bom = bom.rename(columns={"产品编码": "所属产品"})
    return bom


def _bom_to_polars(bom_pd: pd.DataFrame) -> pl.DataFrame:
    rename = {
        "所属产品": "产品编码",
        "材料编码": "组件编码",
        "组件数量": "MENGE",
        "材料单价": "组件单价",
        "基本单位": "MEINS",
        "材料型号": "MAKTX",
        "生价日期": "CREATEDATE",
    }
    df = bom_pd.rename(columns={k: v for k, v in rename.items() if k in bom_pd.columns})
    if "CREATEDATE" in df.columns:
        df["CREATEDATE"] = df["CREATEDATE"].astype(str).str.replace(r"\.0+$", "", regex=True)
    return pl.from_pandas(df)


def score_product(
    pid: str,
    hist: list[dict],
    in_bom: bool,
    in_price: bool,
    price_periods: int,
    bom_units: set[str],
) -> dict:
    pairs = []
    for h in hist:
        bm, pp = history_per_kg_pair(h)
        if bm is not None and pp is not None and bm > 1e-12 and pp > 0:
            pairs.append((float(bm), float(pp)))

    n = len(pairs)
    r2 = 0.0
    intercept = slope = None
    x_std = y_std = 0.0
    if n >= 2:
        xs = np.array([p[0] for p in pairs])
        ys = np.array([p[1] for p in pairs])
        x_std = float(np.std(xs))
        y_std = float(np.std(ys))
    if n >= 3 and x_std > 1e-9:
        xs = np.array([p[0] for p in pairs])
        ys = np.array([p[1] for p in pairs])
        _, r2, _, intercept, slope = _linear_fit_predict(xs, ys, float(np.mean(xs)))

    # 适合度分（越高越好，满分约 100）
    score = 0.0
    if in_bom and in_price:
        score += 25
    score += min(25, n * 1.2)  # 样本数，约 21 期满分
    score += min(25, r2 * 25) if n >= 3 else 0
    if x_std > 1e-6:
        score += 12.5
    if y_std > 1e-9:
        score += 12.5
    if r2 >= 0.35:
        score += 10

    gaps = []
    if not in_bom:
        gaps.append("缺BOM")
    if not in_price:
        gaps.append("缺价格清单")
    if price_periods < 3:
        gaps.append(f"价格表仅{price_periods}期")
    if n < 3:
        gaps.append(f"有效配对仅{n}期")
    elif x_std <= 1e-6:
        gaps.append("BOM材料/kg各期相同(单位或BOM未按月变)")
    elif y_std <= 1e-9:
        gaps.append("产品价/kg几乎不变")
    elif r2 < 0.35:
        gaps.append(f"R²={r2:.2f}未达0.35")
    if "ST" in bom_units and "KG" not in bom_units:
        gaps.append("BOM单位为ST需换算kg")

    return {
        "产品编码": pid,
        "适合度分": round(score, 1),
        "有效样本期数": n,
        "R2": round(r2, 4) if n >= 3 else None,
        "拟合等级": _regression_fit_grade(r2, n),
        "BOM材料kg标准差": round(x_std, 6),
        "产品价kg标准差": round(y_std, 8),
        "截距": round(intercept, 6) if intercept is not None else None,
        "斜率": round(slope, 8) if slope is not None else None,
        "在BOM": in_bom,
        "在价格表": in_price,
        "价格表有重量期数": price_periods,
        "BOM单位": ",".join(sorted(bom_units)) if bom_units else "",
        "待补足": "；".join(gaps) if gaps else "已满足回归主条件",
        "推荐采用回归": n >= 3 and r2 >= 0.35 and x_std > 1e-6,
    }


def main() -> None:
    bom_pd = _load_bom_combined()
    price_xl = pd.read_excel(DATA / "产品价格历史清单.xls")

    bom_pd["所属产品"] = _norm_codes(bom_pd["所属产品"])
    price_xl["产品编码"] = _norm_codes(price_xl["产品编码"])

    bom_codes = set(bom_pd["所属产品"]) - {"", "nan"}
    price_codes = set(price_xl["产品编码"]) - {"", "nan"}
    overlap = bom_codes & price_codes

    px = price_xl.copy()
    px["重量"] = pd.to_numeric(px.get("重量"), errors="coerce")
    price_periods_map: dict[str, int] = {}
    for pid, grp in px.groupby("产品编码"):
        w = grp[grp["重量"].notna() & (grp["重量"] > 0)]
        price_periods_map[pid] = (
            w["报价月份"].nunique() if "报价月份" in w.columns and len(w) else 0
        )

    bom_units_map: dict[str, set[str]] = defaultdict(set)
    if "基本单位" in bom_pd.columns:
        for pid, grp in bom_pd.groupby("所属产品"):
            bom_units_map[pid] = set(grp["基本单位"].astype(str).str.strip().unique()) - {
                "",
                "nan",
            }

    bom_raw = _bom_to_polars(bom_pd)
    pch = build_product_cost_history(bom_raw)
    ppt = build_product_price_timeline(DATA / "产品价格历史清单.xls")

    pbom: dict[str, list[dict]] = defaultdict(list)
    prep = normalize_columns(bom_raw)
    agg = prep.group_by(["产品编码", "组件编码"]).agg(
        pl.col("MENGE").sum().alias("MENGE合计"),
        pl.col("组件单价").mean().alias("组件单价"),
    )
    for row in agg.to_dicts():
        pid = str(row.get("产品编码") or "").strip()
        if pid:
            pbom[pid].append(
                {
                    "组件编码": str(row.get("组件编码") or ""),
                    "MENGE合计": float(row.get("MENGE合计") or 0),
                    "组件单价": float(row.get("组件单价") or 0),
                }
            )

    expanded = expand_product_cost_history_dict(
        dict(pch), dict(pbom), None, ppt
    )

    all_pids = sorted(bom_codes | price_codes)
    rows = []
    for pid in all_pids:
        hist = expanded.get(pid, pch.get(pid, []))
        rows.append(
            score_product(
                pid,
                hist,
                pid in bom_codes,
                pid in price_codes,
                price_periods_map.get(pid, 0),
                bom_units_map.get(pid, set()),
            )
        )

    df = pd.DataFrame(rows).sort_values(
        ["推荐采用回归", "适合度分", "R2", "有效样本期数"],
        ascending=[False, False, False, False],
        na_position="last",
    )
    OUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT, index=False, encoding="utf-8-sig")

    ready = df[df["推荐采用回归"]]
    top = df.head(15)

    print("=" * 60)
    print("BOM 文件:", [p.name for p in [DATA / "报价BOM历史清单.xls", DATA / "报价BOM历史清单_由25年BOM导出.xlsx"] if p.exists()])
    print(f"BOM 产品 {len(bom_codes)} | 价格表 {len(price_codes)} | 交集 {len(overlap)}")
    print(f"可直接采用回归 (R²≥0.35 且 X 有波动): {len(ready)} 个")
    if len(ready):
        print(ready[["产品编码", "适合度分", "有效样本期数", "R2", "待补足"]].to_string(index=False))
    print()
    print("--- 交集内 TOP（按适合度，即使 R² 未达标）---")
    sub = df[df["在BOM"] & df["在价格表"]].head(10)
    if len(sub):
        print(sub[["产品编码", "适合度分", "有效样本期数", "R2", "BOM材料kg标准差", "产品价kg标准差", "待补足"]].to_string(index=False))
    else:
        print("(无交集产品 — 见下方「最接近」)")
    print()
    print("--- 全表适合度 TOP 15 ---")
    print(top[["产品编码", "适合度分", "有效样本期数", "R2", "在BOM", "在价格表", "待补足"]].to_string(index=False))
    print(f"\n完整表: {OUT}")


if __name__ == "__main__":
    main()
