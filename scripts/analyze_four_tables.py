# -*- coding: utf-8 -*-
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import polars as pl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "web"))
sys.path.insert(0, str(ROOT))

from process_excel import normalize_columns
from cost_sim_predict import (
    build_product_cost_history,
    build_product_price_timeline,
    expand_product_cost_history_dict,
    history_per_kg_pair,
    _linear_fit_predict,
    _regression_fit_grade,
)

DATA = Path(r"C:\Users\00109151\Desktop\数据集")


def norm(s):
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
        .replace({"nan": "", "None": ""})
    )


def score_pid(pid, bom_raw, expanded, pch, price_periods):
    hist = expanded.get(pid, pch.get(pid, []))
    pairs = []
    for h in hist:
        bm, pp = history_per_kg_pair(h)
        if bm and pp and bm > 1e-12 and pp > 0:
            pairs.append((bm, pp))
    n = len(pairs)
    r2 = 0.0
    x_std = y_std = 0.0
    if n >= 2:
        xs = np.array([p[0] for p in pairs])
        ys = np.array([p[1] for p in pairs])
        x_std, y_std = float(np.std(xs)), float(np.std(ys))
    if n >= 3 and x_std > 1e-9:
        _, r2, _, _, _ = _linear_fit_predict(
            np.array([p[0] for p in pairs]),
            np.array([p[1] for p in pairs]),
            float(np.mean([p[0] for p in pairs])),
        )
    return {
        "有效样本": n,
        "R2": round(r2, 4),
        "等级": _regression_fit_grade(r2, n),
        "BOMkg波动": x_std,
        "产品价kg波动": y_std,
        "可用回归": n >= 3 and r2 >= 0.35 and x_std > 1e-6,
        "价格表期数": price_periods,
    }


def main():
    bom_list = [
        DATA / "报价BOM历史清单.xls",
        DATA / "报价BOM历史清单_由25年BOM导出.xlsx",
    ]
    bom = pd.concat(
        [pd.read_excel(p) for p in bom_list if p.exists()], ignore_index=True
    )
    price = pd.read_excel(DATA / "产品价格历史清单.xls")

    bom["所属产品"] = norm(bom["所属产品"])
    bom["材料编码"] = norm(bom["材料编码"])
    price["产品编码"] = norm(price["产品编码"])

    mid = "1010010269"
    prods = sorted(bom.loc[bom["材料编码"] == mid, "所属产品"].unique())
    print("材料", mid, "关联产品数(BOM):", len(prods))
    print("  样例:", prods[:8])

    bom_raw = pl.from_pandas(
        bom.rename(
            columns={
                "所属产品": "产品编码",
                "材料编码": "组件编码",
                "组件数量": "MENGE",
                "材料单价": "组件单价",
                "基本单位": "MEINS",
                "材料型号": "MAKTX",
                "生价日期": "CREATEDATE",
            }
        )
    )
    pch = build_product_cost_history(bom_raw)
    ppt = build_product_price_timeline(DATA / "产品价格历史清单.xls")
    pbom = defaultdict(list)
    for row in (
        normalize_columns(bom_raw)
        .group_by(["产品编码", "组件编码"])
        .agg(
            pl.col("MENGE").sum().alias("MENGE合计"),
            pl.col("组件单价").mean().alias("组件单价"),
        )
        .to_dicts()
    ):
        pid = str(row["产品编码"]).strip()
        pbom[pid].append(
            {
                "组件编码": str(row["组件编码"]),
                "MENGE合计": float(row["MENGE合计"] or 0),
                "组件单价": float(row["组件单价"] or 0),
            }
        )
    expanded = expand_product_cost_history_dict(dict(pch), dict(pbom), None, ppt)

    price_periods = {}
    for pid, grp in price.groupby("产品编码"):
        w = grp[pd.to_numeric(grp["重量"], errors="coerce").notna()]
        price_periods[pid] = grp["报价月份"].nunique() if "报价月份" in grp.columns else 0

    print("\n=== 截图相关产品 回归诊断 ===")
    for pid in ["1020361080", "1020110079", "1020110077", "1030011633"]:
        in_bom = pid in set(bom["所属产品"])
        r = score_pid(pid, bom_raw, expanded, pch, price_periods.get(pid, 0))
        print(
            pid,
            "BOM=",
            in_bom,
            "用材料1010010269=",
            pid in prods,
            r,
        )


if __name__ == "__main__":
    main()
