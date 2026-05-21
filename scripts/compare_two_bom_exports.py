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

from process_excel import norm_material_code, normalize_columns
from cost_sim_predict import (
    build_product_cost_history,
    build_product_price_timeline,
    expand_product_cost_history_dict,
    history_per_kg_pair,
    _linear_fit_predict,
    _regression_fit_grade,
)

DATA = Path(r"C:\Users\00109151\Desktop\数据集")


def score_bom(bom_path: Path, pid: str, price: pd.DataFrame, ppt) -> dict:
    bom_pd = pd.read_excel(bom_path)
    bom_raw = pl.from_pandas(
        bom_pd.rename(
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
        p = str(row["产品编码"]).strip()
        pbom[p].append(
            {
                "组件编码": str(row["组件编码"]),
                "MENGE合计": float(row["MENGE合计"] or 0),
                "组件单价": float(row["组件单价"] or 0),
            }
        )
    exp = expand_product_cost_history_dict(dict(pch), dict(pbom), None, ppt)
    hist = exp.get(pid, pch.get(pid, []))
    pairs = []
    for h in hist:
        bm, pp = history_per_kg_pair(h)
        if bm is not None and pp is not None and bm > 1e-12 and pp > 0:
            pairs.append((float(bm), float(pp)))

    n = len(pairs)
    r2 = x_std = y_std = 0.0
    if n >= 2:
        xs = np.array([p[0] for p in pairs])
        ys = np.array([p[1] for p in pairs])
        x_std, y_std = float(np.std(xs)), float(np.std(ys))
    if n >= 3 and x_std > 1e-9:
        _, r2, _, _, _ = _linear_fit_predict(xs, ys, float(np.mean([p[0] for p in pairs])))

    bom_pd = bom_pd.copy()
    bom_pd["_line"] = pd.to_numeric(bom_pd["组件数量"], errors="coerce") * pd.to_numeric(
        bom_pd["材料单价"], errors="coerce"
    )
    monthly_total = bom_pd.groupby("生价日期")["_line"].sum()
    monthly_zdj = bom_pd.groupby("生价日期")["材料单价"].mean()

    pr = price[price["产品编码"] == pid]
    return {
        "产品": pid,
        "BOM行数": len(bom_pd),
        "生价日数": bom_pd["生价日期"].nunique(),
        "材料种数": bom_pd["材料编码"].nunique(),
        "单位": ",".join(sorted(set(bom_pd["基本单位"].astype(str)))),
        "在价格表": pid in set(price["产品编码"]),
        "价格表月份数": pr["报价月份"].nunique() if len(pr) else 0,
        "有效回归样本": n,
        "R2": round(r2, 4),
        "等级": _regression_fit_grade(r2, n),
        "BOM材料kg波动": round(x_std, 4),
        "产品价kg波动": round(y_std, 6),
        "各月材料合计波动": round(float(monthly_total.std()), 4) if len(monthly_total) > 1 else 0,
        "各月均价波动": round(float(monthly_zdj.std()), 4) if len(monthly_zdj) > 1 else 0,
        "可采用回归": n >= 3 and r2 >= 0.35 and x_std > 1e-6,
    }


def main():
    price = pd.read_excel(DATA / "产品价格历史清单.xls")
    price["产品编码"] = norm_material_code(price["产品编码"])
    ppt = build_product_price_timeline(DATA / "产品价格历史清单.xls")

    new_path = DATA / "报价BOM历史清单_由25年BOM(1)导出.xlsx"
    new = score_bom(new_path, "8010010001", price, ppt)

    # 旧文件：若导出 xlsx 不在，则从源表现场转换
    old_export = DATA / "报价BOM历史清单_由25年BOM导出.xlsx"
    old_source = DATA / "25年至今的BOM清单.xlsx"
    if old_export.exists():
        old = score_bom(old_export, "1020110079", price, ppt)
        old_note = "已用导出文件"
    elif old_source.exists():
        sys.path.insert(0, str(ROOT / "scripts"))
        from convert_bom_to_template import convert_bom_df

        tmp = DATA / "_tmp_old_bom_export.xlsx"
        convert_bom_df(pd.read_excel(old_source)).to_excel(tmp, index=False)
        old = score_bom(tmp, "1020110079", price, ppt)
        old_note = "从25年至今的BOM清单.xlsx现场转换"
        tmp.unlink(missing_ok=True)
    else:
        old = {
            "产品": "1020110079",
            "说明": "旧源/导出不在数据集目录，沿用上次测算",
            "有效回归样本": 17,
            "R2": 0.0,
            "BOM材料kg波动": 0.0,
            "产品价kg波动": 0.000456,
            "在价格表": True,
            "价格表月份数": 17,
            "单位": "ST",
            "各月材料合计波动": 0,
            "可采用回归": False,
        }
        old_note = "历史测算"

    print("OLD", old_note, old)
    print("NEW", new)
    print()
    if new.get("可采用回归") and not old.get("可采用回归"):
        print("VERDICT: 新文件更好")
    elif old.get("可采用回归") and not new.get("可采用回归"):
        print("VERDICT: 旧文件更好")
    else:
        print("VERDICT: 两者都未达标；比较结构潜力")


if __name__ == "__main__":
    main()
