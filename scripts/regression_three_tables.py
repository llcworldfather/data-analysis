# -*- coding: utf-8 -*-
"""
三张表联合回归分析（numpy 最小二乘，与根目录 数据分析.py 思路一致）

输入（默认路径可在命令行覆盖）:
  1) EXPORT BOM：产品编码、组件编码、MENGE、组件单价、生价日期
  2) 25年至今产品历史清单：MATNR、CREATEDATE、DMBTR（IDNRK 若全空则忽略）
  3) 材料所属产品：ZMATNR、IDNRK（用于汇总校验；可选过滤 BOM 行）

输出: output/regression_three_tables_<时间戳>.xlsx

「生价日期」与历史表 CREATEDATE 均经 bom_date_key.calendar_date_key_yyyymmdd 规范为 YYYYMMDD 后按 (产品编码=MATNR, 日期) 键合并。
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from bom_date_key import calendar_date_key_yyyymmdd

OUTPUT_DIR = ROOT / "output"

MAX_FEATURES_PER_PRODUCT = 60
RIDGE_ALPHA = 1e-2


def _lstsq_predict_one_product(
    y: np.ndarray, X: np.ndarray, *, ridge_alpha: float = RIDGE_ALPHA
) -> Tuple[np.ndarray, Optional[np.ndarray], float]:
    """X 第一列为常数项；岭回归缓解小样本过拟合。返回 (预测, beta, r2_in_sample)。"""
    n, p = X.shape
    valid = np.isfinite(X).all(axis=1) & np.isfinite(y)
    nv = int(valid.sum())
    fill = float(np.nanmean(y[valid])) if nv else np.nan
    pred = np.full(n, fill, dtype=float)
    if nv < 2:
        return pred, None, float("nan")

    Xv, yv = X[valid], y[valid]
    try:
        if ridge_alpha > 0 and p > 1:
            reg = np.eye(p, dtype=np.float64) * float(ridge_alpha)
            reg[0, 0] = 0.0
            beta = np.linalg.solve(Xv.T @ Xv + reg, Xv.T @ yv)
        else:
            beta, _, _, _ = np.linalg.lstsq(Xv, yv, rcond=None)
        pred = (X @ beta).astype(float, copy=False)
        resid = yv - Xv @ beta
        ss_res = float(np.dot(resid, resid))
        y_mean = float(np.mean(yv))
        ss_tot = float(np.dot(yv - y_mean, yv - y_mean))
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 1e-18 else float("nan")
        return pred, beta.astype(float, copy=False), r2
    except Exception:
        return pred, None, float("nan")


def _norm_id(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def load_export(path: Path) -> pd.DataFrame:
    cols = ["产品编码", "组件编码", "MENGE", "组件单价", "生价日期"]
    df = pd.read_excel(path, sheet_name=0, usecols=cols)
    df["产品编码"] = df["产品编码"].map(_norm_id)
    df["组件编码"] = df["组件编码"].map(_norm_id)
    df["生价日期"] = df["生价日期"].map(calendar_date_key_yyyymmdd)
    df["MENGE"] = pd.to_numeric(df["MENGE"], errors="coerce").fillna(0.0)
    df["组件单价"] = pd.to_numeric(df["组件单价"], errors="coerce").fillna(0.0)
    df = df[(df["产品编码"] != "") & (df["组件编码"] != "") & (df["生价日期"] != "")]
    df["_line"] = df["MENGE"] * df["组件单价"]
    return df


def load_material_map(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, usecols=["ZMATNR", "IDNRK"])
    df["ZMATNR"] = df["ZMATNR"].map(_norm_id)
    df["IDNRK"] = df["IDNRK"].map(_norm_id)
    df = df[(df["ZMATNR"] != "") & (df["IDNRK"] != "")]
    return df.drop_duplicates()


def load_history(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, usecols=["MATNR", "CREATEDATE", "DMBTR"])
    df["MATNR"] = df["MATNR"].map(_norm_id)
    df["CREATEDATE"] = df["CREATEDATE"].map(calendar_date_key_yyyymmdd)
    df["DMBTR"] = pd.to_numeric(df["DMBTR"], errors="coerce").fillna(0.0)
    df = df[(df["MATNR"] != "") & (df["CREATEDATE"] != "")]
    agg = (
        df.groupby(["MATNR", "CREATEDATE"], as_index=False)["DMBTR"]
        .sum()
        .rename(columns={"DMBTR": "历史清单_DMBTR合计"})
    )
    return agg


def run(
    export_path: Path,
    history_path: Path,
    material_path: Path,
    *,
    use_material_filter: bool,
) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = OUTPUT_DIR / f"regression_three_tables_{ts}.xlsx"

    print("读取 BOM 导出 …", flush=True)
    bom = load_export(export_path)
    print(f"  BOM 行数 {len(bom):,}", flush=True)

    print("读取 材料所属产品 …", flush=True)
    mmap = load_material_map(material_path)
    print(f"  映射键 {len(mmap):,} 行（去重后）", flush=True)

    if use_material_filter:
        bom_f = bom.merge(
            mmap,
            left_on=["产品编码", "组件编码"],
            right_on=["ZMATNR", "IDNRK"],
            how="inner",
        )
        bom_f = bom_f.drop(columns=["ZMATNR", "IDNRK"], errors="ignore")
        print(
            f"  经映射过滤后 BOM 行数 {len(bom_f):,}（保留 {len(bom_f) / max(len(bom), 1):.2%}）",
            flush=True,
        )
        if len(bom_f) < 1000:
            print("  警告：过滤后行数过少，已自动退回为「不过滤」", flush=True)
            bom_f = bom
    else:
        bom_f = bom

    print("按产品透视并回归（组件单价 → BOM 总成本，避免全表宽透视占内存）…", flush=True)
    out_rows: list[dict] = []
    r2_list: list[tuple[str, int, float]] = []
    coef_rows: list[dict] = []

    grp = bom_f.groupby("产品编码", sort=False)
    n_products = grp.ngroups
    for gi, (matnr, g) in enumerate(grp):
        if gi % 5000 == 0:
            print(f"  回归进度 {gi}/{n_products}", flush=True)
        matnr = str(matnr)
        costs = g.groupby("生价日期", as_index=True)["_line"].sum()
        costs = costs.replace([np.inf, -np.inf], np.nan).dropna()
        costs.name = "BOM原材料总成本"
        try:
            pv = g.pivot_table(
                index="生价日期",
                columns="组件编码",
                values="组件单价",
                aggfunc="mean",
            ).sort_index()
        except Exception:
            continue
        pv = pv.replace([np.inf, -np.inf], np.nan).ffill().bfill()
        common = costs.index.intersection(pv.index)
        n_common = len(common)
        if n_common == 0:
            continue
        if n_common < 2:
            r2_list.append((matnr, n_common, float("nan")))
            for dt in common:
                yv = float(costs.loc[dt])
                out_rows.append(
                    {
                        "生价日期": dt,
                        "产品编码": matnr,
                        "BOM原材料总成本": yv,
                        "预测BOM总成本": yv,
                    }
                )
            continue

        pv2 = pv.loc[common]
        yg = costs.loc[common].to_numpy(dtype=np.float64, copy=True)
        Xg = pv2.to_numpy(dtype=np.float64, copy=True)
        X_columns = [str(c) for c in pv2.columns]

        col_var = np.nanvar(Xg, axis=0, ddof=0)
        col_mean = np.nanmean(np.abs(Xg), axis=0)
        var_floor = np.maximum(1e-6, col_mean * 1e-5)
        use_idx = np.flatnonzero(np.isfinite(col_var) & (col_var > var_floor))
        if use_idx.size == 0:
            r2_list.append((matnr, n_common, float("nan")))
            fill = float(np.nanmean(yg))
            for dt in common:
                out_rows.append(
                    {
                        "生价日期": dt,
                        "产品编码": matnr,
                        "BOM原材料总成本": float(costs.loc[dt]),
                        "预测BOM总成本": fill,
                    }
                )
            continue

        max_k = min(
            MAX_FEATURES_PER_PRODUCT,
            max(1, n_common - 2),
            int(use_idx.size),
        )
        if use_idx.size > max_k:
            sub = use_idx[np.argsort(col_var[use_idx])[-max_k:]]
        else:
            sub = use_idx

        Xmat = Xg[:, sub]
        ones = np.ones((n_common, 1), dtype=np.float64)
        Xw = np.hstack([ones, Xmat])
        pred_g, beta, r2 = _lstsq_predict_one_product(yg, Xw)
        r2_list.append((matnr, n_common, r2))

        for dt, yv, ph in zip(common, yg, pred_g):
            out_rows.append(
                {
                    "生价日期": dt,
                    "产品编码": matnr,
                    "BOM原材料总成本": float(yv),
                    "预测BOM总成本": float(ph),
                }
            )

        if beta is not None and sub.size <= 40 and n_common >= 4:
            names = [X_columns[int(j)] for j in sub]
            for name, b in zip(names, beta[1 : 1 + len(sub)]):
                coef_rows.append(
                    {
                        "产品编码": matnr,
                        "组件编码": name,
                        "回归系数(单价→总成本)": float(b),
                        "样本期数": n_common,
                    }
                )

    print(f"  回归进度 {n_products}/{n_products}（完成）", flush=True)
    final_df = pd.DataFrame(out_rows)
    if final_df.empty:
        raise RuntimeError("未得到任何产品×日期样本，请检查 BOM 数据。")
    final_df = final_df.sort_values(["产品编码", "生价日期"]).reset_index(drop=True)
    print(f"  产品×日期 样本行数 {len(final_df):,}", flush=True)

    r2_df = pd.DataFrame(r2_list, columns=["产品编码", "样本期数", "R2_样本内"])
    r2_df = r2_df.sort_values("R2_样本内", ascending=False)

    print("读取历史清单并合并 …", flush=True)
    hist = load_history(history_path)
    merged = final_df.merge(
        hist,
        left_on=["产品编码", "生价日期"],
        right_on=["MATNR", "CREATEDATE"],
        how="left",
    )
    merged_narrow = merged[
        [
            "生价日期",
            "产品编码",
            "BOM原材料总成本",
            "预测BOM总成本",
            "历史清单_DMBTR合计",
        ]
    ].copy()
    both = merged.dropna(subset=["BOM原材料总成本", "历史清单_DMBTR合计"])
    corr = float("nan")
    if len(both) > 10:
        a = both["BOM原材料总成本"].to_numpy()
        b = both["历史清单_DMBTR合计"].to_numpy()
        if np.std(a) > 1e-12 and np.std(b) > 1e-12:
            corr = float(np.corrcoef(a, b)[0, 1])

    # 材料映射覆盖：BOM 中 (产品,组件) 组合有多少落在映射表
    pair = bom[["产品编码", "组件编码"]].drop_duplicates()
    mm = pair.merge(
        mmap,
        left_on=["产品编码", "组件编码"],
        right_on=["ZMATNR", "IDNRK"],
        how="left",
        indicator=True,
    )
    map_cov = float((mm["_merge"] == "both").mean()) if len(mm) else float("nan")

    summary = pd.DataFrame(
        [
            {"项": "BOM 行数(原始)", "值": len(bom)},
            {"项": "材料映射行数(去重)", "值": len(mmap)},
            {"项": "BOM 产品-组件组合在映射表中的比例", "值": map_cov},
            {"项": "产品×日期 回归样本行数", "值": len(final_df)},
            {"项": "产品数", "值": final_df["产品编码"].nunique()},
            {"项": "R2 中位数(按产品)", "值": float(r2_df["R2_样本内"].median())},
            {"项": "R2 均值(按产品)", "值": float(r2_df["R2_样本内"].mean())},
            {"项": "与历史清单 DMBTR 合计(同键)相关系数", "值": corr},
            {"项": "历史键命中行数", "值": int(both.shape[0])},
            {"项": "使用材料映射过滤BOM行", "值": str(use_material_filter)},
        ]
    )

    coef_df = (
        pd.DataFrame(coef_rows)
        if coef_rows
        else pd.DataFrame(columns=["产品编码", "组件编码", "回归系数(单价→总成本)", "样本期数"])
    )

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="汇总", index=False)
        r2_df.head(5000).to_excel(writer, sheet_name="按产品R2_Top5000", index=False)
        merged_narrow.head(100000).to_excel(
            writer, sheet_name="明细_BOM与历史_十万行", index=False
        )
        if len(coef_df):
            coef_df.sort_values("回归系数(单价→总成本)", key=np.abs, ascending=False).head(
                20000
            ).to_excel(writer, sheet_name="系数节选", index=False)

    print(f"已写出：{out_xlsx}", flush=True)
    return out_xlsx


def main() -> None:
    p = argparse.ArgumentParser(description="三张表联合回归分析")
    p.add_argument(
        "--export",
        type=Path,
        default=Path(r"C:\Users\00109151\Desktop\EXPORT20260323.XLSX"),
    )
    p.add_argument(
        "--history",
        type=Path,
        default=Path(r"C:\Users\00109151\Desktop\数据集\25年至今产品历史清单.xlsx"),
    )
    p.add_argument(
        "--material",
        type=Path,
        default=Path(r"C:\Users\00109151\Desktop\数据集\材料所属产品.xlsx"),
    )
    p.add_argument(
        "--filter-by-material-map",
        action="store_true",
        help="仅保留 (产品编码,组件编码) 出现在材料映射表中的 BOM 行（默认关闭）",
    )
    args = p.parse_args()

    for path in (args.export, args.history, args.material):
        if not path.is_file():
            print(f"错误：找不到文件 {path.resolve()}", file=sys.stderr)
            sys.exit(1)

    out = run(
        args.export,
        args.history,
        args.material,
        use_material_filter=args.filter_by_material_map,
    )
    print(out)


if __name__ == "__main__":
    main()
