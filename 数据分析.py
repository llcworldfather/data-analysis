from typing import List, Optional, Tuple

import pandas as pd
import numpy as np
import statsmodels.api as sm
import matplotlib.pyplot as plt

from bom_date_key import calendar_date_key_yyyymmdd

# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 每产品最多用多少个「单价有波动」的列。不是越大越好：行数很少时特征多了会过拟合/不可辨识；
# 取方差最大的若干列是在「信息量」和「速度/稳定性」之间的折中。
MAX_FEATURES_PER_PRODUCT = 60


def _pick_zmatnr_for_plot(final_df: pd.DataFrame, k: int = 3) -> np.ndarray:
    """
    避免用 final_df['ZMATNR'].unique()[:3]：前几个产品若成本曲线完全一致，6 条线会两两重合，
    图上就像只剩「一条实线 + 一条虚线」。这里优先选平均成本差异大的产品便于区分。
    """
    agg = final_df.groupby("ZMATNR", sort=False)["Total_Cost"].agg(
        count="count", mean="mean", std="std"
    )
    agg = agg[agg["count"] >= 2]
    if len(agg) == 0:
        return final_df["ZMATNR"].unique()[:k]
    mats = agg.index.to_numpy()
    if len(mats) <= k:
        return mats

    i_hi = agg["mean"].idxmax()
    i_lo = agg["mean"].idxmin()
    picked = [i_hi, i_lo]
    rest = [m for m in mats if m not in picked]
    if rest:
        med = float(agg["mean"].median())
        i_third = max(rest, key=lambda m: abs(float(agg.loc[m, "mean"]) - med))
        picked.append(i_third)

    out = []
    for p in picked:
        if p not in out:
            out.append(p)
    # 若仍有重复或不足 k 个，按成本波动 std 补满
    by_std = agg.sort_values("std", ascending=False).index.tolist()
    for m in by_std:
        if len(out) >= k:
            break
        if m not in out:
            out.append(m)
    return np.array(out[:k], dtype=object)


def _lstsq_predict_one_product(
    y: np.ndarray,
    X: np.ndarray,
) -> Tuple[np.ndarray, Optional[np.ndarray]]:
    """
    y: (n,), X: (n, p) 已含常数列在第一列。
    返回 (全样本预测, beta) beta 长度为 p；失败时预测为有效样本的均值，beta 为 None。
    """
    n, p = X.shape
    valid = np.isfinite(X).all(axis=1) & np.isfinite(y)
    nv = int(valid.sum())
    fill = float(np.nanmean(y[valid])) if nv else np.nan
    pred = np.full(n, fill, dtype=float)

    if nv < p + 1:
        return pred, None

    Xv, yv = X[valid], y[valid]
    try:
        beta, _, _, _ = np.linalg.lstsq(Xv, yv, rcond=None)
        pred = X @ beta
        return pred.astype(float, copy=False), beta
    except Exception:
        return pred, None


def process_price_analysis(file_path):
    # 1. 加载数据
    print("1/5 正在读取 Excel（行数很大时可能要 1～3 分钟）...", flush=True)
    df = pd.read_excel(file_path)
    # 与 EXPORT「生价日期」、历史表 CREATEDATE 同一套 YYYYMMDD 规范化后再转 datetime
    df["CREATEDATE"] = pd.to_datetime(
        df["CREATEDATE"].map(calendar_date_key_yyyymmdd).replace("", pd.NA),
        format="%Y%m%d",
        errors="coerce",
    )

    # 2. 透视：组件单价（列 = MAKTX）
    print("2/5 正在生成透视表...", flush=True)
    pivot_df = df.pivot_table(
        index=["CREATEDATE", "ZMATNR"], columns="MAKTX", values="ZDJ"
    ).reset_index()

    pivot_df = pivot_df.sort_values(["ZMATNR", "CREATEDATE"]).reset_index(drop=True)
    price_cols = [c for c in pivot_df.columns if c not in ("CREATEDATE", "ZMATNR")]
    print(
        f"3/5 按产品填充缺失单价（{len(price_cols)} 列 × {len(pivot_df)} 行，使用向量化 groupby 填充）...",
        flush=True,
    )
    if price_cols:
        pivot_df[price_cols] = pivot_df.groupby("ZMATNR", sort=False)[price_cols].ffill()
        pivot_df[price_cols] = pivot_df.groupby("ZMATNR", sort=False)[price_cols].bfill()

    pivot_df = pivot_df.reset_index(drop=True)

    df = df.assign(_line=df["ZDJ"] * df["MENGE"])
    costs = (
        df.groupby(["CREATEDATE", "ZMATNR"], as_index=False)["_line"]
        .sum()
        .rename(columns={"_line": "Total_Cost"})
    )

    final_df = pd.merge(pivot_df, costs, on=["CREATEDATE", "ZMATNR"], how="inner")
    final_df = final_df.drop_duplicates(subset=["CREATEDATE", "ZMATNR"], keep="last")
    final_df = final_df.replace([np.inf, -np.inf], np.nan)
    print(f"   合并去重后 {len(final_df)} 行（透视宽表允许大量单价为空）", flush=True)

    X_columns = [c for c in pivot_df.columns if c not in ("CREATEDATE", "ZMATNR")]
    final_df = final_df.dropna(subset=["CREATEDATE", "ZMATNR", "Total_Cost"])
    final_df = final_df.reset_index(drop=True)

    n_products = int(final_df["ZMATNR"].nunique())
    print(
        f"4/5 按产品做回归（产品数 {n_products}，行数 {len(final_df)}；"
        f"每产品最多 {MAX_FEATURES_PER_PRODUCT} 个特征，用 numpy 最小二乘）...",
        flush=True,
    )

    if len(final_df) < 10:
        print("警告：有效数据量太少，无法进行回归分析。")
        return None, None

    # 按产品排序后一次性拉成 (n, n_features) 的 C 连续矩阵，用组下标循环；
    # 避免 3 万次 pandas groupby + 每次切 1509 列（那一步会极慢）。
    print("   正在构建特征矩阵并排序（一次性，稍等）...", flush=True)
    order = final_df["ZMATNR"].to_numpy().argsort(kind="mergesort")
    final_sorted = final_df.iloc[order].reset_index(drop=True)
    X_s = final_sorted[X_columns].to_numpy(dtype=np.float64, copy=True)
    y_s = final_sorted["Total_Cost"].to_numpy(dtype=np.float64, copy=True)
    z = final_sorted["ZMATNR"].to_numpy(copy=False)

    boundaries = np.flatnonzero(z[1:] != z[:-1]) + 1
    starts = np.r_[0, boundaries]
    ends = np.r_[boundaries, len(z)]
    n_groups = len(starts)

    pred_s = np.empty_like(y_s)
    coef_tables: List[pd.DataFrame] = []
    max_coef_products = 2000
    coef_products_added = 0

    for gi in range(n_groups):
        if gi % 5000 == 0:
            print(f"   回归进度: {gi}/{n_groups} 个产品", flush=True)

        s, e = int(starts[gi]), int(ends[gi])
        Xg = X_s[s:e]
        yg = y_s[s:e]
        matnr = z[s]

        col_var = np.nanvar(Xg, axis=0)
        use_idx = np.flatnonzero(np.isfinite(col_var) & (col_var > 1e-14))
        if use_idx.size == 0:
            pred_s[s:e] = float(np.nanmean(yg))
            continue

        max_k = min(
            MAX_FEATURES_PER_PRODUCT,
            max(1, (e - s) - 2),
            int(use_idx.size),
        )
        if use_idx.size > max_k:
            sub = use_idx[np.argsort(col_var[use_idx])[-max_k:]]
        else:
            sub = use_idx

        Xmat = Xg[:, sub]
        ones = np.ones((e - s, 1), dtype=np.float64)
        Xw = np.hstack([ones, Xmat])

        pred_g, beta = _lstsq_predict_one_product(yg, Xw)
        pred_s[s:e] = pred_g

        if (
            beta is not None
            and coef_products_added < max_coef_products
            and sub.size <= 80
        ):
            names = [X_columns[int(j)] for j in sub]
            coef_tables.append(
                pd.DataFrame(
                    {
                        "产品编码": matnr,
                        "组件名称": names,
                        "影响系数": beta[1 : 1 + len(sub)],
                        "P值(显著性)": np.nan,
                    }
                )
            )
            coef_products_added += 1

    print(f"   回归进度: {n_groups}/{n_groups} 个产品（完成）", flush=True)

    pred_all = np.empty(len(final_df), dtype=np.float64)
    pred_all[order] = pred_s
    final_df["Predicted_Cost"] = pred_all

    print("5/5 绘图并汇总系数...", flush=True)
    coef_df = (
        pd.concat(coef_tables, ignore_index=True)
        if coef_tables
        else pd.DataFrame(
            columns=["产品编码", "组件名称", "影响系数", "P值(显著性)"]
        )
    )
    if len(coef_df):
        coef_df = coef_df.sort_values("影响系数", ascending=False)

    try:
        plt.figure(figsize=(12, 7))
        unique_mats = _pick_zmatnr_for_plot(final_df, k=3)
        for matnr in unique_mats:
            sub_df = final_df[final_df["ZMATNR"] == matnr].sort_values("CREATEDATE")
            plt.plot(
                sub_df["CREATEDATE"],
                sub_df["Total_Cost"],
                marker="o",
                label=f"产品:{matnr}-实际",
            )
            plt.plot(
                sub_df["CREATEDATE"],
                sub_df["Predicted_Cost"],
                linestyle="--",
                label=f"产品:{matnr}-预测",
            )

        plt.title(
            "组件价格对产品成本影响的拟合图\n"
            f"（示例产品: {', '.join(str(m) for m in unique_mats)}）"
        )
        plt.xlabel("日期")
        plt.ylabel("金额")
        plt.legend()
        plt.grid(True, linestyle=":", alpha=0.6)
        plt.tight_layout()
        plt.savefig("price_analysis_plot.png")
    except Exception as e:
        print(f"绘图出错: {e}")

    global_model = None
    if len(X_columns) <= 120:
        try:
            X_all = sm.add_constant(final_df[X_columns].fillna(0))
            y_all = final_df["Total_Cost"]
            mask = X_all.notna().all(axis=1) & y_all.notna()
            global_model = sm.OLS(y_all[mask], X_all[mask], missing="drop").fit()
        except Exception as e:
            print(f"全样本参考回归未拟合（可忽略）: {e}")
    else:
        print(
            f"   全样本参考回归已跳过（组件列数 {len(X_columns)} 过大，避免长时间占用）",
            flush=True,
        )

    return global_model, coef_df


if __name__ == "__main__":
    print("数据分析开始（请稍候，窗口会持续刷新进度）\n", flush=True)
    try:
        model, sensitivity_table = process_price_analysis("data.xlsx")
        if sensitivity_table is not None:
            print("-" * 30)
            print("✅ 分析成功完成！")
            print("1. 拟合图表已更新: price_analysis_plot.png")
            print("2. 核心组件影响权重（按产品拆分，节选前5行；P值列已省略精确计算）")
            print(sensitivity_table.head(5))
            print("-" * 30)
    except Exception as e:
        print(f"❌ 程序发生崩溃: {e}")

    input("\n按回车键退出...")
