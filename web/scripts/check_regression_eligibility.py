# -*- coding: utf-8 -*-
"""分析三表 Excel 中哪些产品满足成本模拟「历史回归」条件。"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import polars as pl

ROOT = Path(__file__).resolve().parent.parent.parent
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

DEFAULT_DATA = ROOT / "data"
OUT_TXT = ROOT / "output" / "regression_eligibility_report.txt"
OUT_MISSING_XLSX = ROOT / "output" / "missing_bom_for_price_products.xlsx"


def _norm_codes(s: pd.Series) -> pd.Series:
    return norm_material_code(s)


def _prefix_overlap(
    bom_codes: set[str],
    price_codes: set[str],
    lengths: range | None = None,
) -> list[tuple[int, int, list[tuple[str, str]]]]:
    """扫描前缀长度，返回 (L, 匹配对数, 样例对照)。"""
    if lengths is None:
        lengths = range(5, 13)
    results: list[tuple[int, int, list[tuple[str, str]]]] = []
    for n in lengths:
        bom_pref: dict[str, str] = {}
        for c in bom_codes:
            base = c.split("-")[0].strip()
            if len(base) >= n:
                bom_pref[base[:n]] = c
        pairs: list[tuple[str, str]] = []
        for pc in price_codes:
            base = pc.split("-")[0].strip()
            key = base[:n] if len(base) >= n else base
            if key in bom_pref:
                pairs.append((bom_pref[key], pc))
        results.append((n, len(pairs), pairs[:10]))
    return results


def _resolve_data_paths(data_dir: Path) -> tuple[Path, Path]:
    bom_candidates = [
        data_dir / "报价BOM历史清单.xls",
        data_dir / "报价BOM历史清单.xlsx",
    ]
    price_candidates = [
        data_dir / "产品价格历史清单.xls",
        data_dir / "产品价格历史清单.xlsx",
    ]
    bom_path = next((p for p in bom_candidates if p.exists()), bom_candidates[0])
    price_path = next((p for p in price_candidates if p.exists()), price_candidates[0])
    return bom_path, price_path


def main() -> None:
    parser = argparse.ArgumentParser(description="BOM 与价格清单回归条件诊断")
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=DEFAULT_DATA,
        help=f"数据目录（默认 {DEFAULT_DATA}）",
    )
    args = parser.parse_args()
    data_dir: Path = args.data_dir
    bom_path, price_path = _resolve_data_paths(data_dir)

    if not bom_path.exists():
        print(f"未找到 BOM 文件: {bom_path}")
        sys.exit(1)
    if not price_path.exists():
        print(f"未找到价格文件: {price_path}")
        sys.exit(1)

    bom_xl = pd.read_excel(bom_path)
    price_xl = pd.read_excel(price_path)

    bom_product_col = "所属产品" if "所属产品" in bom_xl.columns else "产品编码"
    bom_codes = set(_norm_codes(bom_xl[bom_product_col])) - {"", "nan"}
    price_codes = set(_norm_codes(price_xl["产品编码"])) - {"", "nan"}
    overlap = bom_codes & price_codes

    lines: list[str] = []
    lines.append("=" * 60)
    lines.append("BOM + 产品价格历史回归条件分析报告")
    lines.append("=" * 60)
    lines.append(f"数据目录: {data_dir.resolve()}")
    lines.append(f"BOM 文件: {bom_path.name}")
    lines.append(f"价格文件: {price_path.name}")
    lines.append("")
    lines.append("【回归启用条件（与页面一致）】")
    lines.append("  1. 产品编码在 BOM 与 产品价格历史清单 中能对齐")
    lines.append("  2. 至少 3 期「有效样本」：每期同时有")
    lines.append("     - BOM材料(元/kg) = 整件材料合计 ÷ 重量(kg)")
    lines.append("     - 产品价(元/kg) 来自价格清单")
    lines.append("  3. 每期必须有 重量(kg)（价格清单「重量」列）")
    lines.append("")
    lines.append("【两表规模】")
    lines.append(f"  报价BOM历史清单: {len(bom_xl)} 行, {len(bom_codes)} 个产品({bom_product_col})")
    lines.append(f"  产品价格历史清单: {len(price_xl)} 行, {len(price_codes)} 个产品")
    lines.append("")
    lines.append("【编码对齐情况】")
    lines.append(f"  BOM 产品 ∩ 价格清单产品 = {len(overlap)} 个")
    if overlap:
        lines.append(f"    列表: {', '.join(sorted(overlap)[:30])}")
    else:
        lines.append("    *** 当前数据集中为 0，无法对任何产品做完整回归 ***")
        lines.append(f"    BOM 示例: {', '.join(sorted(bom_codes)[:5])}")
        lines.append(f"    价格表示例: {', '.join(sorted(price_codes)[:5])}")

    lines.append("")
    lines.append("【前缀匹配诊断（仅建议，不自动用于线上预测）】")
    prefix_rows = _prefix_overlap(bom_codes, price_codes)
    for n, cnt, samples in prefix_rows:
        lines.append(f"  前 {n} 位一致（含 split('-')[0] 后截取）: {cnt} 对")
        if samples and cnt > 0 and cnt <= 30:
            for b, p in samples[:5]:
                lines.append(f"    BOM {b}  <->  价格 {p}")
    best = max(prefix_rows, key=lambda x: x[1]) if prefix_rows else (0, 0, [])
    if best[1] > len(overlap):
        lines.append(
            f"  提示: 完全匹配 {len(overlap)} 个；若采用前 {best[0]} 位前缀规则可对齐约 {best[1]} 个，"
            "需业务确认后维护对照表。"
        )

    lines.append("")

    lines.append("-" * 60)
    lines.append("【仅看产品价格历史清单】每期有重量 → 可作「产品价/kg」侧")
    lines.append("-" * 60)
    px = price_xl.copy()
    px["产品编码"] = _norm_codes(px["产品编码"])
    px["重量"] = pd.to_numeric(px["重量"], errors="coerce")
    if "报价月份" in px.columns:
        px["_期"] = px["报价月份"].astype(str).str.replace(r"\.0+$", "", regex=True)
    else:
        px["_期"] = px.get("创建日期", "").astype(str)

    price_stats = []
    for pid, grp in px.groupby("产品编码"):
        if not pid or pid == "nan":
            continue
        with_wt = grp[grp["重量"].notna() & (grp["重量"] > 0)]
        periods = with_wt["_期"].nunique() if len(with_wt) else 0
        price_stats.append(
            {
                "产品编码": pid,
                "价格表行数": len(grp),
                "有重量期数": periods,
                "清单够3期": periods >= 3,
                "在BOM中": pid in bom_codes,
            }
        )
    pdf = pd.DataFrame(price_stats).sort_values("有重量期数", ascending=False)
    ok_price = pdf[pdf["清单够3期"]]
    lines.append(f"  价格清单内 ≥3 期(有重量) 的产品: {len(ok_price)} 个")
    if len(ok_price):
        for _, r in ok_price.head(20).iterrows():
            in_bom = "是" if r["在BOM中"] else "否"
            lines.append(
                f"    {r['产品编码']}: {int(r['有重量期数'])} 期, BOM中有该产品={in_bom}"
            )

    missing_bom = ok_price[~ok_price["在BOM中"]].head(20)
    if len(missing_bom):
        OUT_MISSING_XLSX.parent.mkdir(parents=True, exist_ok=True)
        export_cols = ["产品编码", "有重量期数", "价格表行数", "清单够3期"]
        missing_bom[export_cols].to_excel(OUT_MISSING_XLSX, index=False)
        lines.append("")
        lines.append(
            f"  已导出「价格表够3期但 BOM 缺失」Top {len(missing_bom)} 产品: {OUT_MISSING_XLSX}"
        )

    lines.append("")

    lines.append("-" * 60)
    lines.append("【Web 流水线：BOM + 价格清单 → 有效回归样本】")
    lines.append("-" * 60)
    bom_raw = pl.from_pandas(bom_xl)
    pch = build_product_cost_history(bom_raw)
    price_tl = build_product_price_timeline(price_path)

    pbom: dict[str, list[dict]] = defaultdict(list)
    prep = normalize_columns(pl.from_pandas(bom_xl))
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

    expanded = expand_product_cost_history_dict(dict(pch), dict(pbom), None, price_tl)

    reg_rows = []
    for pid in sorted(bom_codes | price_codes):
        hist = expanded.get(pid, pch.get(pid, []))
        valid = []
        for h in hist:
            bm, pp = history_per_kg_pair(h)
            if bm is not None and pp is not None and bm > 1e-12 and pp > 0:
                valid.append(h)
        n = len(valid)
        r2 = 0.0
        if n >= 3:
            x = np.array([history_per_kg_pair(h)[0] for h in valid], dtype=float)
            y = np.array([history_per_kg_pair(h)[1] for h in valid], dtype=float)
            _, r2, _, _, _ = _linear_fit_predict(x, y, float(np.mean(x)))
        reg_rows.append(
            {
                "产品编码": pid,
                "有效样本期数": n,
                "满足回归": n >= 3,
                "R2": round(r2, 3) if n >= 3 else None,
                "拟合等级": _regression_fit_grade(r2, n),
                "在BOM": pid in bom_codes,
                "在价格表": pid in price_codes,
            }
        )
    rdf = pd.DataFrame(reg_rows)
    eligible = rdf[rdf["满足回归"]]
    lines.append(f"  满足完整回归条件的产品: {len(eligible)} 个")
    if len(eligible):
        lines.append(eligible.to_string(index=False))
    else:
        lines.append("  (无)")
    lines.append("")
    lines.append("【结论与建议】")
    if len(overlap) == 0:
        lines.append(
            "  根因: BOM(所属产品) 与 价格清单(产品编码) 无交集，"
            "系统无法为任何产品同时构造「BOM材料/kg」与「产品价/kg」配对样本。"
        )
        if best[1] > 0:
            lines.append(
                f"  前缀诊断: 前 {best[0]} 位规则可潜在对齐 {best[1]} 对，"
                "请导出对照表经业务确认后再接入预测。"
            )
        lines.append(
            "  建议: ① 确认两表是否应对同一批成品物料号；"
            "② 若业务上为不同编码体系，需维护对照表或统一导出字段；"
            "③ 仅有价格表产品时需提供对应 BOM（见 missing_bom Excel）。"
        )
    elif len(eligible) == 0:
        lines.append("  有产品交集但有效样本仍不足 3 期，请补充 BOM 或产品价格历史月份。")
    else:
        lines.append(f"  可直接做回归的产品见上表（共 {len(eligible)} 个）。")

    OUT_TXT.parent.mkdir(parents=True, exist_ok=True)
    text = "\n".join(lines)
    OUT_TXT.write_text(text, encoding="utf-8")
    print(text)
    print(f"\n已写入: {OUT_TXT}")


if __name__ == "__main__":
    main()
