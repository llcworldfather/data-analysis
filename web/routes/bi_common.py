# -*- coding: utf-8 -*-
"""BI 路由共用：JSON 清洗、成本模拟、预测结果映射。"""
from __future__ import annotations

import math
from typing import Any

import pandas as pd

from process_excel import COL_PRODUCT, COL_CREATEDATE, COL_CATEGORY
from cost_sim_predict import (
    baseline_prices_from_price_snap,
    latest_bom_snapshot_from_rows,
    map_legacy_predict_en_to_zh,
    map_regression_analysis_en_to_zh,
    map_sensitivity_grid_en_to_zh,
    map_sensitivity_item_en_to_zh,
    predict_product_price,
    _predict_product_price_legacy,
)

_PRICE_EXTRA_COLS = ("材料成本", "工费", "重量", "报价月份", "标杆工厂")


def float_safe(x, default: float = 0.0) -> float:
    try:
        if x is None or (isinstance(x, str) and not str(x).strip()):
            return default
        f = float(x)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except (TypeError, ValueError):
        return default


def optional_rounded_float(x, *, ndigits: int = 6) -> float | None:
    if x is None or (isinstance(x, str) and not str(x).strip()):
        return None
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    if math.isnan(f) or math.isinf(f):
        return None
    return round(f, ndigits)


def json_clean_scalar(v: Any) -> Any:
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


def json_clean_row(row: dict) -> dict:
    return {k: json_clean_scalar(v) for k, v in row.items()}


def product_price_snapshot_from_bom_rows(rows: list[dict]) -> dict:
    if not rows:
        return {}

    def _row_sort_key(r: dict) -> tuple[str, str]:
        return (
            str(r.get("报价月份") or ""),
            str(r.get(COL_CREATEDATE) or ""),
        )

    r0 = max(rows, key=_row_sort_key)
    out: dict = {"价格口径": "元/KG"}
    if "产品价格" in r0:
        out["总成本"] = optional_rounded_float(r0.get("产品价格"))
    for c in _PRICE_EXTRA_COLS:
        if c in r0:
            out[c] = optional_rounded_float(r0.get(c))
    return out


def map_new_predict_result(
    pred: dict,
    price_snap: dict,
    *,
    is_bom_load: bool = False,
) -> dict:
    if "error" in pred:
        bl = baseline_prices_from_price_snap(price_snap, None)
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

    point_est = pred.get("point_estimate")
    point_kg = pred.get("point_per_kg")
    weight = pred.get("detail", {}).get("base_weight")
    try:
        weight_f = float(weight) if weight is not None else None
    except (TypeError, ValueError):
        weight_f = None

    bl = baseline_prices_from_price_snap(price_snap, weight_f)
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


def simulate_product_cost(cache: dict, product: str, prices_in: dict) -> dict:
    rows = (cache.get("product_bom") or {}).get(product)
    if not rows:
        raise ValueError(f"未找到产品：{product}")

    prices = {str(k).strip(): float_safe(v) for k, v in prices_in.items()}
    prices_old: dict[str, float] = {}

    baseline = 0.0
    simulated = 0.0
    line_out: list[dict] = []
    for r in rows:
        comp = str(r.get("组件编码") or "").strip()
        m = float_safe(r.get("MENGE合计"))
        old_p = float_safe(r.get("组件单价"))
        prices_old[comp] = old_p
        old_line = float_safe(r.get("行原材料成本"), m * old_p)
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
        nl = float_safe(row.get("模拟行成本"))
        row["成本占产品%"] = (
            round(100.0 * nl / simulated, 4) if simulated > 1e-12 else 0.0
        )

    price_snap = product_price_snapshot_from_bom_rows(rows)
    bom_pd = cache.get("bom_pd")
    price_pd = cache.get("price_pd")
    if bom_pd is not None and price_pd is not None:
        all_prices = {c: prices.get(c, prices_old[c]) for c in prices_old}
        user_changes = {
            c: all_prices[c]
            for c in prices_old
            if abs(all_prices[c] - prices_old[c]) > 1e-9
        }
        bom_snap = latest_bom_snapshot_from_rows(product, rows)
        pred = predict_product_price(
            product,
            user_changes,
            bom_pd,
            price_pd,
            reference_prices=prices_old,
            latest_bom_snapshot=bom_snap,
        )
        pred = map_new_predict_result(pred, price_snap)
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
    return {
        "product": product,
        "baseline原材料总成本": round(baseline, 6),
        "模拟原材料总成本": sim_round,
        "差额": round(simulated - baseline, 6),
        "总成本": ref_total,
        **{k: v for k, v in price_snap.items() if k != "总成本"},
        "差额_相对总成本": (
            round(sim_round - ref_total, 6) if ref_total is not None else None
        ),
        "lines": [json_clean_row(x) for x in line_out],
        **pred,
    }
