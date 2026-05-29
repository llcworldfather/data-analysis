# -*- coding: utf-8 -*-
"""BI 看板、成本模拟与 /api/bi/* 路由。"""
from __future__ import annotations

import traceback
from typing import Callable

import pandas as pd
from flask import Blueprint, jsonify, render_template, request

from process_excel import COL_PRODUCT, COL_COMPONENT, COL_CREATEDATE, COL_CATEGORY
from services import bi_cache
from routes.bi_common import (
    float_safe,
    json_clean_row,
    map_new_predict_result,
    product_price_snapshot_from_bom_rows,
    simulate_product_cost,
)
from cost_sim_predict import (
    latest_bom_snapshot_from_rows,
    map_legacy_predict_en_to_zh,
    predict_product_price,
    _predict_product_price_legacy,
)

bp = Blueprint("bi", __name__)

_external_capabilities_fn: Callable[[], dict] | None = None


def init_blueprint(external_capabilities: Callable[[], dict]) -> None:
    global _external_capabilities_fn
    _external_capabilities_fn = external_capabilities


def session_capabilities(session_id: str) -> dict:
    if _external_capabilities_fn is None:
        raise RuntimeError("bi.init_blueprint 未调用")
    out = _external_capabilities_fn()
    cache = bi_cache.peek(session_id)
    has_price = bi_cache.session_has_price_history(cache)
    predict_reason = ""
    if not has_price:
        predict_reason = (
            "本次分析未附带「产品价格历史清单」，"
            "「计算预测产品价」等高级预测可能不可用；BOM 材料合计模拟仍可使用。"
            "可在首页重新上传 BOM 并附加价格表。"
        )
    out["cost_sim"] = {
        "enabled": True,
        "predict_enabled": has_price,
        "reason": predict_reason,
    }
    return out


def _expired_html() -> tuple[str, int]:
    return "会话已过期或不存在，请返回首页重新上传并执行分析。", 404


@bp.route("/bi/<session_id>")
def bi_view(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    if not bi_cache.session_ready(session_id):
        return _expired_html()
    try:
        bi_cache.get(session_id)
    except FileNotFoundError:
        return _expired_html()
    caps = session_capabilities(session_id)
    return render_template("bi.html", token=session_id, capabilities=caps)


@bp.route("/sim/<session_id>")
def cost_sim_view(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    if not bi_cache.session_ready(session_id):
        return _expired_html()
    try:
        bi_cache.get(session_id)
    except FileNotFoundError:
        return _expired_html()
    caps = session_capabilities(session_id)
    return render_template("cost_sim.html", token=session_id, capabilities=caps)


@bp.route("/api/bi_ready/<session_id>")
def bi_ready(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"ready": False}), 400
    return jsonify({"ready": bi_cache.session_ready(session_id)})


@bp.route("/api/bi/capabilities/<session_id>")
def bi_capabilities(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    if not bi_cache.session_ready(session_id):
        return jsonify({"error": "会话不存在或已过期，请重新分析"}), 404
    return jsonify(session_capabilities(session_id))


@bp.route("/api/bi/products/<session_id>")
def bi_products(session_id: str):
    FULL_LIST_THRESHOLD = 2000
    try:
        limit = int(request.args.get("limit", 100))
    except (TypeError, ValueError):
        limit = 100
    limit = max(1, min(limit, 500))

    cache, _, err = bi_cache.cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None

    all_ids: list[str] = cache.get("all_product_ids") or sorted(
        (cache.get("product_bom") or {}).keys(), key=lambda s: (len(s), s)
    )
    prefix_index: dict[str, list[str]] = cache.get("product_prefix_index") or {}
    total = len(all_ids)
    q = (request.args.get("q") or "").strip()
    min_q_len = 3 if total > 200_000 else 2

    def _search_ids(q: str, ids: list[str], idx: dict[str, list[str]], lim: int) -> list[str]:
        if not q:
            return ids[:lim]
        if len(q) >= 2 and idx:
            candidates = idx.get(q[:2], [])
            matched = [p for p in candidates if q in p]
            if len(matched) < lim:
                seen = set(candidates)
                for p in ids:
                    if p not in seen and q in p:
                        matched.append(p)
                    if len(matched) >= lim:
                        break
        else:
            matched = [p for p in ids if q in p]
        return matched[:lim]

    if total <= FULL_LIST_THRESHOLD:
        items = _search_ids(q, all_ids, prefix_index, limit) if q else all_ids
        return jsonify(
            {
                "count": total,
                "total": total,
                "search_required": False,
                "min_q_len": min_q_len,
                "products": [{"id": p, "label": p} for p in (items if q else all_ids)],
            }
        )

    if not q:
        return jsonify(
            {
                "count": total,
                "total": total,
                "search_required": True,
                "min_q_len": min_q_len,
                "products": [],
                "hint": f"共 {total} 个产品，请在输入框中输入至少 {min_q_len} 位编码片段进行搜索",
            }
        )

    if len(q) < min_q_len:
        return jsonify(
            {
                "count": total,
                "total": total,
                "search_required": True,
                "min_q_len": min_q_len,
                "products": [],
                "hint": f"请至少输入 {min_q_len} 位再搜索（当前 {len(q)} 位）",
            }
        )

    matched = _search_ids(q, all_ids, prefix_index, limit)
    return jsonify(
        {
            "count": total,
            "total": total,
            "search_required": True,
            "min_q_len": min_q_len,
            "q": q,
            "match_count": len(matched),
            "truncated": len(matched) >= limit,
            "products": [{"id": x, "label": x} for x in matched],
        }
    )


@bp.route("/api/bi/product_bom/<session_id>")
def bi_product_bom(session_id: str):
    product = (request.args.get("product") or "").strip()
    if not product:
        return jsonify({"error": "缺少参数 product（产品编码）"}), 400
    cache, _, err = bi_cache.cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None
    rows = (cache.get("product_bom") or {}).get(product)
    if not rows:
        return jsonify({"error": f"未找到产品：{product}"}), 404

    lines: list[dict] = []
    baseline = 0.0
    for r in rows:
        row = dict(r)
        m = float_safe(row.get("MENGE合计"))
        p0 = float_safe(row.get("组件单价"))
        line_cost = float_safe(row.get("行原材料成本"), m * p0)
        baseline += line_cost
        row["行原材料成本"] = line_cost
        row["单价每变动1对总成本的边际"] = m
        lines.append(row)
    if baseline > 0:
        for row in lines:
            lc = float_safe(row.get("行原材料成本"))
            row["成本占产品%"] = round(100.0 * lc / baseline, 4)
    lines.sort(key=lambda x: float_safe(x.get("行原材料成本")), reverse=True)

    price_snap = product_price_snapshot_from_bom_rows(rows)
    product_price = price_snap.get("总成本")
    prices_old = {
        str(r.get("组件编码") or "").strip(): float_safe(r.get("组件单价"))
        for r in rows
    }

    bom_pd = cache.get("bom_pd")
    price_pd = cache.get("price_pd")
    if bom_pd is not None and price_pd is not None:
        bom_snap = latest_bom_snapshot_from_rows(product, rows)
        pred = predict_product_price(
            product,
            {},
            bom_pd,
            price_pd,
            reference_prices=prices_old,
            latest_bom_snapshot=bom_snap,
        )
        pred = map_new_predict_result(pred, price_snap, is_bom_load=True)
    else:
        pred = map_legacy_predict_en_to_zh(
            _predict_product_price_legacy(
                product=product,
                simulated_material=baseline,
                baseline_material=baseline,
                price_snap=price_snap,
                product_cost_history=cache.get("product_cost_history"),
                price_history=cache.get("price_history"),
                bom_rows=rows,
                prices_new=prices_old,
                prices_old=prices_old,
                product_categories=cache.get("product_categories"),
            )
        )

    return jsonify(
        {
            "product": product,
            "baseline原材料总成本": round(baseline, 6),
            "产品价格": product_price,
            "总成本": product_price,
            **{k: v for k, v in price_snap.items() if k != "总成本"},
            **pred,
            "说明": (
                "主结果「预测产品价格」由历史报价关系与成本结构（材料成本+工费）推算，不是简单 BOM 行相加。"
                "「模拟原材料总成本」为各组件数量×模拟单价之和，供对照。"
            ),
            "lines": [json_clean_row(row) for row in lines],
        }
    )


@bp.route("/api/bi/product_cost_simulate/<session_id>", methods=["POST"])
def bi_product_cost_simulate(session_id: str):
    cache, _, err = bi_cache.cache_or_404(session_id)
    if err:
        return err[0], err[1]
    assert cache is not None
    data = request.get_json(silent=True) or {}
    product = str(data.get("product") or "").strip()
    prices_in = data.get("prices") or {}
    if not product:
        return jsonify({"error": "缺少 product"}), 400
    try:
        payload = simulate_product_cost(cache, product, prices_in)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    return jsonify(payload)


@bp.route("/api/bi/summary/<session_id>")
def bi_summary(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    if not bi_cache.session_ready(session_id):
        return jsonify({"error": "会话不存在或已过期，请重新分析"}), 404
    try:
        cache = bi_cache.get(session_id)
        return jsonify({"data": cache["summary"]})
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}), 500


@bp.route("/api/bi/detail/<session_id>")
def bi_detail(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400
    component = request.args.get("component", "").strip()
    if not component:
        return jsonify({"error": "缺少参数 component"}), 400
    if not bi_cache.session_ready(session_id):
        return jsonify({"error": "会话不存在或已过期，请重新分析"}), 404
    try:
        cache = bi_cache.get(session_id)
        records = cache["detail"].get(component, [])
        return jsonify({
            "data": records,
            "component": component,
            "count": len(records),
        })
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}), 500


@bp.route("/api/bi/price_fluctuation/<session_id>")
def bi_price_fluctuation(session_id: str):
    if not bi_cache.uuid_valid(session_id):
        return jsonify({"error": "无效的令牌"}), 400

    start_date = (request.args.get("start_date") or "").strip().replace("-", "")
    end_date = (request.args.get("end_date") or "").strip().replace("-", "")
    min_pct = abs(float(request.args.get("min_change_pct", 0) or 0))

    if not start_date or not end_date:
        return jsonify({"error": "请提供 start_date 和 end_date（格式：YYYYMMDD 或 YYYY-MM-DD）"}), 400

    if not bi_cache.session_ready(session_id):
        return jsonify({"error": "会话不存在或已过期，请重新分析"}), 404

    try:
        from bom_date_key import calendar_date_key_yyyymmdd

        cache = bi_cache.get(session_id)
        price_history = cache.get("price_history", [])

        if not price_history:
            return jsonify({
                "error": "无价格历史数据，请重新上传并执行分析",
                "data": [], "category_impact": [], "has_category": False,
            }), 200

        df = pd.DataFrame(price_history)
        if COL_CREATEDATE in df.columns:
            df[COL_CREATEDATE] = df[COL_CREATEDATE].map(calendar_date_key_yyyymmdd)
            df_range = df[(df[COL_CREATEDATE] >= start_date) & (df[COL_CREATEDATE] <= end_date)].copy()
        else:
            df_range = df.copy()

        if df_range.empty:
            return jsonify({
                "data": [], "category_impact": [], "has_category": False,
                "message": f"所选日期范围（{start_date}–{end_date}）内无数据",
                "total": 0,
            })

        comp_col = COL_COMPONENT
        price_col = "组件单价"
        has_cat = bool(COL_CATEGORY in df_range.columns and df_range[COL_CATEGORY].notna().any())
        name_col = "组件名称" if "组件名称" in df_range.columns else ("MAKTX" if "MAKTX" in df_range.columns else None)

        results = []
        for comp, grp in df_range.groupby(comp_col, sort=False):
            if COL_CREATEDATE in grp.columns:
                grp = grp.sort_values(COL_CREATEDATE)
            price_vals = grp[price_col].dropna().astype(float)
            if price_vals.empty:
                continue
            price_start = float(price_vals.iloc[0])
            price_end = float(price_vals.iloc[-1])
            price_min = float(price_vals.min())
            price_max = float(price_vals.max())
            price_diff = price_end - price_start
            price_pct = round((price_diff / price_start * 100), 2) if price_start != 0 else 0.0

            row = {
                "组件编码": str(comp),
                "期初价格": round(price_start, 4),
                "期末价格": round(price_end, 4),
                "期间最低价": round(price_min, 4),
                "期间最高价": round(price_max, 4),
                "价格差值": round(price_diff, 4),
                "价格变动%": price_pct,
                "记录数": int(len(grp)),
                "涉及产品数": int(grp[COL_PRODUCT].nunique()) if COL_PRODUCT in grp.columns else 0,
            }
            if name_col:
                row["组件名称"] = str(grp[name_col].dropna().iloc[0]) if grp[name_col].notna().any() else ""
            if has_cat:
                row[COL_CATEGORY] = str(grp[COL_CATEGORY].dropna().iloc[0]) if grp[COL_CATEGORY].notna().any() else ""
            results.append(row)

        if min_pct > 0:
            results = [r for r in results if abs(r["价格变动%"]) >= min_pct]
        results.sort(key=lambda r: r["价格差值"], reverse=True)

        category_impact = []
        if has_cat and results:
            cat_df = pd.DataFrame(results)
            cat_df = cat_df[cat_df[COL_CATEGORY].notna() & (cat_df[COL_CATEGORY] != "")]
            if not cat_df.empty:
                for cat, cgrp in cat_df.groupby(COL_CATEGORY, sort=False):
                    pcts = cgrp["价格变动%"]
                    abs_max = float(pcts.abs().max())
                    level = "高" if abs_max >= 10 else ("中" if abs_max >= 3 else "低")
                    category_impact.append({
                        "分类": str(cat),
                        "平均变动%": round(float(pcts.mean()), 2),
                        "最大涨幅%": round(float(pcts.max()), 2),
                        "最大跌幅%": round(float(pcts.min()), 2),
                        "最大绝对变动%": round(abs_max, 2),
                        "涉及组件数": int(len(cgrp)),
                        "上涨组件数": int((pcts > 0).sum()),
                        "下跌组件数": int((pcts < 0).sum()),
                        "影响等级": level,
                    })
                category_impact.sort(key=lambda r: r["最大绝对变动%"], reverse=True)

        return jsonify({
            "data": results,
            "category_impact": category_impact,
            "date_range": {"start": start_date, "end": end_date},
            "total": len(results),
            "has_category": has_cat,
            "has_name": bool(name_col),
        })

    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return jsonify({"error": f"价格波动分析失败：{exc}", "detail": tb[-1]}), 500
