# -*- coding: utf-8 -*-
"""成本模拟：由组件单价变动预测产品总价（非简单 BOM 相加），并评估预测可信度。"""
from __future__ import annotations

import math
import re
import sys
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import polars as pl

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from bom_date_key import calendar_date_key_yyyymmdd
from process_excel import (
    COL_COMPONENT,
    COL_CREATEDATE,
    COL_MAKTX,
    COL_PRODUCT,
    COL_QTY,
    COL_UNIT_PRICE,
    norm_material_code,
    normalize_columns,
)

MAX_REG_FEATURES = 40

# 业务口径：产品价格历史清单中的「单价/成本」均为 元/KG（每公斤）
PRICE_UNIT_LABEL = "元/KG"

_DATE_ALIASES = ("创建日期", "生价日期", "CREATEDATE", "生产日期", "日期")


def _norm_code_series(s: pd.Series) -> pd.Series:
    return norm_material_code(s)


def _timeline_month_key(dk: Any) -> str:
    """BOM 日键 YYYYMMDD 与价格时间线 YYYYMM 对齐到月份。"""
    s = str(dk).strip()
    if not s or s in ("nan", "None", "snapshot"):
        return s
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 6:
        return digits[:6]
    return s


def quote_month_key_from_series(s: pd.Series) -> pd.Series:
    """将「报价月份」或日期列规范为 YYYYMM（如 202512）。"""
    raw = norm_material_code(s)

    def _one(v: str) -> str:
        if not v:
            return ""
        digits = re.sub(r"\D", "", v)
        if len(digits) >= 6:
            return digits[:6]
        if len(digits) == 4:
            return digits
        dk = calendar_date_key_yyyymmdd(v)
        return dk[:6] if dk and len(dk) >= 6 else ""

    return raw.map(_one)


def quote_month_key_from_value(v: Any) -> str:
    if v is None:
        return ""
    return quote_month_key_from_series(pd.Series([v])).iloc[0]


def product_weight_kg(
    price_snap: dict[str, Any] | None,
    bom_rows: list[dict] | None,
) -> float | None:
    """产品重量（kg）：优先价格清单「重量」，否则 BOM 中 KG 用量合计。"""
    if price_snap:
        wt = price_snap.get("重量")
        if wt is not None:
            try:
                w = float(wt)
                if w > 1e-12:
                    return w
            except (TypeError, ValueError):
                pass
    if not bom_rows:
        return None
    total = 0.0
    for r in bom_rows:
        unit = str(r.get("计量单位") or "KG").strip().upper()
        if unit and unit not in ("KG", "KGM", "千克", "G", "克"):
            continue
        m = float(r.get("MENGE合计") or 0)
        if unit in ("G", "克"):
            m /= 1000.0
        total += m
    return total if total > 1e-12 else None


def per_kg_to_product_total(
    per_kg: float | None,
    weight_kg: float | None,
) -> float | None:
    """将元/KG 口径金额换算为该产品规格下的总价（元）。"""
    if per_kg is None:
        return None
    if weight_kg is None or weight_kg <= 1e-12:
        return float(per_kg)
    return float(per_kg) * float(weight_kg)


def _safe_float(v: Any) -> float | None:
    if v is None:
        return None
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


def normalize_price_list_cost_fields(
    mat: float | None,
    labor: float | None,
    total: float | None,
    weight_kg: float | None,
) -> dict[str, Any]:
    """
    识别价格清单中 材料成本/工费/总成本 是 元/件 还是 元/KG，并统一输出两种口径。

    成本结构预测公式（元/件）：
      新材料成本 = 基准材料成本(元/件) × (模拟BOM / 基准BOM)
      新总成本   = 新材料成本 + 工费(元/件)
      价格(元/kg) = 新总成本 / 重量(kg)   ← 只除一次重量
    """
    mat_f = _safe_float(mat)
    labor_f = _safe_float(labor)
    total_f = _safe_float(total)
    wt = _safe_float(weight_kg)

    out: dict[str, Any] = {
        "mat_piece": None,
        "labor_piece": None,
        "total_piece": None,
        "mat_per_kg": None,
        "labor_per_kg": None,
        "total_per_kg": None,
        "costs_are_per_piece": False,
    }
    if mat_f is None or labor_f is None:
        return out

    sum_ml = mat_f + labor_f
    per_piece = False
    if wt and wt > 1e-12:
        if total_f is not None and total_f > 1e-12:
            rel_sum_total = abs(sum_ml - total_f) / max(sum_ml, 1e-9)
            if rel_sum_total < 0.06:
                # 材料+工费与总成本同单位；数值明显大于典型「每公斤」报价 → 元/件
                if sum_ml > 12.0:
                    per_piece = True
            elif abs(sum_ml / wt - total_f) / max(total_f, 1e-9) < 0.08:
                # 材料/工费为整件，总成本列为每公斤
                per_piece = True
        elif sum_ml > 12.0:
            per_piece = True
    elif sum_ml > 12.0:
        per_piece = True

    out["costs_are_per_piece"] = per_piece
    if per_piece and wt and wt > 1e-12:
        out["mat_piece"] = mat_f
        out["labor_piece"] = labor_f
        if total_f is not None and abs(total_f - sum_ml) / max(sum_ml, 1e-9) < 0.06:
            out["total_piece"] = total_f
            out["total_per_kg"] = total_f / wt
        else:
            out["total_piece"] = sum_ml
            out["total_per_kg"] = (
                total_f if total_f is not None else sum_ml / wt
            )
        out["mat_per_kg"] = mat_f / wt
        out["labor_per_kg"] = labor_f / wt
    elif wt and wt > 1e-12:
        out["mat_per_kg"] = mat_f
        out["labor_per_kg"] = labor_f
        out["total_per_kg"] = total_f if total_f is not None else sum_ml
        out["mat_piece"] = mat_f * wt
        out["labor_piece"] = labor_f * wt
        out["total_piece"] = (
            total_f * wt if total_f is not None else (mat_f + labor_f) * wt
        )
    else:
        out["mat_per_kg"] = mat_f
        out["labor_per_kg"] = labor_f
        out["total_per_kg"] = total_f if total_f is not None else sum_ml
        out["mat_piece"] = mat_f
        out["labor_piece"] = labor_f
        out["total_piece"] = total_f if total_f is not None else sum_ml

    return out


def dedupe_warnings(warnings: list[str]) -> list[str]:
    """按完整文案去重，保留首次出现顺序。"""
    return list(dict.fromkeys(warnings))


def actual_price_per_kg_from_price_row(
    mat: Any,
    labor: Any,
    total: Any,
    weight_kg: Any,
) -> float | None:
    """将价格清单一行中的材料+工费/总成本统一为 元/kg（与主预测量纲一致）。"""
    mat_f = _safe_float(mat)
    labor_f = _safe_float(labor)
    total_f = _safe_float(total)
    wt = _safe_float(weight_kg)
    if mat_f is not None and labor_f is not None:
        norm = normalize_price_list_cost_fields(mat_f, labor_f, total_f, wt)
        per_kg = norm.get("total_per_kg")
        if per_kg is not None and per_kg > 1e-12:
            return float(per_kg)
    if total_f is not None and total_f > 1e-12:
        norm = normalize_price_list_cost_fields(None, None, total_f, wt)
        per_kg = norm.get("total_per_kg")
        if per_kg is not None and per_kg > 1e-12:
            return float(per_kg)
        if norm.get("costs_are_per_piece") and wt and wt > 0:
            return float(total_f) / wt
        return float(total_f)
    return None


def cost_structure_predict(
    base_mat: float,
    base_labor: float,
    bom_ratio: float,
    weight_kg: float | None,
    *,
    costs_are_per_piece: bool,
) -> tuple[float | None, float | None]:
    """成本结构法：返回 (新总成本元/件, 价格元/kg)。"""
    if costs_are_per_piece:
        total_piece = float(base_mat) * bom_ratio + float(base_labor)
        if weight_kg is None or weight_kg <= 1e-12:
            return total_piece, None
        return total_piece, total_piece / float(weight_kg)
    per_kg = float(base_mat) * bom_ratio + float(base_labor)
    total_piece = per_kg_to_product_total(per_kg, weight_kg)
    return total_piece, per_kg


def _weight_kg_from_row(row: dict[str, Any]) -> float | None:
    wt = row.get("重量_kg")
    if wt is None:
        wt = row.get("重量")
    if wt is None:
        return None
    try:
        w = float(wt)
        return w if w > 1e-12 else None
    except (TypeError, ValueError):
        return None


def history_per_kg_pair(row: dict[str, Any]) -> tuple[float | None, float | None]:
    """
    从历史样本行提取回归用 (BOM材料元/KG, 产品价元/KG)。
    清单中产品价本身为 元/KG；BOM 材料合计为整件时除以重量(kg)。
    """
    wt = _weight_kg_from_row(row)

    pp_kg = row.get("product_price_per_kg")
    if pp_kg is None:
        pp = row.get("product_price")
        if pp is not None:
            try:
                v = float(pp)
                pp_kg = v / wt if wt else v
            except (TypeError, ValueError):
                pp_kg = None
        else:
            pp_kg = None
    else:
        try:
            pp_kg = float(pp_kg)
        except (TypeError, ValueError):
            pp_kg = None

    bm_kg = row.get("bom_material_per_kg")
    if bm_kg is None:
        bm = row.get("bom_material")
        if bm is not None:
            try:
                v = float(bm)
                bm_kg = v / wt if wt else v
            except (TypeError, ValueError):
                bm_kg = None
    else:
        try:
            bm_kg = float(bm_kg)
        except (TypeError, ValueError):
            bm_kg = None

    return bm_kg, pp_kg


def product_total_from_per_kg_fields(
    fields: dict[str, Any],
    weight_kg: float | None,
) -> float | None:
    """材料成本、工费、总成本（均为元/KG）→ 产品总价。"""
    mat = fields.get("材料成本")
    labor = fields.get("工费")
    if mat is not None and labor is not None:
        try:
            return per_kg_to_product_total(float(mat) + float(labor), weight_kg)
        except (TypeError, ValueError):
            pass
    for key in ("product_price", "总成本", "产品价格"):
        v = fields.get(key)
        if v is not None:
            try:
                return per_kg_to_product_total(float(v), weight_kg)
            except (TypeError, ValueError):
                continue
    return None


def build_product_cost_history(raw_pl: pl.DataFrame) -> dict[str, list[dict[str, Any]]]:
    """从原始上传表按 产品[+日期] 汇总：BOM 材料合计、产品总价、材料成本、工费。"""
    raw = normalize_columns(raw_pl).to_pandas()
    if COL_PRODUCT not in raw.columns or COL_COMPONENT not in raw.columns:
        return {}

    raw[COL_PRODUCT] = _norm_code_series(raw[COL_PRODUCT])
    raw[COL_COMPONENT] = _norm_code_series(raw[COL_COMPONENT])
    raw["_menge"] = pd.to_numeric(
        raw[COL_QTY] if COL_QTY in raw.columns else 0, errors="coerce"
    ).fillna(0.0)
    raw["_price"] = pd.to_numeric(
        raw[COL_UNIT_PRICE] if COL_UNIT_PRICE in raw.columns else 0, errors="coerce"
    ).fillna(0.0)
    raw["_line"] = raw["_menge"] * raw["_price"]

    date_col = COL_CREATEDATE if COL_CREATEDATE in raw.columns else None
    if date_col:
        raw["_dk"] = raw[date_col].map(calendar_date_key_yyyymmdd).replace("", pd.NA)
    else:
        raw["_dk"] = "snapshot"

    product_price_col = "产品价格" if "产品价格" in raw.columns else None
    mat_cost_col = "材料成本" if "材料成本" in raw.columns else None
    labor_col = "工费" if "工费" in raw.columns else None

    quote_col = next(
        (c for c in ("报价号", "ZBJNO", "报价流水号", "ZSNO") if c in raw.columns),
        None,
    )
    if quote_col:
        raw["_qv"] = _norm_code_series(raw[quote_col]).replace("", pd.NA)
    else:
        raw["_qv"] = pd.NA

    # 先按 (产品, 报价版本, 日) 汇总单版本成本，再对同日多版本取均值
    version_daily = (
        raw.groupby([COL_PRODUCT, "_qv", "_dk"], sort=False, dropna=False)["_line"]
        .sum()
        .reset_index()
    )
    bom_by_day = (
        version_daily.groupby([COL_PRODUCT, "_dk"], sort=False)["_line"]
        .mean()
        .reset_index()
    )
    meta_by_day = raw.groupby([COL_PRODUCT, "_dk"], sort=False).last().reset_index()

    out: dict[str, list[dict[str, Any]]] = {}
    for _, bom_row in bom_by_day.iterrows():
        pid = str(bom_row[COL_PRODUCT]).strip()
        dk = bom_row["_dk"]
        bom_material = float(bom_row["_line"])
        if bom_material <= 1e-12:
            continue

        meta = meta_by_day[
            (meta_by_day[COL_PRODUCT] == bom_row[COL_PRODUCT])
            & (meta_by_day["_dk"] == dk)
        ]
        grp = meta.iloc[-1:] if len(meta) else pd.DataFrame()

        row: dict[str, Any] = {
            "date": str(dk) if pd.notna(dk) else "snapshot",
            "bom_material": round(bom_material, 6),
        }
        wt_kg: float | None = None
        if len(grp) and "重量" in grp.columns:
            wts = pd.to_numeric(grp["重量"], errors="coerce").dropna()
            if len(wts):
                wt_kg = float(wts.iloc[-1])

        pp_per_kg: float | None = None
        if product_price_col and len(grp):
            pp = pd.to_numeric(grp[product_price_col], errors="coerce").dropna()
            if len(pp):
                pp_per_kg = float(pp.iloc[-1])
        mat_per_kg: float | None = None
        labor_per_kg: float | None = None
        if mat_cost_col and len(grp):
            mc = pd.to_numeric(grp[mat_cost_col], errors="coerce").dropna()
            if len(mc):
                mat_per_kg = float(mc.iloc[-1])
                row["材料成本_per_kg"] = round(mat_per_kg, 6)
        if labor_col and len(grp):
            lb = pd.to_numeric(grp[labor_col], errors="coerce").dropna()
            if len(lb):
                labor_per_kg = float(lb.iloc[-1])
                row["工费_per_kg"] = round(labor_per_kg, 6)

        if pp_per_kg is None and mat_per_kg is not None and labor_per_kg is not None:
            pp_per_kg = mat_per_kg + labor_per_kg
        if pp_per_kg is None:
            continue

        row["product_price_per_kg"] = round(pp_per_kg, 6)
        row["product_price"] = round(
            per_kg_to_product_total(pp_per_kg, wt_kg) or pp_per_kg, 6
        )
        if wt_kg is not None:
            row["重量_kg"] = round(wt_kg, 6)
            row["bom_material_per_kg"] = round(bom_material / wt_kg, 6)

        out.setdefault(pid, []).append(row)

    for pid in out:
        out[pid].sort(key=lambda x: str(x.get("date") or ""))
    return out


def build_product_price_timeline(price_path: Path) -> dict[str, dict[str, dict[str, float]]]:
    """
    产品价格历史清单：每一行表示「某产品、某报价月份、某重量(kg)」下的
    **材料成本、工费、总成本（均为 元/KG）**。时间轴优先用 **报价月份**（YYYYMM）。
    """
    price = pd.read_excel(price_path, sheet_name=0)
    if COL_PRODUCT not in price.columns:
        if "ZMATNR" in price.columns:
            price = price.rename(columns={"ZMATNR": COL_PRODUCT})
        elif "MATNR" in price.columns:
            price = price.rename(columns={"MATNR": COL_PRODUCT})
    if COL_PRODUCT not in price.columns:
        return {}

    val_col = next(
        (c for c in ("总成本", "产品价格", "DMBTR", "标价", "出厂价", "销售价") if c in price.columns),
        None,
    )
    if not val_col:
        return {}
    price = price.rename(columns={val_col: "product_price"})
    date_col = next((c for c in _DATE_ALIASES if c in price.columns), None)

    price[COL_PRODUCT] = _norm_code_series(price[COL_PRODUCT])
    price["product_price"] = pd.to_numeric(price["product_price"], errors="coerce")
    if "报价月份" in price.columns:
        price["_tk"] = quote_month_key_from_series(price["报价月份"])
    elif date_col:
        price["_tk"] = price[date_col].map(calendar_date_key_yyyymmdd).replace("", pd.NA)
        price["_tk"] = quote_month_key_from_series(price["_tk"])
    else:
        price["_tk"] = "snapshot"
    if "重量" in price.columns:
        price["重量"] = pd.to_numeric(price["重量"], errors="coerce")
    if "报价月份" in price.columns:
        price["报价月份"] = quote_month_key_from_series(price["报价月份"])

    group_cols = [COL_PRODUCT, "_tk"]
    if "重量" in price.columns and price["重量"].notna().any():
        group_cols.append("重量")

    out: dict[str, dict[str, dict[str, float]]] = {}
    for keys, grp in price.groupby(group_cols, sort=False):
        if isinstance(keys, tuple):
            pid_s = str(keys[0]).strip()
            tk_s = str(keys[1]) if pd.notna(keys[1]) else "snapshot"
            wt = keys[2] if len(keys) > 2 else None
        else:
            pid_s, tk_s, wt = str(keys).strip(), "snapshot", None
        pp = grp["product_price"].dropna()
        if pp.empty:
            continue
        pp_kg = float(pp.iloc[-1])
        wt_f = float(wt) if wt is not None and pd.notna(wt) else None
        slot: dict[str, float] = {
            "product_price_per_kg": round(pp_kg, 6),
            "product_price": round(per_kg_to_product_total(pp_kg, wt_f) or pp_kg, 6),
            "价格口径": PRICE_UNIT_LABEL,
        }
        if "材料成本" in grp.columns:
            mc = pd.to_numeric(grp["材料成本"], errors="coerce").dropna()
            if len(mc):
                slot["材料成本"] = round(float(mc.iloc[-1]), 6)
        if "工费" in grp.columns:
            lb = pd.to_numeric(grp["工费"], errors="coerce").dropna()
            if len(lb):
                slot["工费"] = round(float(lb.iloc[-1]), 6)
        if wt_f is not None:
            slot["重量"] = round(wt_f, 6)
        if "报价月份" in grp.columns:
            qm = grp["报价月份"].dropna()
            if len(qm):
                slot["报价月份"] = str(qm.iloc[-1])
        out.setdefault(pid_s, {})[tk_s] = slot
    return out


def _price_at_date(panel: dict[str, float], date: str, default: float) -> float:
    """panel: 已排序 date→price；取 ≤ date 的最近价，否则 default。"""
    if not panel:
        return default
    if date in panel:
        return float(panel[date])
    best = default
    for d, p in panel.items():
        if str(d) <= str(date):
            best = float(p)
        elif str(d) > str(date):
            break
    return best


def expand_product_cost_history_dict(
    product_cost_history: dict[str, list[dict[str, Any]]],
    product_bom: dict[str, list[dict]],
    material_trend_df: pd.DataFrame | None,
    product_price_timeline: dict[str, dict[str, dict[str, float]]] | None,
) -> dict[str, list[dict[str, Any]]]:
    """
    用产品价格时间线补齐「日期 → BOM材料合计 + 产品价」样本。
    报价材料价格趋势已停用；BOM 材料合计使用当前 BOM 单价作为基准。
    """
    material_trend_df = None
    if material_trend_df is None or material_trend_df.empty:
        if not product_price_timeline:
            return dict(product_cost_history)
    trend_df = material_trend_df
    ppt = product_price_timeline or {}

    comp_panels: dict[str, dict[str, float]] = {}
    if trend_df is not None and not trend_df.empty:
        for comp, grp in trend_df.groupby(COL_COMPONENT, sort=False):
            comp_s = str(comp).strip()
            panel: dict[str, float] = {}
            for _, r in grp.sort_values(COL_CREATEDATE).iterrows():
                dk = str(r[COL_CREATEDATE])
                panel[dk] = float(r[COL_UNIT_PRICE])
            comp_panels[comp_s] = panel

    all_products = set(product_bom.keys()) | set(product_cost_history.keys()) | set(ppt.keys())
    out: dict[str, list[dict[str, Any]]] = dict(product_cost_history)

    for product in all_products:
        bom_rows = product_bom.get(product) or []
        if not bom_rows:
            continue
        menges: dict[str, float] = {}
        defaults: dict[str, float] = {}
        for r in bom_rows:
            c = str(r.get("组件编码") or "").strip()
            if not c:
                continue
            menges[c] = float(r.get("MENGE合计") or 0)
            defaults[c] = float(r.get("组件单价") or 0)

        existing = {
            _timeline_month_key(str(h.get("date"))): h for h in out.get(product, [])
        }
        pp_tl = {
            _timeline_month_key(k): v for k, v in (ppt.get(product, {}) or {}).items()
        }

        dates: set[str] = set(existing.keys()) | set(pp_tl.keys())
        for c in menges:
            if c in comp_panels:
                for dk in comp_panels[c]:
                    mk = _timeline_month_key(dk)
                    if mk:
                        dates.add(mk)
        if not dates:
            continue

        expanded_rows: list[dict[str, Any]] = []
        for mk in sorted(dates):
            if mk in ("", "nan", "None"):
                continue
            bom_material = 0.0
            for c, m in menges.items():
                if m <= 0:
                    continue
                p = _price_at_date(comp_panels.get(c, {}), mk, defaults.get(c, 0.0))
                bom_material += m * p
            if bom_material <= 1e-12:
                continue

            pp_kg: float | None = None
            wt_slot: float | None = None
            if mk in pp_tl:
                slot = pp_tl[mk]
                wt_slot = slot.get("重量")
                if slot.get("product_price_per_kg") is not None:
                    pp_kg = float(slot["product_price_per_kg"])
                else:
                    pp_total = slot.get("product_price")
                    if pp_total is not None and wt_slot and float(wt_slot) > 1e-12:
                        pp_kg = float(pp_total) / float(wt_slot)
                    elif pp_total is not None:
                        pp_kg = float(pp_total)
            elif mk in existing:
                ex = existing[mk]
                wt_slot = _weight_kg_from_row(ex)
                _, pp_kg = history_per_kg_pair(ex)

            if pp_kg is None:
                continue
            if wt_slot is None or float(wt_slot) <= 1e-12:
                continue
            wt_f = float(wt_slot)
            bom_per_kg = bom_material / wt_f

            row: dict[str, Any] = {
                "date": mk,
                "bom_material": round(bom_material, 6),
                "bom_material_per_kg": round(bom_per_kg, 6),
                "product_price_per_kg": round(pp_kg, 6),
                "product_price": round(per_kg_to_product_total(pp_kg, wt_f) or pp_kg, 6),
                "重量_kg": round(wt_f, 6),
                "_from_trend": mk not in existing,
            }
            if mk in pp_tl:
                slot = pp_tl[mk]
                if "材料成本" in slot:
                    row["材料成本_per_kg"] = slot["材料成本"]
                if "工费" in slot:
                    row["工费_per_kg"] = slot["工费"]
            expanded_rows.append(row)

        # 按月优先保留 BOM 原始快照（非趋势合成）
        by_month: dict[str, dict[str, Any]] = {}
        for r in expanded_rows:
            m = _timeline_month_key(str(r["date"]))
            prev = by_month.get(m)
            if prev is None or (prev.get("_from_trend") and not r.get("_from_trend")):
                by_month[m] = r
            elif prev.get("_from_trend") and r.get("_from_trend"):
                by_month[m] = r
        for m, h in existing.items():
            if m not in by_month and h.get("product_price") is not None:
                by_month[m] = {**h, "date": m}

        merged = sorted(by_month.values(), key=lambda x: str(x.get("date") or ""))
        if merged:
            out[product] = merged

    return out


def _regression_fit_grade(r2: float, n: int) -> str:
    """回归拟合效果文字等级（供页面展示）。"""
    if n < 3:
        return "未启用"
    if r2 >= 0.7:
        return "优"
    if r2 >= 0.35:
        return "良"
    if r2 > 0.05:
        return "一般"
    return "弱"


def _linear_fit_predict(
    x_hist: np.ndarray, y_hist: np.ndarray, x_new: float
) -> tuple[float | None, float, int, float | None, float | None]:
    """y ≈ a + b·x；返回 (预测值, R², 样本数, 截距a, 斜率b)。"""
    mask = np.isfinite(x_hist) & np.isfinite(y_hist) & (x_hist > 1e-12)
    x = x_hist[mask]
    y = y_hist[mask]
    n = int(x.size)
    if n < 3:
        return None, 0.0, n, None, None

    X = np.column_stack([np.ones(n), x])
    try:
        beta, _, _, _ = np.linalg.lstsq(X, y, rcond=None)
    except Exception:
        return None, 0.0, n, None, None

    intercept = float(beta[0])
    slope = float(beta[1])
    pred = float(intercept + slope * x_new)
    y_hat = X @ beta
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 0.0
    r2 = max(0.0, min(1.0, r2))
    return pred, r2, n, intercept, slope


PASS_THROUGH_MIN = 0.3
PASS_THROUGH_MAX = 1.5


def _history_xy_points(records: list[dict[str, Any]]) -> list[tuple[str, float, float]]:
    """返回按日期排序的 (date, BOM材料/kg, 产品价/kg) 有效点。"""
    pts: list[tuple[str, float, float]] = []
    for idx, h in enumerate(records or []):
        bm_kg, pp_kg = history_per_kg_pair(h)
        if bm_kg is None or pp_kg is None:
            continue
        if not (math.isfinite(bm_kg) and math.isfinite(pp_kg)):
            continue
        if bm_kg <= 1e-12 or pp_kg <= 1e-12:
            continue
        pts.append((str(h.get("date") or idx), float(bm_kg), float(pp_kg)))
    pts.sort(key=lambda x: x[0])
    return pts


def _robust_slopes_from_points(points: list[tuple[str, float, float]]) -> list[float]:
    """
    Theil-Sen 风格斜率样本：用所有成对变化估计产品价对 BOM 材料变化的传导。
    R² 在价格窄幅波动时容易偏低，这里只要求变化方向和量级稳定。
    """
    n = len(points)
    if n < 2:
        return []
    xs = np.array([p[1] for p in points], dtype=float)
    dx_floor = max(1e-9, float(np.nanmedian(np.abs(xs))) * 1e-5)
    slopes: list[float] = []
    for i in range(n - 1):
        xi = points[i][1]
        yi = points[i][2]
        for j in range(i + 1, n):
            dx = points[j][1] - xi
            if abs(dx) <= dx_floor:
                continue
            s = (points[j][2] - yi) / dx
            if math.isfinite(s) and -0.25 <= s <= 3.0:
                slopes.append(float(s))
    return slopes


def _pass_through_from_records(
    records: list[dict[str, Any]],
    source: str,
) -> dict[str, Any] | None:
    points = _history_xy_points(records)
    slopes = _robust_slopes_from_points(points)
    if not slopes:
        return None
    raw = float(np.nanmedian(np.array(slopes, dtype=float)))
    coef = min(PASS_THROUGH_MAX, max(PASS_THROUGH_MIN, raw))
    spread = float(np.nanpercentile(slopes, 75) - np.nanpercentile(slopes, 25)) if len(slopes) >= 4 else 0.0
    return {
        "coefficient": coef,
        "raw_coefficient": raw,
        "source": source,
        "sample_count": len(points),
        "slope_count": len(slopes),
        "spread": spread,
    }


def _estimate_pass_through(
    product: str,
    product_records: list[dict[str, Any]],
    all_history: dict[str, list[dict]] | None,
) -> dict[str, Any] | None:
    own = _pass_through_from_records(product_records, "本产品历史")
    if own is not None and own["slope_count"] >= 2:
        return own

    global_slopes: list[float] = []
    global_points = 0
    for pid, records in (all_history or {}).items():
        info = _pass_through_from_records(records, "全局产品历史")
        if info is None:
            continue
        pts = _history_xy_points(records)
        global_points += len(pts)
        global_slopes.extend(_robust_slopes_from_points(pts))
        if len(global_slopes) >= 5000:
            break

    if global_slopes:
        raw = float(np.nanmedian(np.array(global_slopes, dtype=float)))
        coef = min(PASS_THROUGH_MAX, max(PASS_THROUGH_MIN, raw))
        return {
            "coefficient": coef,
            "raw_coefficient": raw,
            "source": "全局产品历史",
            "sample_count": global_points,
            "slope_count": len(global_slopes),
            "spread": (
                float(np.nanpercentile(global_slopes, 75) - np.nanpercentile(global_slopes, 25))
                if len(global_slopes) >= 4
                else 0.0
            ),
        }

    return own


def _component_price_ranges(
    price_history: list[dict],
    product: str,
    components: list[str],
) -> dict[str, tuple[float, float]]:
    """各组件单价 min/max：仅使用产品 BOM 历史。"""
    ranges: dict[str, tuple[float, float]] = {}

    def _apply_frame(frame: pd.DataFrame, filter_product: bool) -> None:
        if frame.empty or COL_COMPONENT not in frame.columns:
            return
        sub = frame.copy()
        sub[COL_COMPONENT] = _norm_code_series(sub[COL_COMPONENT])
        price_col = COL_UNIT_PRICE if COL_UNIT_PRICE in sub.columns else "组件单价"
        if price_col not in sub.columns:
            return
        if filter_product and COL_PRODUCT in sub.columns:
            sub[COL_PRODUCT] = _norm_code_series(sub[COL_PRODUCT])
            sub = sub[sub[COL_PRODUCT] == product]
        sub["_p"] = pd.to_numeric(sub[price_col], errors="coerce")
        for comp in components:
            g = sub[sub[COL_COMPONENT] == comp]["_p"].dropna()
            if len(g) == 0:
                continue
            lo, hi = float(g.min()), float(g.max())
            if comp in ranges:
                ranges[comp] = (min(ranges[comp][0], lo), max(ranges[comp][1], hi))
            else:
                ranges[comp] = (lo, hi)

    if price_history:
        _apply_frame(pd.DataFrame(price_history), filter_product=True)
    return ranges


def _range_coverage_score(
    prices_new: dict[str, float],
    prices_old: dict[str, float],
    ranges: dict[str, tuple[float, float]],
) -> float:
    """调价后单价落在历史区间内的组件占比（未改价视为 1）。"""
    if not prices_new:
        return 0.5
    ok = 0
    total = 0
    for comp, p_new in prices_new.items():
        p_old = prices_old.get(comp, p_new)
        if abs(p_new - p_old) < 1e-9:
            ok += 1
            total += 1
            continue
        total += 1
        if comp not in ranges:
            ok += 0.5
            continue
        lo, hi = ranges[comp]
        if lo <= p_new <= hi:
            ok += 1
        elif p_new < lo:
            ok += max(0.0, 1.0 - (lo - p_new) / max(lo, 1e-6))
        else:
            ok += max(0.0, 1.0 - (p_new - hi) / max(hi, 1e-6))
    return ok / total if total else 0.5


def _extrapolation_risk(
    prices_new: dict[str, float],
    prices_old: dict[str, float],
    ranges: dict[str, tuple[float, float]],
    simulated_material: float,
    baseline_material: float,
) -> dict[str, Any]:
    """极端调价时给可信度设置上限，避免离谱外推仍显示高可信。"""
    factors: list[float] = []
    for comp, p_new in prices_new.items():
        p_old = float(prices_old.get(comp, p_new) or 0.0)
        if p_old <= 1e-12 or abs(float(p_new) - p_old) <= 1e-9:
            continue

        lo, hi = ranges.get(comp, (p_old, p_old))
        lo = float(lo) if lo and lo > 1e-12 else p_old
        hi = float(hi) if hi and hi > 1e-12 else p_old
        if p_new > hi:
            factors.append(float(p_new) / max(hi, 1e-6))
        elif p_new < lo:
            factors.append(lo / max(float(p_new), 1e-6))
        else:
            factors.append(1.0)

    material_factor = 1.0
    if baseline_material > 1e-12 and simulated_material > 1e-12:
        ratio = simulated_material / baseline_material
        material_factor = max(ratio, 1.0 / ratio)
        factors.append(material_factor)

    max_factor = max(factors) if factors else 1.0
    cap: float | None = None
    if max_factor >= 10.0:
        cap = 25.0
    elif max_factor >= 5.0:
        cap = 35.0
    elif max_factor >= 2.0:
        cap = 55.0
    elif max_factor >= 1.5:
        cap = 70.0

    return {
        "max_factor": max_factor,
        "material_factor": material_factor,
        "credibility_cap": cap,
    }


def _predict_product_price_legacy(
    *,
    product: str,
    simulated_material: float,
    baseline_material: float,
    price_snap: dict[str, float | None],
    product_cost_history: dict[str, list[dict]] | None,
    price_history: list[dict] | None,
    bom_rows: list[dict],
    prices_new: dict[str, float],
    prices_old: dict[str, float],
) -> dict[str, Any]:
    """
    返回预测产品总价、可信度 0–100、方法说明等。
    优先：成本结构与 BOM 成本传导；历史回归仅作为诊断参考。
    """
    weight_kg = product_weight_kg(price_snap, bom_rows)

    baseline_per_kg = price_snap.get("总成本")
    mat_per_kg = price_snap.get("材料成本")
    labor_per_kg = price_snap.get("工费")
    baseline_product = per_kg_to_product_total(baseline_per_kg, weight_kg)
    if baseline_product is None and mat_per_kg is not None and labor_per_kg is not None:
        baseline_product = per_kg_to_product_total(
            float(mat_per_kg) + float(labor_per_kg), weight_kg
        )

    candidates: list[tuple[str, float, float]] = []  # method, value, weight for agreement
    ratio = (simulated_material / baseline_material) if baseline_material > 1e-12 else 1.0

    # —— 方法 1：成本结构（材料/工费随 BOM 比例；元/件口径再 ÷ 重量得 元/kg）——
    mat_snap = price_snap.get("材料成本")
    labor_snap = price_snap.get("工费")
    total_snap = price_snap.get("总成本")
    cost_norm = normalize_price_list_cost_fields(
        _safe_float(mat_snap),
        _safe_float(labor_snap),
        _safe_float(total_snap),
        weight_kg,
    )
    mat_piece = cost_norm.get("mat_piece")
    labor_piece = cost_norm.get("labor_piece")
    if (
        mat_piece is not None
        and labor_piece is not None
        and float(mat_piece) > 1e-12
        and baseline_material > 1e-12
    ):
        pred_struct, pred_per_kg = cost_structure_predict(
            float(mat_piece),
            float(labor_piece),
            ratio,
            weight_kg,
            costs_are_per_piece=bool(cost_norm.get("costs_are_per_piece")),
        )
        if pred_struct is not None and pred_struct > 0:
            candidates.append(("成本结构(材料+工费)", pred_struct, 1.0))
    elif baseline_per_kg is not None and baseline_material > 1e-12:
        pred_per_kg = float(baseline_per_kg) * ratio
        pred_scale = per_kg_to_product_total(pred_per_kg, weight_kg) or pred_per_kg
        candidates.append(("总价比例", pred_scale, 0.85))

    # —— 方法 2：历史回归（元/KG）产品价每公斤 ~ BOM 材料每公斤 ——
    hist = (product_cost_history or {}).get(product, [])
    hist_valid: list[dict[str, Any]] = []
    for h in hist:
        bm_kg, pp_kg = history_per_kg_pair(h)
        if pp_kg is not None and bm_kg is not None and bm_kg > 1e-12:
            hist_valid.append(h)
    n_raw = len(
        [
            h
            for h in hist
            if h.get("product_price") is not None and h.get("bom_material", 0) > 1e-12
        ]
    )
    r2 = 0.0
    n_hist = len(hist_valid)
    reg_intercept: float | None = None
    reg_slope: float | None = None
    pred_reg_per_kg: float | None = None
    reg_enabled = False
    sim_per_kg: float | None = (
        simulated_material / weight_kg
        if weight_kg is not None and weight_kg > 1e-12
        else None
    )
    if n_hist >= 3 and weight_kg is not None and weight_kg > 1e-12:
        x = np.array(
            [history_per_kg_pair(h)[0] for h in hist_valid], dtype=float
        )
        y = np.array(
            [history_per_kg_pair(h)[1] for h in hist_valid], dtype=float
        )
        pred_reg_per_kg, r2, n_hist, reg_intercept, reg_slope = _linear_fit_predict(
            x, y, sim_per_kg
        )
        reg_enabled = True
        if pred_reg_per_kg is not None and pred_reg_per_kg > 0:
            pred_reg = per_kg_to_product_total(pred_reg_per_kg, weight_kg)
            if pred_reg is not None and pred_reg > 0:
                w = 0.9 + 0.1 * min(1.0, r2)
                candidates.append(("历史回归(元/KG)", pred_reg, w))

    # —— 方法 3：稳健成本传导（低波动时不依赖 R²）——
    pass_info = _estimate_pass_through(product, hist_valid, product_cost_history)
    pass_coef: float | None = None
    pred_pass_per_kg: float | None = None
    baseline_bom_per_kg: float | None = (
        baseline_material / weight_kg
        if weight_kg is not None and weight_kg > 1e-12
        else None
    )
    baseline_price_per_kg: float | None = None
    if baseline_per_kg is not None:
        try:
            baseline_price_per_kg = float(baseline_per_kg)
        except (TypeError, ValueError):
            baseline_price_per_kg = None
    if baseline_price_per_kg is None and hist_valid:
        last_bm, last_pp = history_per_kg_pair(hist_valid[-1])
        if last_pp is not None and last_pp > 1e-12:
            baseline_price_per_kg = float(last_pp)
        if baseline_bom_per_kg is None and last_bm is not None and last_bm > 1e-12:
            baseline_bom_per_kg = float(last_bm)

    if (
        pass_info is not None
        and sim_per_kg is not None
        and baseline_bom_per_kg is not None
        and baseline_price_per_kg is not None
    ):
        pass_coef = float(pass_info["coefficient"])
        pred_pass_per_kg = baseline_price_per_kg + pass_coef * (sim_per_kg - baseline_bom_per_kg)
        if pred_pass_per_kg > 0:
            pred_pass = per_kg_to_product_total(pred_pass_per_kg, weight_kg)
            if pred_pass is not None and pred_pass > 0:
                w = 0.98 if pass_info.get("source") == "本产品历史" else 0.9
                candidates.append((f"成本传导({pass_info['source']})", pred_pass, w))

    # 选主预测：成本结构优先，其次成本传导；回归仅作参考诊断
    primary_method = "BOM合计"
    predicted_product: float | None = None

    if candidates:
        struct_cands = [c for c in candidates if c[0].startswith("成本结构")]
        pass_cands = [c for c in candidates if c[0].startswith("成本传导")]
        if struct_cands:
            primary_method, predicted_product, _ = struct_cands[0]
        elif pass_cands:
            primary_method, predicted_product, _ = pass_cands[0]
        else:
            primary_method, predicted_product, _ = max(candidates, key=lambda c: c[2])
    elif simulated_material > 0:
        predicted_product = simulated_material
        primary_method = "BOM合计_fallback"

    if predicted_product is not None:
        predicted_product = max(0.0, round(predicted_product, 6))

    # —— 可信度 0–100 ——
    hist_score = min(1.0, n_hist / 8.0) * 22.0
    fit_score = r2 * 12.0 if n_hist >= 3 else 0.0

    comps = [str(r.get("组件编码") or "").strip() for r in bom_rows]
    price_ranges = _component_price_ranges(price_history or [], product, comps)
    coverage = _range_coverage_score(
        prices_new,
        prices_old,
        price_ranges,
    )
    range_score = coverage * 22.0
    extrapolation = _extrapolation_risk(
        prices_new,
        prices_old,
        price_ranges,
        simulated_material,
        baseline_material,
    )

    agreement_score = 15.0
    if len(candidates) >= 2:
        vals = [c[1] for c in candidates]
        mu = sum(vals) / len(vals)
        if mu > 1e-12:
            rel_spread = abs(max(vals) - min(vals)) / mu
            agreement_score = 15.0 * max(0.0, 1.0 - min(1.0, rel_spread))
    elif len(candidates) == 1:
        agreement_score = 12.0

    structure_score = 0.0
    if mat_per_kg is not None and labor_per_kg is not None:
        structure_score = 18.0
    elif baseline_per_kg is not None:
        structure_score = 10.0

    pass_score = 0.0
    if pass_info is not None:
        if pass_info.get("source") == "本产品历史":
            pass_score = min(16.0, 8.0 + min(8.0, float(pass_info.get("slope_count") or 0)))
        else:
            pass_score = min(12.0, 6.0 + min(6.0, float(pass_info.get("slope_count") or 0) / 10.0))

    credibility = round(
        min(
            100.0,
            max(
                0.0,
                hist_score
                + fit_score
                + range_score
                + agreement_score
                + structure_score
                + pass_score,
            ),
        ),
        1,
    )
    cap = extrapolation.get("credibility_cap")
    if cap is not None and credibility > cap:
        credibility = round(float(cap), 1)

    if credibility >= 75:
        level = "高"
    elif credibility >= 45:
        level = "中"
    else:
        level = "低"

    reasons: list[str] = []
    if n_hist >= 3:
        reasons.append(
            f"历史样本在{PRICE_UNIT_LABEL}口径：{n_hist} 期（产品价与 BOM 材料均已换算为每公斤）"
            + (f"，R²={r2:.2f}" if r2 > 0 else "")
        )
    else:
        reasons.append(
            "历史样本不足或缺少重量(kg)，主要依据当前成本结构或 BOM 合计推算"
        )
    if pass_info is not None:
        reasons.append(
            f"成本传导系数={pass_info['coefficient']:.3f}，来源：{pass_info['source']}；"
            "低波动月份不再以 R² 作为能否预测的硬门槛"
        )
    if mat_per_kg is not None and labor_per_kg is not None:
        qm = price_snap.get("报价月份")
        wt = price_snap.get("重量")
        ctx = []
        if qm is not None:
            ctx.append(f"报价月份{qm}")
        if wt is not None:
            ctx.append(f"重量{wt}kg")
        suffix = f"（{'、'.join(ctx)}）" if ctx else ""
        reasons.append(
            f"基准来自产品价格历史清单{suffix}：材料/工费/总价均为{PRICE_UNIT_LABEL}，"
            f"预测产品总价=每公斤价×重量；材料块随 BOM 材料合计同比例调整"
        )
    if weight_kg is None:
        reasons.append("未匹配到产品重量(kg)，预测价暂按元/KG口径展示，建议核对价格清单「重量」列")
    if coverage < 0.7:
        reasons.append("部分调价超出历史单价区间，外推风险偏高")
    if cap is not None:
        reasons.append(
            f"模拟调价最高约为可参考价格的 {extrapolation['max_factor']:.1f} 倍，"
            f"属于极端外推，可信度已封顶到 {cap:.0f} 分"
        )
    if len(candidates) >= 2:
        spread = abs(candidates[0][1] - candidates[-1][1])
        if spread / max(predicted_product or 1, 1e-6) > 0.15:
            reasons.append("多种算法结果差异较大，建议结合业务判断")

    reg_used = primary_method.startswith("历史回归") if primary_method else False
    reg_grade = _regression_fit_grade(r2, n_hist if reg_enabled else 0)
    reg_analysis: dict[str, Any] = {
        "是否启用": reg_enabled,
        "是否采用为最终预测": reg_used,
        "有效样本数": n_hist if reg_enabled else 0,
        "BOM直接样本数": n_raw,
        "R2": round(r2, 4) if reg_enabled else None,
        "R2_百分比": round(r2 * 100.0, 1) if reg_enabled else None,
        "拟合等级": reg_grade,
        "模型": "产品价(元/KG) ≈ 截距 + 斜率 × BOM材料(元/KG)",
        "截距": round(reg_intercept, 6) if reg_intercept is not None else None,
        "斜率": round(reg_slope, 6) if reg_slope is not None else None,
        "模拟点_BOM材料每公斤": round(sim_per_kg, 6) if sim_per_kg is not None else None,
        "回归预测_产品价每公斤": (
            round(pred_reg_per_kg, 6) if pred_reg_per_kg is not None else None
        ),
        "采用回归阈值R2": None,
        "成本传导系数": round(pass_coef, 6) if pass_coef is not None else None,
        "传导系数来源": pass_info.get("source") if pass_info is not None else None,
        "传导预测_产品价每公斤": (
            round(pred_pass_per_kg, 6) if pred_pass_per_kg is not None else None
        ),
        "说明": (
            f"共 {n_hist} 期有效样本（元/公斤口径）"
            + (f"，R²={r2:.2f}（{reg_grade}）" if reg_enabled else "；样本不足 3 期或未匹配重量，未做回归")
            + ("；回归已作为最终预测" if reg_used else "；回归仅作参考，主预测采用成本结构/成本传导" if reg_enabled else "")
        ),
    }

    return {
        "预测产品价格": predicted_product,
        "预测产品价格_每公斤": (
            round(predicted_product / weight_kg, 6)
            if predicted_product is not None and weight_kg and weight_kg > 1e-12
            else None
        ),
        "基准产品价格": baseline_product,
        "基准产品价格_每公斤": baseline_per_kg,
        "产品重量_kg": round(weight_kg, 6) if weight_kg is not None else None,
        "价格口径": PRICE_UNIT_LABEL,
        "预测方法": primary_method,
        "预测可信度": credibility,
        "可信度等级": level,
        "可信度说明": "；".join(reasons),
        "可信度明细": {
            "历史样本分": round(hist_score, 1),
            "拟合优度分": round(fit_score, 1),
            "价格区间分": round(range_score, 1),
            "方法一致性分": round(agreement_score, 1),
            "成本结构分": round(structure_score, 1),
            "成本传导分": round(pass_score, 1),
            "历史样本数": n_hist if reg_enabled else 0,
            "回归R2": round(r2, 4) if reg_enabled else None,
            "成本传导系数": round(pass_coef, 6) if pass_coef is not None else None,
            "极端外推倍数": round(float(extrapolation["max_factor"]), 4),
            "可信度封顶": cap,
            "满分": 100.0,
        },
        "回归分析": reg_analysis,
        "备选预测": [
            {"方法": m, "价格": round(v, 6)} for m, v, _ in candidates if m != primary_method
        ],
    }


# 敏感性分析：组件单价在用户输入值 ±N% 时的产品价格波动
SENSITIVITY_PCT = 0.10
# 模型历史误差区间：点估计 ± Z×MAE（稳健，对异常月不敏感）
MODEL_ERROR_Z = 1.5


# ---------------------------------------------------------------------------
# 新预测算法（成本结构 / 传导系数 + 敏感性 + 历史回测误差）
# ---------------------------------------------------------------------------

def _sim_bom_from_component_info(component_info: list[dict]) -> float:
    return sum(float(c["quantity"]) * float(c["price"]) for c in component_info)


def _point_per_kg_from_sim_bom(
    sim_bom: float,
    *,
    has_cost_structure: bool,
    base_data: dict,
    coeff: float = 0.030,
    base_for_coeff: float | None = None,
    base_bom_total_override: float | None = None,
) -> float | None:
    """由 BOM 批次合计推算产品价（元/kg）。"""
    base_bom_total = (
        float(base_bom_total_override)
        if base_bom_total_override is not None
        else float(base_data.get("base_bom_total") or 0.0)
    )
    if has_cost_structure and base_bom_total > 1e-12:
        base_mat = base_data.get("base_mat_piece") or base_data["base_mat"]
        base_labor = base_data.get("base_labor_piece") or base_data["base_labor"]
        per_piece = bool(base_data.get("costs_are_per_piece"))
        weight = base_data.get("weight")
        ratio = sim_bom / base_bom_total
        _, per_kg = cost_structure_predict(
            float(base_mat),
            float(base_labor),
            ratio,
            weight if weight and not math.isnan(weight) else None,
            costs_are_per_piece=per_piece,
        )
        return per_kg
    weight = base_data.get("weight")
    if weight is None or math.isnan(weight) or weight <= 1e-12:
        return None
    _base = base_for_coeff if base_for_coeff is not None else base_data.get("base_total") or 0.0
    if math.isnan(_base):
        _base = 0.0
    delta = sim_bom - base_bom_total
    return float(_base + coeff * delta)


def _component_info_with_prices(
    component_info: list[dict],
    price_overrides: dict[str, float],
) -> list[dict]:
    """复制 component_info 并覆盖指定组件单价。"""
    out: list[dict] = []
    for c in component_info:
        mid = str(c["mat_id"])
        price = float(price_overrides[mid]) if mid in price_overrides else float(c["price"])
        out.append({**c, "price": price})
    return out


def _compute_sensitivity_analysis(
    user_modified_prices: dict,
    component_info: list[dict],
    point_per_kg: float,
    *,
    has_cost_structure: bool,
    base_data: dict,
    coeff: float,
    base_for_coeff: float | None,
    pct: float = SENSITIVITY_PCT,
) -> list[dict[str, Any]]:
    """对用户改价的每个组件做 ±pct 确定性敏感性分析。"""
    if not user_modified_prices:
        return []

    base_bom_total = float(base_data.get("base_bom_total") or 0.0)
    base_mat = base_data.get("base_mat")
    denom_ppk = float(point_per_kg) if point_per_kg and point_per_kg > 1e-12 else None
    if denom_ppk is None:
        bom_center = _sim_bom_from_component_info(component_info)
        denom_ppk = _point_per_kg_from_sim_bom(
            bom_center,
            has_cost_structure=has_cost_structure,
            base_data=base_data,
            coeff=coeff,
            base_for_coeff=base_for_coeff,
        )
    if denom_ppk is None or denom_ppk <= 1e-12:
        return []

    items: list[dict[str, Any]] = []
    for mat_id, user_price in user_modified_prices.items():
        mat_id = str(mat_id).strip()
        comp = next((c for c in component_info if str(c["mat_id"]) == mat_id), None)
        if comp is None:
            continue
        user_price = float(user_price)
        quantity = float(comp["quantity"])
        price_lo = user_price * (1.0 - pct)
        price_hi = user_price * (1.0 + pct)

        info_lo = _component_info_with_prices(component_info, {mat_id: price_lo})
        info_hi = _component_info_with_prices(component_info, {mat_id: price_hi})
        bom_lo = _sim_bom_from_component_info(info_lo)
        bom_hi = _sim_bom_from_component_info(info_hi)

        ppk_lo = _point_per_kg_from_sim_bom(
            bom_lo, has_cost_structure=has_cost_structure, base_data=base_data,
            coeff=coeff, base_for_coeff=base_for_coeff,
        )
        ppk_hi = _point_per_kg_from_sim_bom(
            bom_hi, has_cost_structure=has_cost_structure, base_data=base_data,
            coeff=coeff, base_for_coeff=base_for_coeff,
        )
        if ppk_lo is None or ppk_hi is None:
            continue

        # 成本结构：产品价格(元/kg) 对 BOM 批次合计(元) 的边际 = base_mat/base_bom_total
        if (
            has_cost_structure
            and base_mat is not None
            and not math.isnan(base_mat)
            and base_bom_total > 1e-12
        ):
            d_bom_half = quantity * user_price * pct
            mat_for_ratio = base_data.get("base_mat_piece") or base_mat
            half_delta_piece = float(mat_for_ratio) * d_bom_half / base_bom_total
            wt = base_data.get("weight")
            if base_data.get("costs_are_per_piece") and wt and not math.isnan(wt) and wt > 0:
                half_delta_ppk = half_delta_piece / float(wt)
            else:
                half_delta_ppk = half_delta_piece
            product_pct = abs(half_delta_ppk / denom_ppk) * 100.0
        else:
            half_spread = max(abs(ppk_hi - denom_ppk), abs(denom_ppk - ppk_lo))
            product_pct = half_spread / denom_ppk * 100.0

        items.append({
            "材料编码": mat_id,
            "材料型号": comp.get("name") or mat_id,
            "用户单价": round(user_price, 4),
            "单价区间": [round(price_lo, 4), round(price_hi, 4)],
            "单价波动_pct": round(pct * 100, 1),
            "产品价格_kg区间": [round(ppk_lo, 4), round(ppk_hi, 4)],
            "产品价格变动_pct": round(product_pct, 2),
        })
    return items


def _compute_historical_model_error(
    base_data: dict,
    *,
    has_cost_structure: bool,
    coeff: float,
    base_for_coeff: float | None,
) -> dict[str, Any]:
    """
    用历史月份回测：predict(当月BOM) vs 当月实际价（元/kg）。
    剔除材料成本单月环比变动超过 2σ 的异常月后，用 MAE（中位数绝对误差）构造 ±1.5×MAE 区间。
    """
    p_full = base_data.get("p_hist_full")
    bom_m = base_data.get("bom_monthly")
    if p_full is None or bom_m is None or bom_m.empty:
        return {"可用": False, "说明": "缺少价格或 BOM 历史序列"}

    merged = bom_m.merge(p_full, on="报价月份", how="inner").sort_values("报价月份")
    if merged.empty:
        return {"可用": False, "说明": "价格与 BOM 历史月份无法对齐"}

    default_weight = base_data.get("weight")
    rows: list[dict[str, Any]] = []
    for _, row in merged.iterrows():
        bom_t = float(row["BOM批次合计"])
        mat_t = row.get("材料成本")
        labor_t = row.get("工费")
        total_t = row.get("总成本")
        wt_t = row.get("重量") if "重量" in row.index else default_weight
        if wt_t is None or (isinstance(wt_t, float) and math.isnan(wt_t)):
            wt_t = default_weight

        actual = actual_price_per_kg_from_price_row(mat_t, labor_t, total_t, wt_t)
        if actual is None or actual <= 1e-12:
            continue

        pred = _point_per_kg_from_sim_bom(
            bom_t,
            has_cost_structure=has_cost_structure,
            base_data=base_data,
            coeff=coeff,
            base_for_coeff=base_for_coeff,
            base_bom_total_override=bom_t,
        )
        if pred is None:
            continue
        mat_v = float(mat_t) if mat_t is not None and not pd.isna(mat_t) else float("nan")
        rows.append({
            "报价月份": row["报价月份"],
            "实际": round(actual, 6),
            "预测": round(float(pred), 6),
            "误差": round(float(pred) - actual, 6),
            "材料成本": mat_v,
        })

    if len(rows) < 2:
        return {
            "可用": False,
            "说明": f"有效回测月份仅 {len(rows)} 期，不足以估计误差",
            "样本数": len(rows),
        }

    df_err = pd.DataFrame(rows)
    mat_series = df_err["材料成本"].dropna()
    outlier_months: set[str] = set()
    if len(mat_series) >= 3:
        mat_diff = mat_series.diff().dropna()
        if len(mat_diff) >= 2:
            diff_std = float(mat_diff.std())
            if diff_std > 1e-12:
                threshold = 2.0 * diff_std
                for idx, dmat in mat_diff.items():
                    if abs(float(dmat)) > threshold:
                        outlier_months.add(str(df_err.loc[idx, "报价月份"]))

    df_use = df_err[~df_err["报价月份"].astype(str).isin(outlier_months)]
    if len(df_use) < 2:
        df_use = df_err

    err_arr = df_use["误差"].to_numpy(dtype=float)
    mae_robust = float(np.median(np.abs(err_arr)))
    rmse = float(np.sqrt(np.mean(err_arr ** 2)))
    bias = float(np.median(err_arr))

    mat_kg_vals: list[float] = []
    for _, row in merged.iterrows():
        mat_t = row.get("材料成本")
        if mat_t is None or pd.isna(mat_t):
            continue
        wt_t = row.get("重量") if "重量" in row.index else default_weight
        if wt_t is None or (isinstance(wt_t, float) and math.isnan(wt_t)):
            wt_t = default_weight
        norm = normalize_price_list_cost_fields(
            float(mat_t), 0.0, None, _safe_float(wt_t)
        )
        mk = norm.get("mat_per_kg")
        if mk is not None and mk > 1e-12:
            mat_kg_vals.append(float(mk))

    mat_cv: float | None = None
    mat_cv_pct: float | None = None
    n_mat_months = len(mat_kg_vals)
    if n_mat_months >= 2:
        arr = np.array(mat_kg_vals, dtype=float)
        mu = float(arr.mean())
        if mu > 1e-12:
            mat_cv = float(arr.std(ddof=1) / mu)
            mat_cv_pct = round(mat_cv * 100.0, 1)

    narrow_note: str | None = None
    if mat_cv is not None and mat_cv < 0.05 and n_mat_months >= 6:
        narrow_note = (
            f"区间较窄是因为该产品历史成本结构非常稳定（{n_mat_months} 个月材料成本波动 "
            f"<5%，实际约 {mat_cv_pct:g}%），并非计算异常。"
        )

    return {
        "可用": True,
        "样本数": int(len(df_use)),
        "总样本数": int(len(df_err)),
        "剔除异常月数": int(len(outlier_months)),
        "MAE": round(mae_robust, 4),
        "RMSE": round(rmse, 4),
        "平均偏差": round(bias, 4),
        "方法": "cost_structure_ratio" if has_cost_structure else "conduction_coeff",
        "区间口径": f"点估计 ± {MODEL_ERROR_Z}×MAE（稳健）",
        "材料成本历史月数": n_mat_months,
        "材料成本历史波动CV": round(mat_cv, 4) if mat_cv is not None else None,
        "材料成本历史波动CV_百分比": mat_cv_pct,
        "区间较窄说明": narrow_note,
    }


def _model_error_interval(point_per_kg: float, model_error: dict[str, Any]) -> dict[str, Any]:
    """由稳健 MAE 构造点估计 ± Z×MAE 区间（元/kg）。"""
    if not model_error.get("可用") or point_per_kg is None:
        return model_error
    mae = float(model_error.get("MAE") or model_error.get("RMSE") or 0.0)
    z = MODEL_ERROR_Z
    half = z * mae
    lo = max(0.0, point_per_kg - half)
    hi = point_per_kg + half
    out = dict(model_error)
    out["预测区间_kg"] = [round(lo, 4), round(hi, 4)]
    out["区间半宽"] = round(half, 4)
    return out

def _prepare_base_data(
    product_id,
    bom_df: pd.DataFrame,
    price_df: pd.DataFrame,
) -> "tuple[dict | None, str | None]":
    """
    从 BOM 历史与产品价格历史中提取预测所需的基准数据。

    bom_df  : normalize_columns 后的 pandas DataFrame，含以下列：
              COL_PRODUCT(产品编码), COL_COMPONENT(组件编码), COL_QTY(MENGE),
              COL_UNIT_PRICE(组件单价), COL_MAKTX(MAKTX), COL_CREATEDATE(CREATEDATE)
    price_df: 产品价格历史清单 DataFrame，含列：
              产品编码, 报价月份(YYYYMM), 重量, 材料成本, 工费, 总成本
    返回 (data_dict, None) 或 (None, error_msg)。
    """
    if price_df is None or price_df.empty:
        return None, "未上传产品价格历史清单，无法预测"

    pid_str = norm_material_code(pd.Series([product_id])).iloc[0]

    # ── 价格历史 ──────────────────────────────────────────────────────────
    price_pid_col = "产品编码" if "产品编码" in price_df.columns else (
        "ZMATNR" if "ZMATNR" in price_df.columns else None
    )
    if price_pid_col is None:
        return None, "价格历史缺少产品编码列"

    p_hist = price_df[norm_material_code(price_df[price_pid_col]) == pid_str].copy()

    if "报价月份" not in p_hist.columns:
        return None, "价格历史缺少「报价月份」列"

    p_hist["_qm"] = norm_material_code(p_hist["报价月份"]).str[:6]
    p_hist = p_hist.sort_values("_qm").drop_duplicates("_qm", keep="last")

    if p_hist.empty:
        return None, "价格历史中无该产品记录"

    latest_price = p_hist.iloc[-1]

    # ── BOM 历史 ──────────────────────────────────────────────────────────
    if bom_df is None or bom_df.empty or COL_PRODUCT not in bom_df.columns:
        return None, "BOM历史数据不可用"

    b_hist = bom_df[norm_material_code(bom_df[COL_PRODUCT]) == pid_str].copy()

    if b_hist.empty:
        return None, "BOM历史中无该产品记录"

    b_hist[COL_QTY] = pd.to_numeric(b_hist[COL_QTY], errors="coerce").fillna(0.0)
    b_hist[COL_UNIT_PRICE] = pd.to_numeric(b_hist[COL_UNIT_PRICE], errors="coerce").fillna(0.0)

    # 规范化日期列为 8 位字符串
    if COL_CREATEDATE in b_hist.columns:
        b_hist["_date"] = norm_material_code(b_hist[COL_CREATEDATE]).str[:8]
    else:
        b_hist["_date"] = "00000000"

    latest_date = b_hist["_date"].max()
    latest_bom_raw = b_hist[b_hist["_date"] == latest_date].copy()

    # ── 组件历史统计 ──────────────────────────────────────────────────────
    maktx_col = COL_MAKTX if COL_MAKTX in b_hist.columns else COL_COMPONENT
    comp_group = b_hist.groupby([COL_COMPONENT, maktx_col], sort=False)[COL_UNIT_PRICE]
    comp_stats = comp_group.agg(
        历史均价="mean",
        历史std="std",
        历史样本数="count",
        历史最低="min",
        历史最高="max",
    ).reset_index().rename(columns={COL_COMPONENT: "材料编码", maktx_col: "材料型号"})
    comp_stats["材料编码"] = norm_material_code(comp_stats["材料编码"])
    comp_stats["历史std"] = comp_stats["历史std"].fillna(comp_stats["历史均价"] * 0.05)

    # ── 历史BOM合计月度序列 ───────────────────────────────────────────────
    b_hist["_line_cost"] = b_hist[COL_QTY] * b_hist[COL_UNIT_PRICE]
    bom_daily = (
        b_hist.groupby("_date", sort=True)["_line_cost"]
        .sum()
        .reset_index()
        .rename(columns={"_date": "生价日期", "_line_cost": "BOM批次合计"})
    )
    bom_daily["报价月份"] = bom_daily["生价日期"].str[:6]
    bom_monthly = bom_daily.groupby("报价月份")["BOM批次合计"].mean().reset_index()

    # ── 基准BOM合计（最新期：用量 × 当前 BOM 单价，与页面「加载 BOM」一致）──
    lb_comp = norm_material_code(latest_bom_raw[COL_COMPONENT])
    lb_qty = latest_bom_raw[COL_QTY].values
    lb_unit_prices = pd.to_numeric(
        latest_bom_raw[COL_UNIT_PRICE], errors="coerce"
    ).fillna(0.0).values
    base_bom_total = float(np.dot(lb_qty, lb_unit_prices))

    # ── 构建标准化的 latest_bom DataFrame ────────────────────────────────
    latest_bom = latest_bom_raw[[COL_COMPONENT, maktx_col, COL_QTY, COL_UNIT_PRICE]].copy()
    latest_bom["材料编码"] = norm_material_code(latest_bom[COL_COMPONENT])
    latest_bom = latest_bom.rename(columns={
        maktx_col: "材料型号",
        COL_QTY: "组件数量",
        COL_UNIT_PRICE: "材料单价",
    }).drop(columns=[COL_COMPONENT], errors="ignore")

    # ── 从 latest_price 提取基准成本字段 ─────────────────────────────────
    def _get_num(row, col):
        v = row.get(col) if isinstance(row, pd.Series) else None
        if v is None:
            return float("nan")
        try:
            return float(v)
        except (TypeError, ValueError):
            return float("nan")

    weight = _get_num(latest_price, "重量")
    base_mat = _get_num(latest_price, "材料成本")
    base_labor = _get_num(latest_price, "工费")
    base_total = _get_num(latest_price, "总成本")

    # ── 构建 p_hist 供传导系数使用（需要 报价月份 & 材料成本） ─────────────
    p_hist_coeff = p_hist[["_qm"] + [c for c in ["材料成本"] if c in p_hist.columns]].copy()
    p_hist_coeff = p_hist_coeff.rename(columns={"_qm": "报价月份"})
    if "材料成本" not in p_hist_coeff.columns:
        p_hist_coeff["材料成本"] = float("nan")

    # 完整价格序列（供历史回测 MAE；含重量以便元/件 → 元/kg）
    p_hist_full_cols = ["_qm"] + [
        c for c in ("材料成本", "工费", "总成本", "重量") if c in p_hist.columns
    ]
    p_hist_full = p_hist[p_hist_full_cols].copy().rename(columns={"_qm": "报价月份"})
    for c in ("材料成本", "工费", "总成本", "重量"):
        if c in p_hist_full.columns:
            p_hist_full[c] = pd.to_numeric(p_hist_full[c], errors="coerce")

    cost_norm = normalize_price_list_cost_fields(
        None if math.isnan(base_mat) else base_mat,
        None if math.isnan(base_labor) else base_labor,
        None if math.isnan(base_total) else base_total,
        None if math.isnan(weight) else weight,
    )

    return {
        "latest_price": latest_price,
        "latest_bom": latest_bom,
        "comp_stats": comp_stats,
        "p_hist": p_hist_coeff,
        "p_hist_full": p_hist_full,
        "bom_monthly": bom_monthly,
        "base_bom_total": base_bom_total,
        "n_price_months": len(p_hist),
        "n_bom_months": bom_monthly.shape[0],
        "weight": weight,
        "base_mat": base_mat,
        "base_labor": base_labor,
        "base_total": base_total,
        "base_mat_piece": cost_norm.get("mat_piece"),
        "base_labor_piece": cost_norm.get("labor_piece"),
        "base_total_piece": cost_norm.get("total_piece"),
        "base_total_per_kg": cost_norm.get("total_per_kg"),
        "costs_are_per_piece": cost_norm.get("costs_are_per_piece"),
    }, None


def _estimate_conduction_coeff(base_data: dict) -> "tuple[float, str]":
    """
    用历史月度数据估计：产品材料成本 随 BOM批次合计 的变化率。
    复用 Theil-Sen 风格稳健斜率（_robust_slopes_from_points），系数区间随样本自适应。
    返回 (稳健传导系数, 样本质量评级)。
    """
    fallback = 0.030
    abs_min, abs_max = 0.005, 0.15

    bom_m = base_data["bom_monthly"]
    p_h = base_data["p_hist"]

    merged = bom_m.merge(p_h[["报价月份", "材料成本"]], on="报价月份", how="inner")
    merged = merged.dropna(subset=["材料成本", "BOM批次合计"])
    merged = merged.sort_values("报价月份")

    if len(merged) < 3:
        return fallback, "global_fallback"

    points = [
        (
            str(r["报价月份"]),
            float(r["BOM批次合计"]),
            float(r["材料成本"]),
        )
        for _, r in merged.iterrows()
        if float(r["BOM批次合计"]) > 1e-12 and float(r["材料成本"]) > 1e-12
    ]
    slopes = _robust_slopes_from_points(points)
    if len(slopes) < 2:
        return fallback, "global_fallback"

    coeff = float(np.nanmedian(np.array(slopes, dtype=float)))
    if len(slopes) >= 5:
        lo, hi = np.nanpercentile(slopes, [5, 95])
        lo = max(abs_min, float(lo))
        hi = min(abs_max, float(hi))
    else:
        lo, hi = abs_min, abs_max
    coeff = float(np.clip(coeff, lo, hi))

    quality = "product_specific" if len(slopes) >= 5 else "product_limited"
    return coeff, quality


def predict_product_price(
    product_id,
    user_modified_prices: dict,
    bom_df: "pd.DataFrame | None",
    price_df: "pd.DataFrame | None",
    *,
    reference_prices: dict | None = None,
) -> dict[str, Any]:
    """
    成本预测主函数（四级优先级 + 敏感性分析 + 历史回测误差区间）。

    Parameters
    ----------
    product_id          : 产品编码（str 或 int）
    user_modified_prices: {组件编码(str): 单价(float)}；未出现的组件用 reference_prices 或 BOM 当前单价。
    reference_prices  : 加载 BOM 时页面上的组件单价快照；「是否改价」= 模拟价相对此快照是否变化。
    bom_df / price_df   : BOM 与产品价格历史 DataFrame

    Returns
    -------
    dict：method、point_estimate、point_per_kg、confidence_score、warnings、
    sensitivity（改价组件 ±10% 敏感性）、model_error（历史回测 RMSE 与 ±2σ 区间，与调价无关）。
    """
    base_data, err = _prepare_base_data(product_id, bom_df, price_df)
    if base_data is None:
        return {"error": err}

    weight = base_data["weight"]
    base_mat = base_data["base_mat"]
    base_labor = base_data["base_labor"]
    base_total = base_data["base_total"]
    lb = base_data["latest_bom"]
    cs = base_data["comp_stats"]
    ref_map = {
        str(k).strip(): float(v)
        for k, v in (reference_prices or {}).items()
        if str(k).strip()
    }
    if ref_map:
        base_bom_total = 0.0
        for _, row in lb.iterrows():
            mid = str(row["材料编码"]).strip()
            qty = float(row.get("组件数量") or 0.0)
            unit = ref_map.get(mid, float(row.get("材料单价") or 0.0))
            base_bom_total += qty * unit
    else:
        base_bom_total = float(base_data["base_bom_total"])

    overrides = {
        str(k).strip(): float(v) for k, v in user_modified_prices.items() if str(k).strip()
    }

    warnings: list[str] = []
    score = 100
    # 同一材料型号可能对应多行 BOM / 多个编码，按「材料型号」只提示一次
    material_warned: set[str] = set()

    # ── 步骤A：计算模拟BOM合计，收集各组件信息 ───────────────────────────
    sim_bom_point = 0.0
    component_info: list[dict] = []

    cs_lookup = cs.set_index("材料编码")

    for _, row in lb.iterrows():
        mat_id = str(row["材料编码"]).strip()
        quantity = float(row.get("组件数量") or 0.0)
        mat_name = str(row.get("材料型号") or mat_id)

        stat = cs_lookup.loc[mat_id] if mat_id in cs_lookup.index else None

        row_unit = float(row.get("材料单价") or 0.0)
        ref_unit = ref_map.get(mat_id, row_unit)
        if mat_id in overrides:
            new_price = float(overrides[mat_id])
            is_fixed = abs(new_price - ref_unit) > 1e-9
        else:
            new_price = ref_unit
            is_fixed = False

        if stat is not None and is_fixed:
            hist_mean_chk = float(stat["历史均价"])
            hist_min = float(stat["历史最低"])
            hist_max = float(stat["历史最高"])
            hist_std_chk = float(stat["历史std"])

            if hist_std_chk > 0 and mat_name not in material_warned:
                shock_sigma = abs(new_price - hist_mean_chk) / hist_std_chk
                if shock_sigma > 10.0:
                    score = min(score, 20)
                    warnings.append(
                        f"🚨 【{mat_name}】调价幅度 {shock_sigma:.1f}σ，属于极端异常值，"
                        f"预测结果不具有参考意义（历史均价 {hist_mean_chk:.2f}，"
                        f"历史区间 [{hist_min:.2f}, {hist_max:.2f}]）"
                    )
                    material_warned.add(mat_name)
                elif shock_sigma > 5.0:
                    score = min(score, 50)
                    warnings.append(
                        f"⚠️ 【{mat_name}】调价幅度 {shock_sigma:.1f}σ，远超历史范围，可信度极低"
                    )
                    material_warned.add(mat_name)
                elif shock_sigma > 3.0:
                    score = min(score, 75)
                    warnings.append(
                        f"⚠️ 【{mat_name}】调价幅度 {shock_sigma:.1f}σ，超出历史波动范围"
                    )
                    material_warned.add(mat_name)
                elif shock_sigma > 2.0:
                    score = min(score, 92)

            if hist_mean_chk > 1e-12 and mat_name not in material_warned:
                abs_ratio = abs(new_price - hist_mean_chk) / hist_mean_chk
                if abs_ratio > 5.0:
                    score = min(score, 15)
                    warnings.append(
                        f"🚨 【{mat_name}】模拟单价（{new_price:.2f}）为历史均价（{hist_mean_chk:.2f}）"
                        f"的 {abs_ratio + 1:.0f} 倍，预测结果不具参考价值"
                    )
                    material_warned.add(mat_name)
                elif abs_ratio > 2.0:
                    score = min(score, 45)

        sim_bom_point += quantity * new_price
        hist_mean = float(stat["历史均价"]) if stat is not None else new_price
        hist_std = float(stat["历史std"]) if stat is not None else new_price * 0.05
        component_info.append({
            "mat_id": mat_id,
            "name": mat_name,
            "quantity": quantity,
            "price": new_price,
            "is_fixed": is_fixed,
            "hist_mean": hist_mean,
            "hist_std": max(hist_std, 1e-9),
        })

    # ── 步骤B：选择预测方法（四级优先级） ────────────────────────────────
    has_cost_structure = (
        not (math.isnan(base_mat) if base_mat is not None else True)
        and not (math.isnan(base_labor) if base_labor is not None else True)
        and weight is not None and not math.isnan(weight) and weight > 0
        and base_mat > 0
        and base_bom_total > 0
    )

    coeff: float = 0.030
    coeff_quality: str = "global_fallback"
    _base_for_coeff: float = 0.0

    if has_cost_structure:
        ratio = sim_bom_point / base_bom_total
        mat_piece = base_data.get("base_mat_piece") or base_mat
        labor_piece = base_data.get("base_labor_piece") or base_labor
        per_piece = bool(base_data.get("costs_are_per_piece"))
        point_total, point_per_kg = cost_structure_predict(
            float(mat_piece),
            float(labor_piece),
            ratio,
            weight if weight and not math.isnan(weight) else None,
            costs_are_per_piece=per_piece,
        )
        if point_total is None:
            point_total = 0.0
        if point_per_kg is None:
            point_per_kg = point_total
        method = "cost_structure_ratio"
    else:
        coeff, coeff_quality = _estimate_conduction_coeff(base_data)
        delta_bom = sim_bom_point - base_bom_total

        _base_for_coeff = base_data.get("base_total_per_kg")
        if _base_for_coeff is None or (
            isinstance(_base_for_coeff, float) and math.isnan(_base_for_coeff)
        ):
            _base_for_coeff = base_total if (base_total and not math.isnan(base_total)) else 0.0
            if base_data.get("costs_are_per_piece") and weight and not math.isnan(weight) and weight > 0:
                _base_for_coeff = float(_base_for_coeff) / float(weight)
        point_per_kg = float(_base_for_coeff) + coeff * delta_bom
        _weight_safe = weight if (weight and not math.isnan(weight) and weight > 0) else 1.0
        point_total = point_per_kg * _weight_safe
        method = f"conduction_coeff({coeff_quality})"
        score = min(score, 80)
        if coeff_quality == "global_fallback":
            score = min(score, 65)
            warnings.append("该产品BOM历史样本不足，传导系数使用全局兜底值，可信度较低")

        if base_bom_total <= 0:
            warnings.append("基准BOM合计为零（数据异常），已降级至传导系数模型")

    any_adjusted = any(c.get("is_fixed") for c in component_info)
    # 未改任何组件单价时，主结果与价格清单基准一致（避免加载 BOM 就出现虚假差额）
    if not any_adjusted:
        bl_piece = base_data.get("base_total_piece")
        bl_kg = base_data.get("base_total_per_kg")
        if bl_piece is not None and not (
            isinstance(bl_piece, float) and math.isnan(bl_piece)
        ):
            point_total = float(bl_piece)
        if bl_kg is not None and not (isinstance(bl_kg, float) and math.isnan(bl_kg)):
            point_per_kg = float(bl_kg)

    # ── 步骤C：敏感性分析（仅实际改价组件，±10% 确定性推算）────────────────
    modified_only = {
        str(c["mat_id"]): float(c["price"])
        for c in component_info
        if c.get("is_fixed")
    }
    sensitivity = _compute_sensitivity_analysis(
        modified_only,
        component_info,
        float(point_per_kg),
        has_cost_structure=has_cost_structure,
        base_data=base_data,
        coeff=coeff,
        base_for_coeff=_base_for_coeff,
    )

    # ── 步骤D：模型历史误差（回测 MAE + ±1.5×MAE，与用户调价无关）────────
    model_error_raw = _compute_historical_model_error(
        base_data,
        has_cost_structure=has_cost_structure,
        coeff=coeff,
        base_for_coeff=_base_for_coeff,
    )
    model_error = _model_error_interval(float(point_per_kg), model_error_raw)

    # ── 步骤E：可信度评分修正项 ───────────────────────────────────────────
    n_price_months = base_data["n_price_months"]
    if n_price_months >= 12:
        score = min(100, score + 5)
    elif n_price_months < 4:
        score = min(score, 85)
        warnings.append(f"价格历史仅 {n_price_months} 个月，样本偏少")

    p_hist_df = base_data["p_hist"]
    labor_vals = p_hist_df["材料成本"].dropna() if "材料成本" in p_hist_df.columns else pd.Series(dtype=float)
    if len(labor_vals) >= 2:
        labor_mean = float(labor_vals.mean())
        labor_std = float(labor_vals.std())
        if labor_mean > 0:
            labor_cv = labor_std / labor_mean
            if labor_cv > 0.10:
                score = min(score, 90)
                warnings.append(
                    f"材料成本历史波动较大（CV={labor_cv:.1%}），固定基准假设可能低估不确定性"
                )

    score = int(np.clip(score, 0, 100))
    warnings = dedupe_warnings(warnings)

    return {
        "method": method,
        "point_estimate": round(float(point_total), 4),
        "point_per_kg": round(float(point_per_kg), 4),
        "user_adjusted_prices": any_adjusted,
        "confidence_score": score,
        "warnings": warnings,
        "sensitivity": sensitivity,
        "model_error": model_error,
        "detail": {
            "base_mat_cost": base_mat,
            "base_labor": base_labor,
            "base_weight": weight,
            "base_bom_total": round(base_bom_total, 4),
            "sim_bom_point": round(sim_bom_point, 4),
            "bom_ratio": round(sim_bom_point / base_bom_total, 6) if base_bom_total > 0 else None,
            "n_price_months": n_price_months,
            "components": component_info,
        },
    }
