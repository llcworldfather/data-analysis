# -*- coding: utf-8 -*-
"""成本模拟：由组件单价变动预测产品总价（非简单 BOM 相加），并评估预测可信度。"""
from __future__ import annotations

import math
import os
import random
import re
import sys
from pathlib import Path
from typing import Any, Literal

import numpy as np
import pandas as pd
import polars as pl

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from bom_date_key import calendar_date_key_yyyymmdd
from process_excel import (
    COL_CATEGORY,
    COL_COMPONENT,
    COL_CREATEDATE,
    COL_MAKTX,
    COL_PRODUCT,
    COL_QTY,
    COL_UNIT,
    COL_UNIT_PRICE,
    norm_material_code,
    norm_material_code_expr,
    norm_material_code_scalar,
    normalize_columns,
    pandas_material_code_dtypes,
    read_price_excel,
)

MAX_REG_FEATURES = 40

# 业务口径：产品价格历史清单中的「单价/成本」均为 元/KG（每公斤）
PRICE_UNIT_LABEL = "元/KG"

# 传导系数：同分类兜底最少斜率条数 / 最少产品数
PASS_THROUGH_CATEGORY_MIN_SLOPES = 30
PASS_THROUGH_CATEGORY_MIN_PRODUCTS = 3
# Theil-Sen 时间衰减（月）：w = exp(-λ * months_ago)，λ≈0.2 → 半衰期约 3.5 月
TIME_DECAY_LAMBDA = 0.2
# 斜率池 MAD 离群过滤（样本数不足时跳过）
SLOPE_MAD_K = 3.5
SLOPE_MAD_MIN_SAMPLES = 4
SLOPE_POOL_MAX = 5000
# 传导系数 (元/kg)/(元/件) 默认裁剪；有重量时上界随重量收紧
CONDUCTION_COEFF_ABS_MIN = 0.005
CONDUCTION_COEFF_ABS_MAX = 0.15
# 低价值物料：不参与超区间扣分 / σ 冲击
LOW_VALUE_UNIT_PRICE_YUAN = 0.5
MIN_BASE_AMOUNT_YUAN = 1.0
INVALID_BOM_DATE = "00000000"
MIN_VALID_BOM_DATE = "19000101"
# 价格清单：数值明显大于典型「每公斤」报价时推断为元/件（可用环境变量覆盖）
def _per_piece_total_threshold_yuan() -> float:
    try:
        return float(os.environ.get("PER_PIECE_TOTAL_THRESHOLD_YUAN", "12"))
    except ValueError:
        return 12.0


PER_PIECE_TOTAL_THRESHOLD_YUAN = _per_piece_total_threshold_yuan()

UnitHint = Literal["piece", "kg", "unknown"]

_PIECE_UNIT_TOKENS = frozenset(
    {"PC", "ST", "EA", "SET", "PCS", "PCE", "件", "个", "只", "套", "台"}
)
_KG_UNIT_TOKENS = frozenset(
    {"KG", "KGM", "千克", "公斤", "G", "克", "GRAM", "GR", "GM"}
)
_UNIT_HINT_COLUMN_KEYS = ("MEINS", "基本单位", "计量单位", COL_UNIT)
# Shock σ：非核心件（行成本占基准 BOM 低于此比例）豁免封顶扣分
SHOCK_NON_CORE_BOM_SHARE = 0.01
SHOCK_MIN_EFFECTIVE_STD_YUAN = 0.05
SHOCK_ABS_DELTA_CAP_YUAN = 0.5
# 单价低于此值时，σ 封顶链改用相对涨幅阈值（辅料微量波动）
LOW_VALUE_REL_SIGMA_UNIT_YUAN = 2.0
LOW_VALUE_REL_SIGMA_RATIO = 3.0
# 相对加载基准倍数 R：分段可信度上限（见 _credibility_cap_from_ref_multiple）
CREDIBILITY_REF_NORMAL_MAX = 1.1
CREDIBILITY_REF_ABNORMAL_MIN = 5.0
CREDIBILITY_REF_FLOOR_CAP = 15
# (R 下界, R 上界, 区间起始分, 区间结束分, 提示文案；下界不含、上界含)
CREDIBILITY_REF_BANDS: list[tuple[float, float, int, int, str]] = [
    (1.1, 1.3, 100, 90, "价格略微偏离加载基准价（当前 {R} 倍）"),
    (1.3, 1.6, 90, 75, "价格明显偏离加载基准价（当前 {R} 倍）"),
    (1.6, 2.0, 75, 55, "价格严重偏离，请核对单位（当前 {R} 倍）"),
    (2.0, 3.5, 55, 35, "价格存在极大偏差（当前 {R} 倍）"),
    (3.5, 5.0, 35, 15, "价格极度异常，可信度极低（当前 {R} 倍）"),
]
# 模型历史误差区间：随模拟 BOM 偏离基准而放大
MODEL_ERROR_GAMMA = 0.5

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
    """将「报价月份」或日期列规范为 YYYYMM；优先 pd.to_datetime，再回退 calendar_date_key。"""
    raw = norm_material_code(s)
    ts = pd.to_datetime(raw, errors="coerce", dayfirst=False)
    out = pd.Series("", index=s.index, dtype=object)
    valid = ts.notna()
    if valid.any():
        out.loc[valid] = ts.loc[valid].dt.strftime("%Y%m")

    def _fallback(v: str) -> str:
        if not v:
            return ""
        digits = re.sub(r"\D", "", v)
        if len(digits) >= 6:
            return digits[:6]
        if len(digits) == 4:
            return digits
        dk = calendar_date_key_yyyymmdd(v)
        return dk[:6] if dk and len(dk) >= 6 else ""

    miss = ~valid
    if miss.any():
        out.loc[miss] = raw.loc[miss].map(_fallback)
    return out.astype(str)


def quote_month_key_scalar(v: Any) -> str:
    """单值报价月份 → YYYYMM（与 quote_month_key_from_series 一致）。"""
    if v is None:
        return ""
    ser = quote_month_key_from_series(pd.Series([v]))
    return str(ser.iloc[0]) if len(ser) else ""


def quote_month_key_from_value(v: Any) -> str:
    return quote_month_key_scalar(v)


def normalize_predict_dataframes(
    bom_df: pd.DataFrame | None,
    price_df: pd.DataFrame | None,
) -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    """预测入口一次性规范化物料编码与报价月份，避免热路径反复构造 Series。"""
    if bom_df is not None and not bom_df.empty:
        bom_df = bom_df.copy()
        if COL_PRODUCT in bom_df.columns:
            bom_df[COL_PRODUCT] = norm_material_code(bom_df[COL_PRODUCT])
        if COL_COMPONENT in bom_df.columns:
            bom_df[COL_COMPONENT] = norm_material_code(bom_df[COL_COMPONENT])
        if COL_CREATEDATE in bom_df.columns:
            bom_df[COL_CREATEDATE] = bom_df[COL_CREATEDATE].map(calendar_date_key_yyyymmdd)
    if price_df is not None and not price_df.empty:
        price_df = price_df.copy()
        pid_col = COL_PRODUCT if COL_PRODUCT in price_df.columns else None
        if pid_col is None:
            for alias in ("ZMATNR", "MATNR", "所属产品"):
                if alias in price_df.columns:
                    pid_col = alias
                    break
        if pid_col:
            price_df[pid_col] = norm_material_code(price_df[pid_col])
        if "报价月份" in price_df.columns:
            price_df["报价月份"] = quote_month_key_from_series(price_df["报价月份"])
    return bom_df, price_df


def build_product_category_index(bom_df: pd.DataFrame | pl.DataFrame) -> dict[str, str]:
    """
    按 BOM 行成本（MENGE×组件单价）汇总，取贡献最大的组件「分类」作为产品 dominant 分类。
    """
    if isinstance(bom_df, pl.DataFrame):
        df = normalize_columns(bom_df)
    else:
        if bom_df is None or (hasattr(bom_df, "empty") and bom_df.empty):
            return {}
        df = normalize_columns(pl.from_pandas(bom_df))

    if COL_PRODUCT not in df.columns or COL_CATEGORY not in df.columns:
        return {}

    qty = pl.col(COL_QTY).cast(pl.Float64, strict=False).fill_null(0.0) if COL_QTY in df.columns else pl.lit(0.0)
    price = (
        pl.col(COL_UNIT_PRICE).cast(pl.Float64, strict=False).fill_null(0.0)
        if COL_UNIT_PRICE in df.columns
        else pl.lit(0.0)
    )
    ranked = (
        df.with_columns(
            norm_material_code_expr(COL_PRODUCT).alias(COL_PRODUCT),
            (qty * price).alias("_line_cost"),
        )
        .filter(
            pl.col(COL_PRODUCT).is_not_null()
            & (pl.col(COL_PRODUCT) != "")
            & pl.col(COL_CATEGORY).is_not_null()
            & (pl.col(COL_CATEGORY).cast(pl.Utf8) != "")
        )
        .group_by([COL_PRODUCT, COL_CATEGORY])
        .agg(pl.col("_line_cost").sum().alias("_cat_cost"))
        .sort([COL_PRODUCT, "_cat_cost"], descending=[False, True])
    )
    top = ranked.group_by(COL_PRODUCT).first()
    return {
        str(r[COL_PRODUCT]): str(r[COL_CATEGORY])
        for r in top.to_dicts()
        if r.get(COL_PRODUCT) and r.get(COL_CATEGORY)
    }


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


def normalize_unit_hint(unit: Any) -> UnitHint:
    """将 MEINS / 基本单位等规范为 piece | kg | unknown。"""
    if unit is None:
        return "unknown"
    s = str(unit).strip().upper()
    if not s or s in ("NAN", "NONE", "<NA>", ""):
        return "unknown"
    if s in _PIECE_UNIT_TOKENS:
        return "piece"
    if s in _KG_UNIT_TOKENS:
        return "kg"
    if any(tok in s for tok in ("件", "个", "只")):
        return "piece"
    if any(tok in s for tok in ("千克", "公斤", "KG")):
        return "kg"
    return "unknown"


def _latest_valid_bom_date(dates: "pd.Series | pl.Series") -> str:
    """忽略 00000000 / 无效日期后取最大 8 位日期；全无有效日期时退回全表 max。"""
    if hasattr(dates, "to_pandas"):
        dates = dates.to_pandas()
    s = dates.astype(str).str.strip().str[:8]
    valid = s[(s > MIN_VALID_BOM_DATE) & (s != INVALID_BOM_DATE)]
    if len(valid):
        return str(valid.max())
    return str(s.max()) if len(s) else INVALID_BOM_DATE


def unit_hint_from_row(row: Any) -> UnitHint:
    """从价格快照 dict 或 pandas Series 读取计量单位倾向。"""
    if row is None:
        return "unknown"
    for key in _UNIT_HINT_COLUMN_KEYS:
        if isinstance(row, dict):
            if key in row and row[key] not in (None, ""):
                return normalize_unit_hint(row[key])
        elif hasattr(row, "index") and key in row.index:
            v = row.get(key) if hasattr(row, "get") else row[key]
            if v is not None and not (isinstance(v, float) and math.isnan(v)):
                if str(v).strip():
                    return normalize_unit_hint(v)
    return "unknown"


def unit_hint_from_bom_dataframe(
    bom_df: pd.DataFrame | None,
    product_id: str,
) -> UnitHint:
    """按最新 BOM 行成本加权，推断成品计量单位倾向（PC vs KG）。"""
    if bom_df is None or bom_df.empty or COL_PRODUCT not in bom_df.columns:
        return "unknown"
    pid = norm_material_code_scalar(product_id)
    sub = bom_df[norm_material_code(bom_df[COL_PRODUCT]) == pid].copy()
    if sub.empty:
        return "unknown"
    unit_col = next((c for c in _UNIT_HINT_COLUMN_KEYS if c in sub.columns), None)
    if not unit_col:
        return "unknown"
    if COL_CREATEDATE in sub.columns:
        sub["_date"] = norm_material_code(sub[COL_CREATEDATE]).str[:8]
        latest = _latest_valid_bom_date(sub["_date"])
        sub = sub[sub["_date"] == latest]
    qty = pd.to_numeric(
        sub[COL_QTY] if COL_QTY in sub.columns else 0, errors="coerce"
    ).fillna(0.0)
    price = pd.to_numeric(
        sub[COL_UNIT_PRICE] if COL_UNIT_PRICE in sub.columns else 0, errors="coerce"
    ).fillna(0.0)
    weights = qty * price
    votes: dict[str, float] = {"piece": 0.0, "kg": 0.0}
    for u, w in zip(sub[unit_col], weights):
        hint = normalize_unit_hint(u)
        if hint in votes and w > 0:
            votes[hint] += float(w)
    if votes["piece"] > votes["kg"] * 1.05:
        return "piece"
    if votes["kg"] > votes["piece"] * 1.05:
        return "kg"
    return "unknown"


def _resolve_cost_unit_hint(
    unit_hint: UnitHint | str | None,
) -> UnitHint:
    if unit_hint in ("piece", "kg", "unknown"):
        return unit_hint  # type: ignore[return-value]
    return normalize_unit_hint(unit_hint)


def _ref_multiple_magnitude(ref_mult: float) -> float | None:
    """相对加载基准的偏离倍数 R（涨价为 ref_mult，跌价为 1/ref_mult）。"""
    if not math.isfinite(ref_mult) or ref_mult <= 0:
        return None
    return ref_mult if ref_mult >= 1.0 else 1.0 / ref_mult


def _format_ref_multiple_label(r: float) -> str:
    if r >= 100:
        return f"{r:.1f}"
    if r >= 10:
        return f"{r:.2f}"
    return f"{r:.2f}"


def _credibility_cap_from_ref_multiple(
    ref_mult: float,
) -> tuple[int | None, str | None, str | None]:
    """
    相对「加载 BOM 时单价」的倍数 R → (可信度上限, 提示文案, R 展示串)。

    R≤1.1：不扣分；1.1～5 分段线性插值；R>5：保底 15 分。
    """
    r = _ref_multiple_magnitude(ref_mult)
    if r is None:
        return None, None, None
    r_label = _format_ref_multiple_label(r)
    if r <= CREDIBILITY_REF_NORMAL_MAX:
        return None, None, None
    if r > CREDIBILITY_REF_ABNORMAL_MIN:
        msg = f"价格已脱离实际市场参考（当前 {r_label} 倍）"
        return CREDIBILITY_REF_FLOOR_CAP, msg, r_label
    for r_lo, r_hi, cap_lo, cap_hi, template in CREDIBILITY_REF_BANDS:
        if r <= r_hi:
            span = r_hi - r_lo
            t = (r - r_lo) / span if span > 0 else 1.0
            cap = round(cap_lo + t * (cap_hi - cap_lo))
            return cap, template.format(R=r_label), r_label
    return CREDIBILITY_REF_FLOOR_CAP, (
        f"价格已脱离实际市场参考（当前 {r_label} 倍）"
    ), r_label


def _price_list_values_imply_per_piece(
    sum_ml: float,
    total_f: float | None,
    wt: float | None,
) -> bool:
    """
    根据清单数值判断是否为「整件报价」（元/件）。
    优先于 BOM 行上的 KG 单位提示——行单位是原料计量，成品价仍是元/件。
    """
    threshold = _per_piece_total_threshold_yuan()
    if wt is None or wt <= 1e-12 or total_f is None or total_f <= 1e-12:
        return False
    if abs(sum_ml - total_f) / max(sum_ml, 1e-9) >= 0.06:
        return False
    if not (15.0 <= sum_ml <= threshold * 5):
        return False
    # 整件报价时「材料+工费」/重量 常为个位数元/kg；清单若已是元/kg 则往往 ≥ 阈值
    if sum_ml / float(wt) < threshold:
        return True
    return False


def _infer_costs_are_per_piece(
    sum_ml: float,
    total_f: float | None,
    wt: float | None,
    unit_hint: UnitHint,
) -> bool:
    """在材料+工费可用时，判定清单口径为元/件还是元/KG。"""
    if _price_list_values_imply_per_piece(sum_ml, total_f, wt):
        return True
    if unit_hint == "piece":
        return True
    if unit_hint == "kg":
        return False

    threshold = _per_piece_total_threshold_yuan()
    has_wt = wt is not None and wt > 1e-12
    if has_wt:
        if total_f is not None and total_f > 1e-12:
            rel_sum_total = abs(sum_ml - total_f) / max(sum_ml, 1e-9)
            if rel_sum_total < 0.06:
                return sum_ml > threshold
            if abs(sum_ml / wt - total_f) / max(total_f, 1e-9) < 0.08:
                return True
            return False
        return sum_ml > threshold
    return sum_ml > threshold


def normalize_price_list_cost_fields(
    mat: float | None,
    labor: float | None,
    total: float | None,
    weight_kg: float | None,
    *,
    unit_hint: UnitHint | str | None = None,
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
    uh = _resolve_cost_unit_hint(unit_hint)
    per_piece = _infer_costs_are_per_piece(sum_ml, total_f, wt, uh)
    # 兜底：1kg 轻量件 sum_ml≈元/kg，误判为元/件时按元/kg 重算
    if per_piece and wt and wt > 1e-12:
        probe_per_kg = (
            (total_f / wt)
            if total_f is not None and total_f > 1e-12
            else sum_ml / wt
        )
        if float(probe_per_kg) > _per_piece_total_threshold_yuan() * 1.5:
            per_piece = False

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


def baseline_prices_from_price_snap(
    price_snap: dict,
    weight_kg: float | None,
) -> dict[str, Any]:
    """
    从价格清单快照解析基准整件价(元/件)与基准价(元/kg)。
    供 Web 映射层与预测模块共用，避免 app 内重复口径逻辑。
    """
    wt = _safe_float(weight_kg)
    if wt is None:
        wt = _safe_float(price_snap.get("重量"))
    norm = normalize_price_list_cost_fields(
        _safe_float(price_snap.get("材料成本")),
        _safe_float(price_snap.get("工费")),
        _safe_float(price_snap.get("总成本")),
        wt,
        unit_hint=unit_hint_from_row(price_snap),
    )
    piece = norm.get("total_piece")
    per_kg = norm.get("total_per_kg")
    raw_total = _safe_float(price_snap.get("总成本"))
    if piece is None and raw_total is not None:
        if norm.get("costs_are_per_piece"):
            piece = raw_total
            per_kg = raw_total / wt if wt and wt > 0 else None
        elif wt and wt > 0:
            per_kg = raw_total
            piece = raw_total * wt
        else:
            piece = raw_total
            per_kg = raw_total
    list_unit = "元/件" if norm.get("costs_are_per_piece") else "元/KG"
    return {
        "基准产品价格": round(piece, 4) if piece is not None else None,
        "基准产品价格_每公斤": round(per_kg, 4) if per_kg is not None else None,
        "清单材料工费口径": list_unit,
    }


def dedupe_warnings(warnings: list[str]) -> list[str]:
    """按完整文案去重，保留首次出现顺序。"""
    return list(dict.fromkeys(warnings))


def _infer_total_per_kg_when_mat_labor_missing(
    total_f: float,
    weight_kg: float | None,
    *,
    unit_hint: UnitHint | str | None = None,
) -> float | None:
    """
    仅总成本列可用时推断 元/kg：与 normalize 共用「>阈值≈元/件」启发式。
    无有效重量且 total 像整件价时返回 None，避免回测量纲混乱。
    """
    if total_f <= 1e-12:
        return None
    wt = _safe_float(weight_kg)
    has_wt = wt is not None and wt > 1e-12
    uh = _resolve_cost_unit_hint(unit_hint)
    threshold = _per_piece_total_threshold_yuan()
    if uh == "kg":
        return float(total_f)
    if uh == "piece":
        if has_wt:
            return float(total_f) / float(wt)
        return None
    if has_wt:
        if total_f > threshold:
            return float(total_f) / float(wt)
        return float(total_f)
    if total_f > threshold:
        return None
    return float(total_f)


def actual_price_per_kg_from_price_row(
    mat: Any,
    labor: Any,
    total: Any,
    weight_kg: Any,
    *,
    unit_hint: UnitHint | str | None = None,
) -> float | None:
    """将价格清单一行中的材料+工费/总成本统一为 元/kg（与主预测量纲一致）。"""
    mat_f = _safe_float(mat)
    labor_f = _safe_float(labor)
    total_f = _safe_float(total)
    wt = _safe_float(weight_kg)
    uh = _resolve_cost_unit_hint(unit_hint)
    if mat_f is not None and labor_f is not None:
        norm = normalize_price_list_cost_fields(
            mat_f, labor_f, total_f, wt, unit_hint=uh
        )
        per_kg = norm.get("total_per_kg")
        if per_kg is not None and per_kg > 1e-12:
            return float(per_kg)
    if total_f is not None and total_f > 1e-12:
        inferred = _infer_total_per_kg_when_mat_labor_missing(
            total_f, wt, unit_hint=uh
        )
        if inferred is not None and inferred > 1e-12:
            return inferred
    return None


def _cost_structure_bases(base_data: dict) -> tuple[float, float, bool]:
    """按清单口径取材料/工费基准，避免元/件字段误入元/kg 分支。"""
    per_piece = bool(base_data.get("costs_are_per_piece"))
    if per_piece:
        mat = base_data.get("base_mat_piece")
        labor = base_data.get("base_labor_piece")
        if mat is None:
            mat = base_data.get("base_mat")
        if labor is None:
            labor = base_data.get("base_labor")
    else:
        mat = base_data.get("base_mat_per_kg")
        labor = base_data.get("base_labor_per_kg")
        if mat is None:
            mat = base_data.get("base_mat")
        if labor is None:
            labor = base_data.get("base_labor")
    return float(mat), float(labor), per_piece


def _sum_bom_from_latest_bom(
    lb: pd.DataFrame,
    unit_for_row: Any,
) -> float:
    """与预测循环一致：对 latest_bom 各行 用量×单价 求和。"""
    total = 0.0
    for _, row in lb.iterrows():
        qty = float(row.get("组件数量") or 0.0)
        if qty <= 0:
            continue
        mid = str(row.get("材料编码") or "").strip()
        unit = float(unit_for_row(mid, row))
        total += qty * unit
    return total


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
    """从原始上传表按 产品[+日期] 汇总：BOM 材料合计、产品总价、材料成本、工费（Polars）。"""
    df = normalize_columns(raw_pl)
    if COL_PRODUCT not in df.columns or COL_COMPONENT not in df.columns:
        return {}

    qty = (
        pl.col(COL_QTY).cast(pl.Float64, strict=False).fill_null(0.0)
        if COL_QTY in df.columns
        else pl.lit(0.0)
    )
    price = (
        pl.col(COL_UNIT_PRICE).cast(pl.Float64, strict=False).fill_null(0.0)
        if COL_UNIT_PRICE in df.columns
        else pl.lit(0.0)
    )

    if COL_CREATEDATE in df.columns:
        # 向量化替代 map_elements(calendar_date_key_yyyymmdd)：
        # 1. cast→String 兼容 Int/Date/Datetime/String 列
        # 2. 去掉末尾 ".0"（整数被 pandas 读为浮点时出现）
        # 3. 去掉所有非数字字符（处理 "2026-01-07" 等带分隔符格式）
        # 4. 取前 8 位；长度不足 8 则回退 "snapshot"
        dk_expr = (
            pl.col(COL_CREATEDATE)
            .cast(pl.String)
            .str.strip_chars()
            .str.replace(r"\.0$", "")
            .str.replace_all(r"[^\d]", "")
            .str.slice(0, 8)
            .pipe(
                lambda e: pl.when(e.str.len_chars() == 8)
                .then(e)
                .otherwise(pl.lit("snapshot"))
            )
            .alias("_dk")
        )
    else:
        dk_expr = pl.lit("snapshot").alias("_dk")

    quote_col = next(
        (c for c in ("报价号", "ZBJNO", "报价流水号", "ZSNO") if c in df.columns),
        None,
    )
    if quote_col:
        qv_expr = norm_material_code_expr(quote_col).alias("_qv")
    else:
        qv_expr = pl.lit(None).cast(pl.Utf8).alias("_qv")

    base = df.with_columns(
        norm_material_code_expr(COL_PRODUCT).alias(COL_PRODUCT),
        norm_material_code_expr(COL_COMPONENT).alias(COL_COMPONENT),
        (qty * price).alias("_line"),
        dk_expr,
        qv_expr,
    )

    version_daily = base.group_by([COL_PRODUCT, "_qv", "_dk"]).agg(
        pl.col("_line").sum().alias("_line")
    )
    # 同日多报价版本：取排序后最后一条（最新报价），避免对版本做均值稀释
    bom_by_day = (
        version_daily.sort([COL_PRODUCT, "_dk", "_qv"], nulls_last=True)
        .group_by([COL_PRODUCT, "_dk"])
        .agg(pl.col("_line").last().alias("bom_material"))
    )

    meta_cols = [c for c in ("产品价格", "材料成本", "工费", "重量") if c in base.columns]
    sort_keys = [COL_PRODUCT, "_dk"]
    if COL_CREATEDATE in base.columns:
        sort_keys.append(COL_CREATEDATE)
    meta = base.sort(sort_keys).group_by([COL_PRODUCT, "_dk"]).last()
    if meta_cols:
        meta = meta.select([COL_PRODUCT, "_dk", *meta_cols])
    else:
        meta = meta.select([COL_PRODUCT, "_dk"])

    joined = bom_by_day.join(meta, on=[COL_PRODUCT, "_dk"], how="left")

    out: dict[str, list[dict[str, Any]]] = {}
    for rec in joined.to_dicts():
        pid = str(rec.get(COL_PRODUCT) or "").strip()
        if not pid:
            continue
        dk = rec.get("_dk")
        bom_material = float(rec.get("bom_material") or 0.0)
        if bom_material <= 1e-12:
            continue

        wt_kg: float | None = None
        if rec.get("重量") is not None:
            try:
                w = float(rec["重量"])
                if w > 1e-12:
                    wt_kg = w
            except (TypeError, ValueError):
                pass

        pp_per_kg: float | None = None
        mat_per_kg: float | None = None
        labor_per_kg: float | None = None
        if rec.get("产品价格") is not None:
            try:
                pp_per_kg = float(rec["产品价格"])
            except (TypeError, ValueError):
                pass
        if rec.get("材料成本") is not None:
            try:
                mat_per_kg = float(rec["材料成本"])
            except (TypeError, ValueError):
                pass
        if rec.get("工费") is not None:
            try:
                labor_per_kg = float(rec["工费"])
            except (TypeError, ValueError):
                pass

        if pp_per_kg is None and mat_per_kg is not None and labor_per_kg is not None:
            pp_per_kg = mat_per_kg + labor_per_kg
        if pp_per_kg is None:
            continue

        row: dict[str, Any] = {
            "date": str(dk) if dk is not None else "snapshot",
            "bom_material": round(bom_material, 6),
            "product_price_per_kg": round(pp_per_kg, 6),
            "product_price": round(
                per_kg_to_product_total(pp_per_kg, wt_kg) or pp_per_kg, 6
            ),
        }
        if mat_per_kg is not None:
            row["材料成本_per_kg"] = round(mat_per_kg, 6)
        if labor_per_kg is not None:
            row["工费_per_kg"] = round(labor_per_kg, 6)
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
    price = read_price_excel(price_path)
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
    product_price_timeline: dict[str, dict[str, dict[str, float]]] | None,
) -> dict[str, list[dict[str, Any]]]:
    """
    用产品价格时间线补齐「日期 → BOM材料合计 + 产品价」样本。
    BOM 材料合计使用当前 BOM 默认单价 × 用量（报价材料价格趋势已停用）。
    """
    ppt = product_price_timeline or {}
    if not ppt:
        return dict(product_cost_history)

    all_products = set(product_bom.keys()) | set(product_cost_history.keys()) | set(ppt.keys())
    out: dict[str, list[dict[str, Any]]] = dict(product_cost_history)

    static_bom: dict[str, float] = {}
    for product, bom_rows in product_bom.items():
        if not bom_rows:
            continue
        total = 0.0
        for r in bom_rows:
            c = str(r.get("组件编码") or "").strip()
            if not c:
                continue
            qty = float(r.get("MENGE合计") or 0)
            if qty <= 0:
                continue
            total += qty * float(r.get("组件单价") or 0)
        if total > 1e-12:
            static_bom[product] = total

    for product in all_products:
        bom_material_static = static_bom.get(product)
        if not bom_material_static:
            continue

        existing = {
            _timeline_month_key(str(h.get("date"))): h for h in out.get(product, [])
        }
        pp_tl = {
            _timeline_month_key(k): v for k, v in (ppt.get(product, {}) or {}).items()
        }

        dates: set[str] = set(existing.keys()) | set(pp_tl.keys())
        if not dates:
            continue

        expanded_rows: list[dict[str, Any]] = []
        for mk in sorted(dates):
            if mk in ("", "nan", "None"):
                continue
            bom_material = bom_material_static

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
    if r2 < 0:
        return "无效"
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

    if float(np.std(x)) < 1e-4 or float(np.ptp(x)) < 1e-4:
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
    r2 = min(1.0, float(r2))
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


def _month_index_from_key(dk: str) -> int | None:
    mk = _timeline_month_key(dk)
    if not mk or mk in ("snapshot", "nan", "None"):
        return None
    digits = re.sub(r"\D", "", mk)
    if len(digits) < 6:
        return None
    try:
        return int(digits[:4]) * 12 + int(digits[4:6])
    except ValueError:
        return None


def _weighted_median(values: list[float], weights: list[float]) -> float:
    if not values:
        return float("nan")
    if not weights or len(weights) != len(values) or sum(weights) <= 0:
        return float(np.nanmedian(np.array(values, dtype=float)))
    pairs = sorted(zip(values, weights), key=lambda x: x[0])
    half = sum(weights) / 2.0
    cum = 0.0
    for v, w in pairs:
        cum += w
        if cum >= half:
            return float(v)
    return float(pairs[-1][0])


def _aggregate_slopes(
    slopes: list[float],
    slope_weights: list[float] | None,
) -> float:
    if slope_weights and len(slope_weights) == len(slopes):
        return _weighted_median(slopes, slope_weights)
    return float(np.nanmedian(np.array(slopes, dtype=float)))


def _filter_slopes_mad(
    slopes: list[float],
    weights: list[float] | None = None,
) -> tuple[list[float], list[float] | None]:
    """中位数绝对偏差（MAD）自适应剔除斜率离群值；样本过少时不过滤。"""
    if len(slopes) < SLOPE_MAD_MIN_SAMPLES:
        return slopes, weights
    arr = np.array(slopes, dtype=float)
    med = float(np.median(arr))
    mad = float(np.median(np.abs(arr - med)))
    if mad / max(abs(med), 1e-9) < 1e-6:
        return slopes, weights
    keep = np.abs(arr - med) <= SLOPE_MAD_K * mad
    if not keep.any():
        return slopes, weights
    new_slopes = [s for s, k in zip(slopes, keep) if k]
    if weights and len(weights) == len(slopes):
        new_w = [w for w, k in zip(weights, keep) if k]
        return new_slopes, new_w if new_w else None
    return new_slopes, None


def _robust_slopes_from_points(
    points: list[tuple[str, float, float]],
) -> tuple[list[float], list[float] | None]:
    """
    Theil-Sen 风格斜率样本：用所有成对变化估计产品价对 BOM 材料变化的传导。
    返回 (slopes, weights)；若时间键可解析则对较新样本对加权（指数衰减）。
    """
    n = len(points)
    if n < 2:
        return [], None
    xs = np.array([p[1] for p in points], dtype=float)
    dx_floor = max(1e-9, float(np.nanmedian(np.abs(xs))) * 1e-5)
    month_idxs = [_month_index_from_key(p[0]) for p in points]
    use_decay = all(m is not None for m in month_idxs)
    latest_mi = max(month_idxs) if use_decay else None

    slopes: list[float] = []
    weights: list[float] = []
    for i in range(n - 1):
        xi = points[i][1]
        yi = points[i][2]
        for j in range(i + 1, n):
            dx = points[j][1] - xi
            if abs(dx) <= dx_floor:
                continue
            s = (points[j][2] - yi) / dx
            if not math.isfinite(s):
                continue
            slopes.append(float(s))
            if use_decay and latest_mi is not None:
                mi_i, mi_j = month_idxs[i], month_idxs[j]
                if mi_i is not None and mi_j is not None:
                    avg_months_ago = (
                        (latest_mi - mi_i) + (latest_mi - mi_j)
                    ) / 2.0
                    weights.append(
                        math.exp(-TIME_DECAY_LAMBDA * max(0.0, avg_months_ago))
                    )
                elif month_idxs[j] is not None:
                    months_ago = max(0, latest_mi - month_idxs[j])
                    weights.append(math.exp(-TIME_DECAY_LAMBDA * months_ago))
                else:
                    weights.append(1.0)
            elif use_decay:
                weights.append(1.0)

    w_out: list[float] | None = weights if use_decay and len(weights) == len(slopes) else None
    return _filter_slopes_mad(slopes, w_out)


def _pass_through_from_records(
    records: list[dict[str, Any]],
    source: str,
) -> dict[str, Any] | None:
    points = _history_xy_points(records)
    slopes, slope_weights = _robust_slopes_from_points(points)
    if not slopes:
        return None
    raw = _aggregate_slopes(slopes, slope_weights)
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


def _conduction_coeff_clip_bounds(weight_kg: float | None) -> tuple[float, float]:
    """按产品重量收紧传导系数 (元/kg)/(元/件) 的裁剪区间。"""
    lo, hi = CONDUCTION_COEFF_ABS_MIN, CONDUCTION_COEFF_ABS_MAX
    if weight_kg is not None and weight_kg > 1e-12:
        w = float(weight_kg)
        hi = min(hi, 2.0 / w)
        lo = max(lo, min(0.02, 0.05 / w))
        if lo > hi:
            lo = hi * 0.5
    return lo, hi


def _reservoir_add_slopes(
    pooled: list[float],
    pooled_w: list[float],
    new_slopes: list[float],
    new_w: list[float] | None,
    *,
    max_size: int,
    seen: int,
) -> int:
    """将新斜率流式并入池；超过 max_size 时用 reservoir 抽样，避免偏向前序产品。"""
    paired_w = bool(new_w) and len(new_w) == len(new_slopes)
    for idx, s in enumerate(new_slopes):
        seen += 1
        wv = float(new_w[idx]) if paired_w else None
        if len(pooled) < max_size:
            pooled.append(s)
            if paired_w and wv is not None:
                pooled_w.append(wv)
        else:
            j = random.randint(0, seen - 1)
            if j < max_size:
                pooled[j] = s
                if paired_w and wv is not None and j < len(pooled_w):
                    pooled_w[j] = wv
    return seen


def _pool_slopes_from_history(
    all_history: dict[str, list[dict]] | None,
    *,
    product_filter: Any | None = None,
) -> tuple[list[float], list[float], int, int]:
    """汇总斜率池；product_filter(pid)->bool，None 表示全部。"""
    pooled: list[float] = []
    pooled_w: list[float] = []
    n_points = 0
    n_products = 0
    seen = 0
    for pid, records in (all_history or {}).items():
        if product_filter is not None and not product_filter(pid):
            continue
        pts = _history_xy_points(records)
        if len(pts) < 2:
            continue
        n_products += 1
        n_points += len(pts)
        slopes, w = _robust_slopes_from_points(pts)
        if not slopes:
            continue
        seen = _reservoir_add_slopes(
            pooled,
            pooled_w,
            slopes,
            w,
            max_size=SLOPE_POOL_MAX,
            seen=seen,
        )
    pooled, pooled_w = _filter_slopes_mad(pooled, pooled_w if pooled_w else None)
    return pooled, pooled_w, n_points, n_products


def _pass_through_from_pooled_slopes(
    slopes: list[float],
    slope_weights: list[float],
    *,
    source: str,
    sample_count: int,
    n_products: int,
) -> dict[str, Any] | None:
    if not slopes:
        return None
    w = slope_weights if len(slope_weights) == len(slopes) else None
    raw = _aggregate_slopes(slopes, w)
    coef = min(PASS_THROUGH_MAX, max(PASS_THROUGH_MIN, raw))
    return {
        "coefficient": coef,
        "raw_coefficient": raw,
        "source": source,
        "sample_count": sample_count,
        "slope_count": len(slopes),
        "product_count": n_products,
        "spread": (
            float(np.nanpercentile(slopes, 75) - np.nanpercentile(slopes, 25))
            if len(slopes) >= 4
            else 0.0
        ),
    }


def _estimate_pass_through(
    product: str,
    product_records: list[dict[str, Any]],
    all_history: dict[str, list[dict]] | None,
    *,
    product_categories: dict[str, str] | None = None,
) -> dict[str, Any] | None:
    own = _pass_through_from_records(product_records, "本产品历史")
    if own is not None and own["slope_count"] >= 2:
        return own

    cat = (product_categories or {}).get(product) if product_categories else None
    if cat:

        def _same_cat(pid: str) -> bool:
            return pid != product and (product_categories or {}).get(pid) == cat

        cat_slopes, cat_w, cat_pts, cat_prods = _pool_slopes_from_history(
            all_history, product_filter=_same_cat
        )
        if (
            len(cat_slopes) >= PASS_THROUGH_CATEGORY_MIN_SLOPES
            or cat_prods >= PASS_THROUGH_CATEGORY_MIN_PRODUCTS
        ):
            info = _pass_through_from_pooled_slopes(
                cat_slopes,
                cat_w,
                source="同分类产品历史",
                sample_count=cat_pts,
                n_products=cat_prods,
            )
            if info is not None:
                info["category"] = cat
                return info

    global_slopes, global_w, global_points, global_prods = _pool_slopes_from_history(all_history)
    if global_slopes:
        info = _pass_through_from_pooled_slopes(
            global_slopes,
            global_w,
            source="全局产品历史",
            sample_count=global_points,
            n_products=global_prods,
        )
        if info is not None:
            return info

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
        if p_old < LOW_VALUE_UNIT_PRICE_YUAN:
            continue

        lo, hi = ranges.get(comp, (p_old, p_old))
        lo = max(float(lo) if lo else p_old, LOW_VALUE_UNIT_PRICE_YUAN, p_old)
        hi = max(float(hi) if hi else p_old, LOW_VALUE_UNIT_PRICE_YUAN, p_old)
        if p_new > hi:
            factors.append(float(p_new) / max(hi, LOW_VALUE_UNIT_PRICE_YUAN))
        elif p_new < lo:
            factors.append(lo / max(float(p_new), LOW_VALUE_UNIT_PRICE_YUAN))
        else:
            factors.append(1.0)

    material_factor = 1.0
    base_mat_eff = max(baseline_material, MIN_BASE_AMOUNT_YUAN)
    sim_mat_eff = max(simulated_material, MIN_BASE_AMOUNT_YUAN)
    if baseline_material > 1e-12 and simulated_material > 1e-12:
        ratio = sim_mat_eff / base_mat_eff
        material_factor = max(ratio, 1.0 / ratio)
        factors.append(material_factor)

    max_factor = max(factors) if factors else 1.0
    cap: float | None = None
    if max_factor > 1.5:
        cap = max(20.0, 70.0 - 15.0 * (max_factor - 1.5))

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
    product_categories: dict[str, str] | None = None,
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
        unit_hint=unit_hint_from_row(price_snap),
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
    pass_info = _estimate_pass_through(
        product,
        hist_valid,
        product_cost_history,
        product_categories=product_categories,
    )
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
                src_pt = pass_info.get("source")
                w = (
                    0.98
                    if src_pt == "本产品历史"
                    else (0.92 if src_pt == "同分类产品历史" else 0.9)
                )
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
        src = pass_info.get("source")
        if src == "本产品历史":
            pass_score = min(16.0, 8.0 + min(8.0, float(pass_info.get("slope_count") or 0)))
        elif src == "同分类产品历史":
            pass_score = min(14.0, 7.0 + min(7.0, float(pass_info.get("slope_count") or 0) / 12.0))
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
        "enabled": reg_enabled,
        "used_as_primary": reg_used,
        "valid_sample_count": n_hist if reg_enabled else 0,
        "bom_direct_sample_count": n_raw,
        "r2": round(r2, 4) if reg_enabled else None,
        "r2_pct": round(r2 * 100.0, 1) if reg_enabled else None,
        "fit_grade": reg_grade,
        "model": "product_price_per_kg ~ intercept + slope * bom_material_per_kg",
        "intercept": round(reg_intercept, 6) if reg_intercept is not None else None,
        "slope": round(reg_slope, 6) if reg_slope is not None else None,
        "sim_bom_material_per_kg": round(sim_per_kg, 6) if sim_per_kg is not None else None,
        "reg_predicted_product_price_per_kg": (
            round(pred_reg_per_kg, 6) if pred_reg_per_kg is not None else None
        ),
        "pass_through_coefficient": round(pass_coef, 6) if pass_coef is not None else None,
        "pass_through_source": pass_info.get("source") if pass_info is not None else None,
        "pass_predicted_product_price_per_kg": (
            round(pred_pass_per_kg, 6) if pred_pass_per_kg is not None else None
        ),
        "note": (
            f"共 {n_hist} 期有效样本（元/公斤口径）"
            + (f"，R²={r2:.2f}（{reg_grade}）" if reg_enabled else "；样本不足 3 期或未匹配重量，未做回归")
            + ("；回归已作为最终预测" if reg_used else "；回归仅作参考，主预测采用成本结构/成本传导" if reg_enabled else "")
        ),
    }

    return {
        "predicted_product_price": predicted_product,
        "predicted_product_price_per_kg": (
            round(predicted_product / weight_kg, 6)
            if predicted_product is not None and weight_kg and weight_kg > 1e-12
            else None
        ),
        "baseline_product_price": baseline_product,
        "baseline_product_price_per_kg": baseline_per_kg,
        "product_weight_kg": round(weight_kg, 6) if weight_kg is not None else None,
        "price_unit_label": PRICE_UNIT_LABEL,
        "primary_method": primary_method,
        "credibility_score": credibility,
        "credibility_level": level,
        "credibility_reasons": reasons,
        "credibility_breakdown": {
            "history_sample_score": round(hist_score, 1),
            "fit_score": round(fit_score, 1),
            "price_range_score": round(range_score, 1),
            "agreement_score": round(agreement_score, 1),
            "cost_structure_score": round(structure_score, 1),
            "pass_through_score": round(pass_score, 1),
            "history_sample_count": n_hist if reg_enabled else 0,
            "regression_r2": round(r2, 4) if reg_enabled else None,
            "pass_through_coefficient": round(pass_coef, 6) if pass_coef is not None else None,
            "extrapolation_max_factor": round(float(extrapolation["max_factor"]), 4),
            "credibility_cap": cap,
            "max_score": 100.0,
        },
        "regression_analysis": reg_analysis,
        "alternative_predictions": [
            {"method": m, "price": round(v, 6)} for m, v, _ in candidates if m != primary_method
        ],
    }


def map_sensitivity_item_en_to_zh(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "材料编码": item.get("material_code"),
        "材料型号": item.get("material_name"),
        "用户单价": item.get("user_price"),
        "单价区间": item.get("price_range"),
        "单价波动_pct": item.get("price_swing_pct"),
        "产品价格_kg区间": item.get("product_price_kg_range"),
        "产品价格变动_pct": item.get("product_price_change_pct"),
    }


def map_sensitivity_grid_en_to_zh(grid: dict[str, Any]) -> dict[str, Any]:
    if not grid.get("available"):
        return {
            "可用": False,
            "说明": grid.get("reason") or "需要至少两个改价组件",
        }
    comps = grid.get("components") or []
    return {
        "可用": True,
        "组件": [
            {"编码": c.get("id"), "名称": c.get("name"), "用户单价": c.get("user_price")}
            for c in comps
        ],
        "扰动步长_pct": grid.get("pct_steps"),
        "基准产品价格_kg": grid.get("baseline_per_kg"),
        "矩阵": grid.get("matrix"),
        "行轴": "组件A价格扰动%",
        "列轴": "组件B价格扰动%",
    }


def map_regression_analysis_en_to_zh(reg: dict[str, Any]) -> dict[str, Any]:
    return {
        "是否启用": reg.get("enabled"),
        "是否采用为最终预测": reg.get("used_as_primary"),
        "有效样本数": reg.get("valid_sample_count"),
        "BOM直接样本数": reg.get("bom_direct_sample_count"),
        "R2": reg.get("r2"),
        "R2_百分比": reg.get("r2_pct"),
        "拟合等级": reg.get("fit_grade"),
        "模型": reg.get("model"),
        "截距": reg.get("intercept"),
        "斜率": reg.get("slope"),
        "模拟点_BOM材料每公斤": reg.get("sim_bom_material_per_kg"),
        "回归预测_产品价每公斤": reg.get("reg_predicted_product_price_per_kg"),
        "成本传导系数": reg.get("pass_through_coefficient"),
        "传导系数来源": reg.get("pass_through_source"),
        "传导预测_产品价每公斤": reg.get("pass_predicted_product_price_per_kg"),
        "说明": reg.get("note"),
    }


def map_legacy_predict_en_to_zh(pred: dict[str, Any]) -> dict[str, Any]:
    """legacy 预测英文结构 → 前端中文键。"""
    if "error" in pred:
        return {"error": pred["error"]}
    reg = pred.get("regression_analysis") or {}
    alts = pred.get("alternative_predictions") or []
    breakdown = pred.get("credibility_breakdown") or {}
    return {
        "预测产品价格": pred.get("predicted_product_price"),
        "预测产品价格_每公斤": pred.get("predicted_product_price_per_kg"),
        "基准产品价格": pred.get("baseline_product_price"),
        "基准产品价格_每公斤": pred.get("baseline_product_price_per_kg"),
        "产品重量_kg": pred.get("product_weight_kg"),
        "价格口径": pred.get("price_unit_label"),
        "预测方法": pred.get("primary_method"),
        "预测可信度": pred.get("credibility_score"),
        "可信度等级": pred.get("credibility_level"),
        "可信度说明": "；".join(pred.get("credibility_reasons") or []),
        "可信度明细": {
            "历史样本分": breakdown.get("history_sample_score"),
            "拟合优度分": breakdown.get("fit_score"),
            "价格区间分": breakdown.get("price_range_score"),
            "方法一致性分": breakdown.get("agreement_score"),
            "成本结构分": breakdown.get("cost_structure_score"),
            "成本传导分": breakdown.get("pass_through_score"),
            "历史样本数": breakdown.get("history_sample_count"),
            "回归R2": breakdown.get("regression_r2"),
            "成本传导系数": breakdown.get("pass_through_coefficient"),
            "极端外推倍数": breakdown.get("extrapolation_max_factor"),
            "可信度封顶": breakdown.get("credibility_cap"),
            "满分": breakdown.get("max_score"),
        },
        "回归分析": map_regression_analysis_en_to_zh(reg),
        "备选预测": [
            {"方法": a.get("method"), "价格": a.get("price")} for a in alts
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
        base_mat, base_labor, per_piece = _cost_structure_bases(base_data)
        weight = base_data.get("weight")
        ratio = sim_bom / base_bom_total
        _, per_kg = cost_structure_predict(
            base_mat,
            base_labor,
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

    base_bom_total = float(
        base_data.get("baseline_bom_for_ratio")
        or base_data.get("base_bom_total")
        or 0.0
    )
    base_mat = base_data.get("base_mat")
    sim_bom_center = _sim_bom_from_component_info(component_info)
    ratio_base = max(base_bom_total, MIN_BASE_AMOUNT_YUAN)

    denom_ppk = float(point_per_kg) if point_per_kg and point_per_kg > 1e-12 else None
    if denom_ppk is None:
        denom_ppk = _point_per_kg_from_sim_bom(
            sim_bom_center,
            has_cost_structure=has_cost_structure,
            base_data=base_data,
            coeff=coeff,
            base_for_coeff=base_for_coeff,
            base_bom_total_override=ratio_base,
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
            bom_lo,
            has_cost_structure=has_cost_structure,
            base_data=base_data,
            coeff=coeff,
            base_for_coeff=base_for_coeff,
            base_bom_total_override=ratio_base,
        )
        ppk_hi = _point_per_kg_from_sim_bom(
            bom_hi,
            has_cost_structure=has_cost_structure,
            base_data=base_data,
            coeff=coeff,
            base_for_coeff=base_for_coeff,
            base_bom_total_override=ratio_base,
        )
        if ppk_lo is None or ppk_hi is None:
            continue

        # 成本结构：产品价格(元/kg) 对 BOM 批次合计(元) 的边际 = base_mat/base_bom_total
        if (
            has_cost_structure
            and base_mat is not None
            and not math.isnan(base_mat)
            and ratio_base > 1e-12
        ):
            d_bom_half = quantity * user_price * pct
            mat_for_ratio = base_data.get("base_mat_piece") or base_mat
            half_delta_piece = float(mat_for_ratio) * d_bom_half / ratio_base
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
            "material_code": mat_id,
            "material_name": comp.get("name") or mat_id,
            "user_price": round(user_price, 4),
            "price_range": [round(price_lo, 4), round(price_hi, 4)],
            "price_swing_pct": round(pct * 100, 1),
            "product_price_kg_range": [round(ppk_lo, 4), round(ppk_hi, 4)],
            "product_price_change_pct": round(product_pct, 2),
        })
    return items


def _compute_sensitivity_grid(
    user_modified_prices: dict,
    component_info: list[dict],
    point_per_kg: float,
    *,
    has_cost_structure: bool,
    base_data: dict,
    coeff: float,
    base_for_coeff: float | None,
    pct: float = SENSITIVITY_PCT,
) -> dict[str, Any]:
    """对本次改价 Top2 组件做 3×3 交叉敏感性网格（±pct）。"""
    if len(user_modified_prices) < 2:
        return {"available": False, "reason": "need_at_least_two_modified_components"}

    ranked: list[tuple[str, dict]] = []
    for mid, up in user_modified_prices.items():
        mid = str(mid).strip()
        comp = next((c for c in component_info if str(c["mat_id"]) == mid), None)
        if comp is None:
            continue
        line_cost = float(comp.get("quantity") or 0) * float(up)
        ranked.append((mid, comp))
    ranked.sort(
        key=lambda x: float(x[1].get("quantity") or 0) * float(user_modified_prices.get(x[0], 0)),
        reverse=True,
    )
    if len(ranked) < 2:
        return {"available": False, "reason": "insufficient_modified_components"}

    id_a, comp_a = ranked[0]
    id_b, comp_b = ranked[1]
    base_ppk = float(point_per_kg)
    if base_ppk <= 1e-12:
        return {"available": False, "reason": "invalid_baseline_per_kg"}

    pct_steps = [-pct, 0.0, pct]
    pct_labels = [round(p * 100, 1) for p in pct_steps]
    matrix: list[list[float | None]] = []

    for pa in pct_steps:
        row_vals: list[float | None] = []
        price_a = float(user_modified_prices[id_a]) * (1.0 + pa)
        for pb in pct_steps:
            price_b = float(user_modified_prices[id_b]) * (1.0 + pb)
            overrides = {id_a: price_a, id_b: price_b}
            info = _component_info_with_prices(component_info, overrides)
            bom = _sim_bom_from_component_info(info)
            ppk = _point_per_kg_from_sim_bom(
                bom,
                has_cost_structure=has_cost_structure,
                base_data=base_data,
                coeff=coeff,
                base_for_coeff=base_for_coeff,
                base_bom_total_override=max(
                    float(
                        base_data.get("baseline_bom_for_ratio")
                        or base_data.get("base_bom_total")
                        or 0.0
                    ),
                    MIN_BASE_AMOUNT_YUAN,
                ),
            )
            row_vals.append(round(float(ppk), 4) if ppk is not None else None)
        matrix.append(row_vals)

    return {
        "available": True,
        "components": [
            {
                "id": id_a,
                "name": comp_a.get("name") or id_a,
                "user_price": round(float(user_modified_prices[id_a]), 4),
            },
            {
                "id": id_b,
                "name": comp_b.get("name") or id_b,
                "user_price": round(float(user_modified_prices[id_b]), 4),
            },
        ],
        "pct_steps": pct_labels,
        "baseline_per_kg": round(base_ppk, 4),
        "matrix": matrix,
        "axis_row": "component_a_pct",
        "axis_col": "component_b_pct",
    }


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
    outlier_months: set[str] = set()
    if len(df_err) >= 3:

        def _month_ord(qm: Any) -> int | None:
            s = quote_month_key_scalar(qm)
            if len(s) < 6:
                return None
            try:
                return int(s[:4]) * 12 + int(s[4:6])
            except ValueError:
                return None

        df_err = df_err.copy()
        df_err["_mord"] = df_err["报价月份"].map(_month_ord)
        df_err = df_err.sort_values("_mord")
        mat_diffs: list[float] = []
        diff_row_idx: list[Any] = []
        prev_m: int | None = None
        prev_mat: float | None = None
        for idx, row in df_err.iterrows():
            mord = row["_mord"]
            mat_v = row["材料成本"]
            if mord is None or mat_v is None or pd.isna(mat_v):
                continue
            if prev_m is not None and prev_mat is not None and mord - prev_m == 1:
                mat_diffs.append(float(mat_v) - prev_mat)
                diff_row_idx.append(idx)
            prev_m = int(mord)
            prev_mat = float(mat_v)
        if len(mat_diffs) >= 2:
            diff_std = float(np.std(mat_diffs))
            if diff_std > 1e-12:
                threshold = 2.0 * diff_std
                for i, dmat in enumerate(mat_diffs):
                    if abs(dmat) > threshold:
                        outlier_months.add(str(df_err.loc[diff_row_idx[i], "报价月份"]))

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
        row_uh = unit_hint_from_row(row)
        norm = normalize_price_list_cost_fields(
            float(mat_t), 0.0, None, _safe_float(wt_t), unit_hint=row_uh
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


def _model_error_interval(
    point_per_kg: float,
    model_error: dict[str, Any],
    *,
    bom_perturbation_ratio: float = 0.0,
) -> dict[str, Any]:
    """
    由稳健 MAE 构造点估计区间（元/kg）。
    半宽 = Z×MAE×(1 + γ×|ΔBOM|/基准BOM)，模拟调价幅度越大区间越宽。
    """
    if not model_error.get("可用") or point_per_kg is None:
        return model_error
    mae = float(model_error.get("MAE") or model_error.get("RMSE") or 0.0)
    z = MODEL_ERROR_Z
    gamma = MODEL_ERROR_GAMMA
    perturb = max(0.0, float(bom_perturbation_ratio))
    interval_scale = 1.0 + gamma * perturb
    half = z * mae * interval_scale
    lo = max(0.0, point_per_kg - half)
    hi = point_per_kg + half
    out = dict(model_error)
    out["预测区间_kg"] = [round(lo, 4), round(hi, 4)]
    out["区间半宽"] = round(half, 4)
    out["BOM扰动比例"] = round(perturb, 6)
    out["区间放大系数"] = round(interval_scale, 4)
    out["区间口径"] = (
        f"点估计 ± {z}×MAE×(1+{gamma}×|ΔBOM|/基准BOM)"
        if perturb > 1e-9
        else f"点估计 ± {z}×MAE（稳健）"
    )
    return out

def latest_bom_snapshot_from_rows(
    product_id: str,
    rows: list[dict],
) -> pd.DataFrame:
    """将 BI 页 BOM 表格行转为预测用快照，保证与页面展示的组件列表一致。"""
    pid = norm_material_code_scalar(product_id)
    recs: list[dict[str, Any]] = []
    for r in rows or []:
        comp = str(r.get("组件编码") or "").strip()
        if not comp:
            continue
        try:
            qty = float(r.get("MENGE合计") or 0)
        except (TypeError, ValueError):
            qty = 0.0
        try:
            unit = float(r.get("组件单价") or 0)
        except (TypeError, ValueError):
            unit = 0.0
        name = str(r.get("组件名称") or r.get("MAKTX") or comp)
        recs.append({
            COL_PRODUCT: pid,
            COL_COMPONENT: comp,
            COL_QTY: qty,
            COL_UNIT_PRICE: unit,
            COL_MAKTX: name,
            COL_CREATEDATE: "snapshot",
        })
    return pd.DataFrame(recs) if recs else pd.DataFrame()


def _prepare_base_data(
    product_id,
    bom_df: pd.DataFrame,
    price_df: pd.DataFrame,
    *,
    latest_bom_snapshot: pd.DataFrame | None = None,
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

    pid_str = norm_material_code_scalar(product_id)

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

    if latest_bom_snapshot is not None and not latest_bom_snapshot.empty:
        latest_bom_raw = latest_bom_snapshot.copy()
        latest_bom_raw[COL_QTY] = pd.to_numeric(
            latest_bom_raw[COL_QTY], errors="coerce"
        ).fillna(0.0)
        latest_bom_raw[COL_UNIT_PRICE] = pd.to_numeric(
            latest_bom_raw[COL_UNIT_PRICE], errors="coerce"
        ).fillna(0.0)
    else:
        latest_date = _latest_valid_bom_date(b_hist["_date"])
        latest_bom_raw = b_hist[b_hist["_date"] == latest_date].copy()

    # ── 组件历史统计 ──────────────────────────────────────────────────────
    maktx_col = COL_MAKTX if COL_MAKTX in b_hist.columns else COL_COMPONENT
    if maktx_col == COL_COMPONENT:
        comp_group = b_hist.groupby(COL_COMPONENT, sort=False)[COL_UNIT_PRICE]
        comp_stats = comp_group.agg(
            历史均价="mean",
            历史std="std",
            历史样本数="count",
            历史最低="min",
            历史最高="max",
        ).reset_index().rename(columns={COL_COMPONENT: "材料编码"})
        comp_stats["材料型号"] = comp_stats["材料编码"]
    else:
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
    name_col = (
        latest_bom_raw[COL_MAKTX]
        if COL_MAKTX in latest_bom_raw.columns and maktx_col != COL_COMPONENT
        else latest_bom_raw[COL_COMPONENT]
    )
    latest_bom = pd.DataFrame({
        "材料编码": norm_material_code(latest_bom_raw[COL_COMPONENT]),
        "材料型号": name_col.astype(str),
        "组件数量": latest_bom_raw[COL_QTY].values,
        "材料单价": latest_bom_raw[COL_UNIT_PRICE].values,
    })

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

    # ── 构建 p_hist 供传导系数使用（报价月份、材料成本、重量） ─────────────
    coeff_cols = ["_qm", "材料成本", "重量"]
    p_hist_coeff = p_hist[
        ["_qm"] + [c for c in coeff_cols[1:] if c in p_hist.columns]
    ].copy()
    p_hist_coeff = p_hist_coeff.rename(columns={"_qm": "报价月份"})
    if "材料成本" not in p_hist_coeff.columns:
        p_hist_coeff["材料成本"] = float("nan")

    # 完整价格序列（供历史回测 MAE、传导系数；含重量以便元/件 → 元/kg）
    p_hist_full_cols = ["_qm"] + [
        c for c in ("材料成本", "工费", "总成本", "重量") if c in p_hist.columns
    ]
    p_hist_full = p_hist[p_hist_full_cols].copy().rename(columns={"_qm": "报价月份"})
    for c in ("材料成本", "工费", "总成本", "重量"):
        if c in p_hist_full.columns:
            p_hist_full[c] = pd.to_numeric(p_hist_full[c], errors="coerce")

    price_uh = unit_hint_from_row(latest_price)
    bom_uh = unit_hint_from_bom_dataframe(bom_df, pid_str)
    sum_ml = (
        (0.0 if math.isnan(base_mat) else base_mat)
        + (0.0 if math.isnan(base_labor) else base_labor)
    )
    wt_for_uh = None if math.isnan(weight) else weight
    total_for_uh = None if math.isnan(base_total) else base_total
    if _price_list_values_imply_per_piece(sum_ml, total_for_uh, wt_for_uh):
        cost_uh: UnitHint = "piece"
    elif price_uh != "unknown":
        cost_uh = price_uh
    else:
        cost_uh = bom_uh

    cost_norm = normalize_price_list_cost_fields(
        None if math.isnan(base_mat) else base_mat,
        None if math.isnan(base_labor) else base_labor,
        None if math.isnan(base_total) else base_total,
        None if math.isnan(weight) else weight,
        unit_hint=cost_uh,
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
        "base_mat_per_kg": cost_norm.get("mat_per_kg"),
        "base_labor_per_kg": cost_norm.get("labor_per_kg"),
        "costs_are_per_piece": cost_norm.get("costs_are_per_piece"),
    }, None


def _estimate_conduction_coeff(base_data: dict) -> "tuple[float, str]":
    """
    用历史月度数据估计材料成本(元/kg)对 BOM 材料合计(元/件)的传导系数。

    回归优先将 BOM 批次合计 ÷ 当月重量 换算为 元/kg，与材料成本同量纲；
    无重量时退回 元/件 BOM × 材料成本(元/kg)，系数量纲为 1/kg，与
    predict 中「Δ价/kg = coeff × ΔBOM/件」一致。
    返回 (稳健传导系数, 样本质量评级)。
    """
    fallback = 0.030

    bom_m = base_data["bom_monthly"]
    p_h = base_data["p_hist"]
    weight_col = "重量" if "重量" in p_h.columns else None
    merge_cols = ["报价月份", "材料成本"] + ([weight_col] if weight_col else [])
    merged = bom_m.merge(p_h[merge_cols], on="报价月份", how="inner")
    merged = merged.dropna(subset=["材料成本", "BOM批次合计"])
    merged = merged.sort_values("报价月份")

    if len(merged) < 3:
        return fallback, "global_fallback"

    points: list[tuple[str, float, float]] = []
    month_weights: list[float] = []
    for _, r in merged.iterrows():
        bom_piece = float(r["BOM批次合计"])
        mat_kg = float(r["材料成本"])
        if bom_piece <= 1e-12 or mat_kg <= 1e-12:
            continue
        wt = None
        if weight_col:
            try:
                w = float(r.get(weight_col))
                if w > 1e-12:
                    wt = w
            except (TypeError, ValueError):
                pass
        if wt:
            x = bom_piece / wt
            month_weights.append(wt)
        else:
            x = bom_piece
        points.append((str(r["报价月份"]), x, mat_kg))
    slopes, slope_weights = _robust_slopes_from_points(points)
    if len(slopes) < 2:
        return fallback, "global_fallback"

    coeff = _aggregate_slopes(slopes, slope_weights)
    # 同量纲(元/kg)回归 → 无量纲斜率；换算为 predict 用的 (元/kg)/(元/件)=1/kg
    w_ref: float | None = None
    if month_weights:
        w_ref = float(np.median(month_weights))
        if w_ref > 1e-12:
            coeff = coeff / w_ref
    if w_ref is None:
        bw = base_data.get("weight")
        try:
            w_ref = float(bw) if bw is not None and not math.isnan(float(bw)) else None
        except (TypeError, ValueError):
            w_ref = None
    abs_min, abs_max = _conduction_coeff_clip_bounds(w_ref)
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
    latest_bom_snapshot: pd.DataFrame | None = None,
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
    base_data, err = _prepare_base_data(
        product_id,
        bom_df,
        price_df,
        latest_bom_snapshot=latest_bom_snapshot,
    )
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
    catalog_bom_total = _sum_bom_from_latest_bom(
        lb,
        lambda mid, row: float(row.get("材料单价") or 0.0),
    )
    if ref_map:
        base_bom_total = _sum_bom_from_latest_bom(
            lb,
            lambda mid, row: ref_map.get(
                mid, float(row.get("材料单价") or 0.0)
            ),
        )
    else:
        base_bom_total = catalog_bom_total

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

        if is_fixed:
            ref_mult = new_price / max(ref_unit, LOW_VALUE_UNIT_PRICE_YUAN)
            cap, cred_msg, _r_lbl = _credibility_cap_from_ref_multiple(ref_mult)
            if cap is not None and cred_msg:
                score = min(score, cap)
                emoji = "🚨" if cap <= 20 else "⚠️"
                warnings.append(f"{emoji} 【{mat_name}】{cred_msg}")
                if cap <= 20:
                    warnings.append(
                        f"🚨 【{mat_name}】预测结果不具参考价值"
                    )
                material_warned.add(mat_name)

        if (
            stat is not None
            and is_fixed
            and ref_unit >= LOW_VALUE_UNIT_PRICE_YUAN
            and mat_name not in material_warned
        ):
            hist_mean_chk = float(stat["历史均价"])
            hist_min = float(stat["历史最低"])
            hist_max = float(stat["历史最高"])
            hist_std_chk = float(stat["历史std"])
            hist_mean_eff = max(hist_mean_chk, LOW_VALUE_UNIT_PRICE_YUAN)
            # 改价后用模拟单价估算行占比，避免「单价低但模拟后行成本很大」被误判为非核心
            unit_for_share = new_price if is_fixed else ref_unit
            line_share = (quantity * unit_for_share) / max(
                base_bom_total, MIN_BASE_AMOUNT_YUAN
            )
            is_non_core = line_share < SHOCK_NON_CORE_BOM_SHARE
            abs_delta = abs(new_price - hist_mean_chk)
            rel_ratio = abs_delta / hist_mean_eff
            mild_abs_move = (
                abs_delta < SHOCK_ABS_DELTA_CAP_YUAN and rel_ratio < 2.0
            )

            if hist_std_chk > 0 and not is_non_core:
                std_eff = max(
                    hist_std_chk,
                    SHOCK_MIN_EFFECTIVE_STD_YUAN,
                    hist_mean_chk * 0.05,
                )
                shock_sigma = abs_delta / std_eff
                use_rel_for_low_unit = ref_unit < LOW_VALUE_REL_SIGMA_UNIT_YUAN

                if use_rel_for_low_unit:
                    cap_lu, cred_msg_lu, _r_lu = _credibility_cap_from_ref_multiple(
                        ref_mult
                    )
                    if cap_lu is not None and cred_msg_lu:
                        score = min(score, cap_lu)
                        emoji_lu = "🚨" if cap_lu <= 20 else "⚠️"
                        warnings.append(f"{emoji_lu} 【{mat_name}】{cred_msg_lu}")
                        if cap_lu <= 20:
                            warnings.append(
                                f"🚨 【{mat_name}】预测结果不具参考价值"
                            )
                        material_warned.add(mat_name)
                    elif rel_ratio > 2.0 and (
                        _ref_multiple_magnitude(ref_mult) or 0
                    ) <= CREDIBILITY_REF_NORMAL_MAX:
                        score = min(score, 92)
                        warnings.append(
                            f"⚠️ 【{mat_name}】低单价辅料调价相对历史均价 {rel_ratio:.0%}，"
                            f"请关注波动"
                        )
                        material_warned.add(mat_name)
                elif not mild_abs_move:
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
                            f"⚠️ 【{mat_name}】调价幅度 {shock_sigma:.1f}σ，"
                            f"远超历史范围，可信度极低"
                        )
                        material_warned.add(mat_name)
                    elif shock_sigma > 3.0:
                        score = min(score, 75)
                        warnings.append(
                            f"⚠️ 【{mat_name}】调价幅度 {shock_sigma:.1f}σ，"
                            f"超出历史波动范围"
                        )
                        material_warned.add(mat_name)
                    elif shock_sigma > 2.0:
                        score = min(score, 92)

            if (
                hist_mean_eff > LOW_VALUE_UNIT_PRICE_YUAN
                and mat_name not in material_warned
                and not is_non_core
            ):
                cap_h, cred_msg_h, _r_h = _credibility_cap_from_ref_multiple(
                    new_price / max(hist_mean_chk, LOW_VALUE_UNIT_PRICE_YUAN)
                )
                if cap_h is not None and cred_msg_h:
                    score = min(score, cap_h)
                    emoji_h = "🚨" if cap_h <= 20 else "⚠️"
                    warnings.append(
                        f"{emoji_h} 【{mat_name}】相对历史均价：{cred_msg_h}"
                    )
                    if cap_h <= 20:
                        warnings.append(
                            f"🚨 【{mat_name}】预测结果不具参考价值"
                        )
                    material_warned.add(mat_name)
                elif rel_ratio > 2.0:
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
        base_eff = max(base_bom_total, MIN_BASE_AMOUNT_YUAN)
        ratio = sim_bom_point / base_eff
        mat_base, labor_base, per_piece = _cost_structure_bases(base_data)
        point_total, point_per_kg = cost_structure_predict(
            mat_base,
            labor_base,
            ratio,
            weight if weight and not math.isnan(weight) else None,
            costs_are_per_piece=per_piece,
        )
        bl_kg = base_data.get("base_total_per_kg")
        bom_perturb_early = abs(sim_bom_point - base_bom_total) / max(
            base_bom_total, MIN_BASE_AMOUNT_YUAN
        )
        # 仅在「BOM 倍率不大但单价/kg 飙高」时提示元/件/元/kg 口径错乱，避免辅料调价误伤
        if (
            point_per_kg is not None
            and bl_kg is not None
            and not (isinstance(bl_kg, float) and math.isnan(bl_kg))
            and float(bl_kg) > 1e-12
            and float(point_per_kg) > float(bl_kg) * 4.0
            and bom_perturb_early < 0.2
        ):
            warnings.append(
                f"预测价({point_per_kg:.2f}元/kg)远高于清单基准({float(bl_kg):.2f}元/kg)，"
                "请核对价格清单口径是否为元/件"
            )
            score = min(score, 40)
        if point_total is None:
            point_total = 0.0
        if point_per_kg is None:
            point_per_kg = point_total
        method = "cost_structure_ratio"
    else:
        coeff, coeff_quality = _estimate_conduction_coeff(base_data)
        # ΔBOM 为元/件；coeff 量纲 1/kg → Δ价/kg = coeff × ΔBOM/件
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

    base_data["baseline_bom_for_ratio"] = base_bom_total

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
    sensitivity_grid = _compute_sensitivity_grid(
        modified_only,
        component_info,
        float(point_per_kg),
        has_cost_structure=has_cost_structure,
        base_data=base_data,
        coeff=coeff,
        base_for_coeff=_base_for_coeff,
    )

    # ── 步骤D：模型历史误差（回测 MAE；区间半宽随本次 ΔBOM 放大）────────
    model_error_raw = _compute_historical_model_error(
        base_data,
        has_cost_structure=has_cost_structure,
        coeff=coeff,
        base_for_coeff=_base_for_coeff,
    )
    bom_perturbation_ratio = abs(sim_bom_point - base_bom_total) / max(
        base_bom_total, MIN_BASE_AMOUNT_YUAN
    )
    model_error = _model_error_interval(
        float(point_per_kg),
        model_error_raw,
        bom_perturbation_ratio=bom_perturbation_ratio,
    )

    # ── 步骤E：可信度评分修正项 ───────────────────────────────────────────
    n_price_months = base_data["n_price_months"]
    if n_price_months >= 12 and score >= 95 and not material_warned:
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
        "sensitivity_grid": sensitivity_grid,
        "model_error": model_error,
        "detail": {
            "base_mat_cost": base_mat,
            "base_labor": base_labor,
            "base_weight": weight,
            "base_bom_total": round(base_bom_total, 4),
            "sim_bom_point": round(sim_bom_point, 4),
            "bom_ratio": (
                round(sim_bom_point / max(base_bom_total, MIN_BASE_AMOUNT_YUAN), 6)
                if base_bom_total > 0
                else None
            ),
            "n_price_months": n_price_months,
            "components": component_info,
        },
    }
