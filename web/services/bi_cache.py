# -*- coding: utf-8 -*-
"""BI 会话内存缓存、侧车持久化与统一 entry 构建。"""
from __future__ import annotations

import json
import os
import re
import threading
import time
from collections import OrderedDict
from pathlib import Path
from typing import Any

import pandas as pd
import polars as pl

from process_excel import (
    COL_PRODUCT,
    COL_COMPONENT,
    COL_CREATEDATE,
    COL_CATEGORY,
    COL_MAKTX,
    COL_QTY,
    COL_UNIT_PRICE,
    COL_UNIT,
)

_ROOT = Path(__file__).resolve().parent.parent.parent
import sys

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from bom_date_key import calendar_date_key_yyyymmdd as _calendar_date_key_yyyymmdd

from cost_sim_predict import expand_product_cost_history_dict

_SIDECAR_VERSION = 1
_UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$"
)

_RESULT_DIR: Path | None = None
_BI_CACHE: OrderedDict[str, dict] = OrderedDict()
_BI_CACHE_LOCK = threading.Lock()

_BOM_PREDICT_COLS = frozenset({
    COL_PRODUCT, COL_COMPONENT, COL_QTY, COL_UNIT_PRICE, COL_CREATEDATE,
    COL_MAKTX, COL_UNIT, COL_CATEGORY,
    "报价月份", "产品价格", "材料成本", "工费", "重量", "标杆工厂",
    "报价号", "ZBJNO", "报价流水号", "ZSNO",
    "MEINS", "基本单位", "计量单位",
})
_PRICE_PREDICT_COLS = frozenset({
    COL_PRODUCT, COL_CREATEDATE,
    "ZMATNR", "MATNR", "所属产品",
    "总成本", "产品价格", "DMBTR", "标价", "出厂价", "销售价",
    "报价月份", "重量", "材料成本", "工费", "标杆工厂",
})


def init(result_dir: Path) -> None:
    global _RESULT_DIR
    _RESULT_DIR = result_dir


def result_dir() -> Path:
    if _RESULT_DIR is None:
        raise RuntimeError("bi_cache.init(result_dir) 未调用")
    return _RESULT_DIR


def cache_max() -> int:
    try:
        return max(1, int(os.environ.get("BOM_BI_CACHE_MAX", "8")))
    except ValueError:
        return 8


def write_excel_enabled() -> bool:
    return os.environ.get("BOM_WRITE_EXCEL", "1").strip().lower() not in (
        "0", "false", "no", "off",
    )


def uuid_valid(session_id: str) -> bool:
    return bool(_UUID_RE.match(session_id))


def _slim_df(df: pd.DataFrame | None, keep_cols: frozenset) -> pd.DataFrame | None:
    if df is None or df.empty:
        return df
    cols = [c for c in df.columns if c in keep_cols]
    return df[cols] if cols else df


def _detail_frames_from_pl(detail_pl: pl.DataFrame) -> tuple[dict[str, list], dict[str, list], list[str], dict[str, list[str]]]:
    detail_pd = detail_pl.to_pandas()
    detail_pd["组件编码"] = detail_pd["组件编码"].astype(str).str.strip()
    detail_pd[COL_PRODUCT] = detail_pd[COL_PRODUCT].astype(str).str.strip()
    if "MENGE合计" in detail_pd.columns:
        detail_pd = detail_pd.sort_values("MENGE合计", ascending=False)

    detail_index: dict[str, list] = {}
    for comp, grp in detail_pd.groupby("组件编码", sort=False):
        detail_index[str(comp)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    product_bom: dict[str, list] = {}
    for pid, grp in detail_pd.groupby(COL_PRODUCT, sort=False):
        product_bom[str(pid)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    all_product_ids = sorted(product_bom.keys(), key=lambda s: (len(s), s))
    product_prefix_index: dict[str, list[str]] = {}
    for pid in all_product_ids:
        prefix = pid[:2] if len(pid) >= 2 else pid
        product_prefix_index.setdefault(prefix, []).append(pid)

    return detail_index, product_bom, all_product_ids, product_prefix_index


def _detail_frames_from_pandas(detail_df: pd.DataFrame) -> tuple[dict[str, list], dict[str, list], list[str], dict[str, list[str]]]:
    detail_df = detail_df.copy()
    detail_df["组件编码"] = detail_df["组件编码"].astype(str).str.strip()
    detail_df[COL_PRODUCT] = detail_df[COL_PRODUCT].astype(str).str.strip()
    if "MENGE合计" in detail_df.columns:
        detail_df = detail_df.sort_values("MENGE合计", ascending=False)

    detail_index: dict[str, list] = {}
    for comp, grp in detail_df.groupby("组件编码", sort=False):
        detail_index[str(comp)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    product_bom: dict[str, list] = {}
    for pid, grp in detail_df.groupby(COL_PRODUCT, sort=False):
        product_bom[str(pid)] = grp.where(pd.notna(grp), None).to_dict(orient="records")

    all_product_ids = sorted(product_bom.keys(), key=lambda s: (len(s), s))
    product_prefix_index: dict[str, list[str]] = {}
    for pid in all_product_ids:
        prefix = pid[:2] if len(pid) >= 2 else pid
        product_prefix_index.setdefault(prefix, []).append(pid)

    return detail_index, product_bom, all_product_ids, product_prefix_index


def build_bi_cache_entry(
    *,
    summary_pl: pl.DataFrame | None = None,
    summary_records: list | None = None,
    detail_pl: pl.DataFrame | None = None,
    detail_df: pd.DataFrame | None = None,
    price_history_pl: pl.DataFrame | None = None,
    price_history_records: list | None = None,
    product_cost_history: dict | None = None,
    product_price_timeline: dict | None = None,
    bom_pd: pd.DataFrame | None = None,
    price_pd: pd.DataFrame | None = None,
    product_categories: dict | None = None,
) -> dict:
    """从 Polars/pandas 表或已有 records 构建统一 BI 缓存 entry。"""
    if summary_records is None:
        if summary_pl is None:
            summary_records = []
        else:
            summary_pd = summary_pl.to_pandas()
            summary_records = summary_pd.where(pd.notna(summary_pd), None).to_dict(orient="records")

    if detail_pl is not None:
        detail_index, product_bom, all_product_ids, product_prefix_index = _detail_frames_from_pl(detail_pl)
    elif detail_df is not None:
        detail_index, product_bom, all_product_ids, product_prefix_index = _detail_frames_from_pandas(detail_df)
    else:
        detail_index = {}
        product_bom = {}
        all_product_ids = []
        product_prefix_index = {}

    if price_history_records is None:
        price_history_records = []
        if price_history_pl is not None:
            try:
                ph_pd = price_history_pl.to_pandas()
                if COL_CREATEDATE in ph_pd.columns:
                    ph_pd[COL_CREATEDATE] = ph_pd[COL_CREATEDATE].map(_calendar_date_key_yyyymmdd)
                price_history_records = ph_pd.where(pd.notna(ph_pd), None).to_dict(orient="records")
            except Exception:
                pass

    bom_pd = _slim_df(bom_pd, _BOM_PREDICT_COLS)
    price_pd = _slim_df(price_pd, _PRICE_PREDICT_COLS)

    pch = expand_product_cost_history_dict(
        product_cost_history or {},
        product_bom,
        product_price_timeline,
    )

    return {
        "summary": summary_records,
        "detail": detail_index,
        "product_bom": product_bom,
        "all_product_ids": all_product_ids,
        "product_prefix_index": product_prefix_index,
        "price_history": price_history_records,
        "product_cost_history": pch,
        "bom_pd": bom_pd,
        "price_pd": price_pd,
        "product_categories": product_categories or {},
        "ts": time.time(),
    }


def _bi_cache_put(token: str, entry: dict) -> None:
    with _BI_CACHE_LOCK:
        if token in _BI_CACHE:
            del _BI_CACHE[token]
        elif len(_BI_CACHE) >= cache_max():
            _BI_CACHE.popitem(last=False)
        _BI_CACHE[token] = entry


def sidecar_json_path(session_id: str) -> Path:
    return result_dir() / f"{session_id}_bi.json"


def sidecar_bom_path(session_id: str) -> Path:
    return result_dir() / f"{session_id}_bom.parquet"


def sidecar_price_path(session_id: str) -> Path:
    return result_dir() / f"{session_id}_price.parquet"


def session_ready(session_id: str) -> bool:
    """BI 是否可用：内存缓存、侧车或（可选）结果 xlsx。"""
    if not uuid_valid(session_id):
        return False
    with _BI_CACHE_LOCK:
        c = _BI_CACHE.get(session_id)
        if c and "product_bom" in c:
            return True
    if sidecar_json_path(session_id).exists():
        return True
    return (result_dir() / f"{session_id}.xlsx").exists()


def _entry_to_sidecar_payload(entry: dict) -> dict:
    return {
        "version": _SIDECAR_VERSION,
        "ts": entry.get("ts", time.time()),
        "summary": entry["summary"],
        "detail": entry["detail"],
        "product_bom": entry["product_bom"],
        "all_product_ids": entry["all_product_ids"],
        "product_prefix_index": entry["product_prefix_index"],
        "price_history": entry["price_history"],
        "product_cost_history": entry["product_cost_history"],
        "product_categories": entry.get("product_categories") or {},
    }


def _entry_from_sidecar_payload(payload: dict, bom_pd: pd.DataFrame | None, price_pd: pd.DataFrame | None) -> dict:
    entry = {
        "summary": payload.get("summary") or [],
        "detail": payload.get("detail") or {},
        "product_bom": payload.get("product_bom") or {},
        "all_product_ids": payload.get("all_product_ids") or [],
        "product_prefix_index": payload.get("product_prefix_index") or {},
        "price_history": payload.get("price_history") or [],
        "product_cost_history": payload.get("product_cost_history") or {},
        "product_categories": payload.get("product_categories") or {},
        "bom_pd": bom_pd,
        "price_pd": price_pd,
        "ts": payload.get("ts", time.time()),
    }
    if not entry["all_product_ids"] and entry["product_bom"]:
        entry["all_product_ids"] = sorted(
            entry["product_bom"].keys(), key=lambda s: (len(s), s)
        )
    if not entry["product_prefix_index"] and entry["all_product_ids"]:
        idx: dict[str, list[str]] = {}
        for pid in entry["all_product_ids"]:
            prefix = pid[:2] if len(pid) >= 2 else pid
            idx.setdefault(prefix, []).append(pid)
        entry["product_prefix_index"] = idx
    return entry


def save_sidecar(session_id: str, entry: dict) -> None:
    path = sidecar_json_path(session_id)
    payload = _entry_to_sidecar_payload(entry)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    bom = entry.get("bom_pd")
    price = entry.get("price_pd")
    bom_p = sidecar_bom_path(session_id)
    price_p = sidecar_price_path(session_id)
    if bom is not None and not bom.empty:
        bom.to_parquet(bom_p, index=False)
    elif bom_p.exists():
        bom_p.unlink(missing_ok=True)
    if price is not None and not price.empty:
        price.to_parquet(price_p, index=False)
    elif price_p.exists():
        price_p.unlink(missing_ok=True)


def load_sidecar(session_id: str) -> dict | None:
    path = sidecar_json_path(session_id)
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    bom_pd = None
    price_pd = None
    bom_p = sidecar_bom_path(session_id)
    price_p = sidecar_price_path(session_id)
    if bom_p.exists():
        try:
            bom_pd = pd.read_parquet(bom_p)
        except Exception:
            pass
    if price_p.exists():
        try:
            price_pd = pd.read_parquet(price_p)
        except Exception:
            pass
    return _entry_from_sidecar_payload(payload, bom_pd, price_pd)


def load_from_excel(session_id: str, file_path: Path) -> dict:
    summary_df = pd.read_excel(file_path, sheet_name="组件全局成本排名")
    detail_df = pd.read_excel(file_path, sheet_name="BOM明细_占比与成本")
    price_history_records: list = []
    try:
        ph_df = pd.read_excel(file_path, sheet_name="价格历史明细")
        if COL_CREATEDATE in ph_df.columns:
            ph_df[COL_CREATEDATE] = ph_df[COL_CREATEDATE].map(_calendar_date_key_yyyymmdd)
        price_history_records = ph_df.where(pd.notna(ph_df), None).to_dict(orient="records")
    except Exception:
        pass
    entry = build_bi_cache_entry(
        detail_df=detail_df,
        summary_records=summary_df.where(pd.notna(summary_df), None).to_dict(orient="records"),
        price_history_records=price_history_records,
        product_cost_history={},
    )
    _bi_cache_put(session_id, entry)
    save_sidecar(session_id, entry)
    return entry


def fill_and_persist(
    token: str,
    summary_pl: pl.DataFrame,
    detail_pl: pl.DataFrame,
    price_history_pl: pl.DataFrame | None = None,
    product_cost_history: dict | None = None,
    product_price_timeline: dict | None = None,
    bom_pd: pd.DataFrame | None = None,
    price_pd: pd.DataFrame | None = None,
    product_categories: dict | None = None,
) -> None:
    entry = build_bi_cache_entry(
        summary_pl=summary_pl,
        detail_pl=detail_pl,
        price_history_pl=price_history_pl,
        product_cost_history=product_cost_history,
        product_price_timeline=product_price_timeline,
        bom_pd=bom_pd,
        price_pd=price_pd,
        product_categories=product_categories,
    )
    _bi_cache_put(token, entry)
    save_sidecar(token, entry)


def get(session_id: str, file_path: Path | None = None) -> dict:
    """内存 → 侧车 → Excel 回退。"""
    with _BI_CACHE_LOCK:
        if session_id in _BI_CACHE:
            c = _BI_CACHE[session_id]
            if "product_bom" not in c:
                del _BI_CACHE[session_id]
            else:
                _BI_CACHE[session_id]["ts"] = time.time()
                return _BI_CACHE[session_id]

    loaded = load_sidecar(session_id)
    if loaded is not None:
        _bi_cache_put(session_id, loaded)
        return loaded

    fp = file_path or (result_dir() / f"{session_id}.xlsx")
    if not fp.exists():
        raise FileNotFoundError("分析结果不存在或已过期，请重新上传并分析")
    return load_from_excel(session_id, fp)


def peek(session_id: str) -> dict | None:
    with _BI_CACHE_LOCK:
        c = _BI_CACHE.get(session_id)
        if c and "product_bom" in c:
            return c
    return load_sidecar(session_id)


def session_has_price_history(cache: dict | None) -> bool:
    if not cache:
        return False
    price_pd = cache.get("price_pd")
    if price_pd is not None:
        try:
            if not price_pd.empty:
                return True
        except AttributeError:
            pass
    return bool(cache.get("product_price_timeline"))


def invalidate(session_id: str) -> None:
    with _BI_CACHE_LOCK:
        _BI_CACHE.pop(session_id, None)


def cache_or_404(session_id: str) -> tuple[dict | None, Path | None, tuple | None]:
    """供 Flask 路由使用：(cache, file_path, error_response)。"""
    from flask import jsonify
    import traceback

    if not uuid_valid(session_id):
        return None, None, (jsonify({"error": "无效的令牌"}), 400)
    file_path = result_dir() / f"{session_id}.xlsx"
    if not session_ready(session_id):
        return None, None, (
            jsonify({"error": "会话不存在或已过期，请重新上传并分析"}),
            404,
        )
    try:
        return get(session_id, file_path), file_path, None
    except FileNotFoundError as exc:
        return None, None, (jsonify({"error": str(exc)}), 404)
    except Exception as exc:
        tb = traceback.format_exc().strip().split("\n")
        return None, None, (
            jsonify({"error": f"读取失败：{exc}", "detail": tb[-1]}),
            500,
        )

