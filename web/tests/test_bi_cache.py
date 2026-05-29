# -*- coding: utf-8 -*-
"""BI cache entry build and sidecar round-trip."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import polars as pl
import pytest

_WEB = Path(__file__).resolve().parent.parent
if str(_WEB) not in sys.path:
    sys.path.insert(0, str(_WEB))

from process_excel import COL_PRODUCT, COL_COMPONENT, prepare, aggregate_lines, enrich
from services import bi_cache


@pytest.fixture
def result_dir(tmp_path: Path) -> Path:
    d = tmp_path / "results"
    d.mkdir()
    bi_cache.init(d)
    return d


def test_build_bi_cache_entry_has_prefix_index() -> None:
    raw = pl.DataFrame({
        "产品编码": ["P1", "P1", "P2"],
        "组件编码": ["C1", "C2", "C1"],
        "MENGE": [1.0, 2.0, 1.0],
        "组件单价": [10.0, 5.0, 8.0],
    })
    lines = enrich(aggregate_lines(prepare(raw)))
    detail = lines.select([
        COL_PRODUCT, COL_COMPONENT, "MENGE合计", "用量占比%", "组件单价", "行原材料成本",
    ])
    summary = pl.DataFrame({
        "组件编码": ["C1"],
        "全局总成本贡献": [18.0],
    })
    entry = bi_cache.build_bi_cache_entry(summary_pl=summary, detail_pl=detail)
    assert "all_product_ids" in entry
    assert "product_prefix_index" in entry
    assert set(entry["all_product_ids"]) == {"P1", "P2"}
    assert entry["product_prefix_index"].get("P1") == ["P1"]
    assert "product_bom" in entry and "P1" in entry["product_bom"]


def test_sidecar_roundtrip(result_dir: Path) -> None:
    token = "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
    entry = {
        "summary": [{"组件编码": "C1"}],
        "detail": {"C1": [{"产品编码": "P1", "组件编码": "C1"}]},
        "product_bom": {"P1": [{"产品编码": "P1", "组件编码": "C1", "MENGE合计": 1}]},
        "all_product_ids": ["P1"],
        "product_prefix_index": {"P1": ["P1"]},
        "price_history": [],
        "product_cost_history": {},
        "product_categories": {},
        "bom_pd": None,
        "price_pd": None,
        "ts": 0.0,
    }
    bi_cache.save_sidecar(token, entry)
    bi_cache.invalidate(token)
    loaded = bi_cache.load_sidecar(token)
    assert loaded is not None
    assert loaded["all_product_ids"] == ["P1"]
    assert loaded["product_prefix_index"]["P1"] == ["P1"]


def test_load_from_excel_fallback(result_dir: Path, tmp_path: Path) -> None:
    import pandas as pd

    token = "b2c3d4e5-f6a7-8901-bcde-f12345678901"
    xlsx = result_dir / f"{token}.xlsx"
    detail = pd.DataFrame({
        "产品编码": ["P1"],
        "组件编码": ["C1"],
        "MENGE合计": [1.0],
        "用量占比%": [100.0],
        "组件单价": [10.0],
        "行原材料成本": [10.0],
    })
    summary = pd.DataFrame({
        "组件编码": ["C1"],
        "全局总成本贡献": [10.0],
        "全局成本占比%": [100.0],
        "涉及产品数": [1],
    })
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="组件全局成本排名", index=False)
        detail.to_excel(w, sheet_name="BOM明细_占比与成本", index=False)
    entry = bi_cache.load_from_excel(token, xlsx)
    assert entry["all_product_ids"] == ["P1"]
    assert "product_prefix_index" in entry
    sidecar = bi_cache.sidecar_json_path(token)
    assert sidecar.exists()
