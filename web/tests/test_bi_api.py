# -*- coding: utf-8 -*-
"""BI API smoke tests with Flask test client."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import polars as pl
import pytest

_WEB = Path(__file__).resolve().parent.parent
if str(_WEB) not in sys.path:
    sys.path.insert(0, str(_WEB))

from process_excel import prepare, aggregate_lines, enrich, sheet_component_global_rank, sheet_bom_detail
from services import bi_cache


@pytest.fixture
def client(tmp_path: Path):
    results = tmp_path / "results"
    uploads = tmp_path / "uploads"
    results.mkdir()
    uploads.mkdir()

    import app as app_module

    app_module.RESULT_DIR = results
    app_module.UPLOAD_DIR = uploads
    bi_cache.init(results)
    app_module.app.config["TESTING"] = True
    return app_module.app.test_client()


def _seed_session(results: Path, token: str) -> None:
    raw = pl.DataFrame({
        "产品编码": ["P100", "P100"],
        "组件编码": ["C1", "C2"],
        "MENGE": [1.0, 2.0],
        "组件单价": [10.0, 5.0],
    })
    lines = enrich(aggregate_lines(prepare(raw)))
    bi_cache.fill_and_persist(
        token,
        sheet_component_global_rank(lines),
        sheet_bom_detail(lines),
        price_history_pl=None,
        product_cost_history={},
        product_price_timeline=None,
    )


def test_bi_ready_and_products(client, tmp_path: Path) -> None:
    token = "c3d4e5f6-a7b8-9012-cdef-123456789012"
    _seed_session(tmp_path / "results", token)

    r = client.get(f"/api/bi_ready/{token}")
    assert r.status_code == 200
    assert r.get_json()["ready"] is True

    r2 = client.get(f"/api/bi/products/{token}")
    assert r2.status_code == 200
    data = r2.get_json()
    assert data["search_required"] is False
    ids = [p["id"] for p in data["products"]]
    assert "P100" in ids


def test_bi_detail_and_simulate(client, tmp_path: Path) -> None:
    token = "d4e5f6a7-b8c9-0123-def0-234567890123"
    _seed_session(tmp_path / "results", token)

    r = client.get(f"/api/bi/detail/{token}?component=C1")
    assert r.status_code == 200
    assert r.get_json()["count"] >= 1

    r2 = client.post(
        f"/api/bi/product_cost_simulate/{token}",
        data=json.dumps({"product": "P100", "prices": {}}),
        content_type="application/json",
    )
    assert r2.status_code == 200
    body = r2.get_json()
    assert body["product"] == "P100"
    assert "模拟原材料总成本" in body
