# -*- coding: utf-8 -*-
"""Unit tests for cost_sim_predict price-row infer and model error interval."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_WEB = Path(__file__).resolve().parent.parent
_ROOT = _WEB.parent
if str(_WEB) not in sys.path:
    sys.path.insert(0, str(_WEB))
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import polars as pl

import pandas as pd

from cost_sim_predict import (  # noqa: E402
    COL_COMPONENT,
    COL_CREATEDATE,
    COL_MAKTX,
    COL_PRODUCT,
    COL_QTY,
    COL_UNIT_PRICE,
    _credibility_cap_from_ref_multiple,
    predict_product_price,
    MODEL_ERROR_GAMMA,
    MODEL_ERROR_Z,
    SLOPE_POOL_MAX,
    actual_price_per_kg_from_price_row,
    baseline_prices_from_price_snap,
    build_product_cost_history,
    latest_bom_snapshot_from_rows,
    _conduction_coeff_clip_bounds,
    _infer_total_per_kg_when_mat_labor_missing,
    _latest_valid_bom_date,
    _linear_fit_predict,
    _model_error_interval,
    _pool_slopes_from_history,
    _regression_fit_grade,
    _robust_slopes_from_points,
    normalize_price_list_cost_fields,
)
from process_excel import normalize_columns  # noqa: E402


class TestInferTotalPerKg:
    def test_piece_total_with_weight(self):
        assert _infer_total_per_kg_when_mat_labor_missing(45.0, 2.5) == pytest.approx(18.0)

    def test_per_kg_total_with_weight(self):
        assert _infer_total_per_kg_when_mat_labor_missing(8.5, 2.5) == pytest.approx(8.5)

    def test_piece_total_without_weight_returns_none(self):
        assert _infer_total_per_kg_when_mat_labor_missing(45.0, None) is None

    def test_small_total_without_weight_ok(self):
        assert _infer_total_per_kg_when_mat_labor_missing(8.5, None) == pytest.approx(8.5)

    def test_actual_price_row_total_only(self):
        assert actual_price_per_kg_from_price_row(None, None, 45.0, 2.5) == pytest.approx(18.0)
        assert actual_price_per_kg_from_price_row(None, None, 45.0, None) is None

    def test_kg_unit_hint_overrides_high_total(self):
        assert _infer_total_per_kg_when_mat_labor_missing(
            45.0, 2.5, unit_hint="KG"
        ) == pytest.approx(45.0)


class TestNormalizePriceListCostFields:
    def test_pc_unit_forces_per_piece(self):
        norm = normalize_price_list_cost_fields(
            30.0, 15.0, 45.0, 2.5, unit_hint="PC"
        )
        assert norm["costs_are_per_piece"] is True
        assert norm["total_piece"] == pytest.approx(45.0)
        assert norm["total_per_kg"] == pytest.approx(18.0)

    def test_kg_unit_avoids_piece_misread(self):
        norm = normalize_price_list_cost_fields(
            30.0, 15.0, 45.0, 2.5, unit_hint="KG"
        )
        assert norm["costs_are_per_piece"] is False
        assert norm["total_per_kg"] == pytest.approx(45.0)

    def test_unknown_unit_low_sum_is_per_kg(self):
        norm = normalize_price_list_cost_fields(5.0, 3.0, 8.0, 2.5)
        assert norm["costs_are_per_piece"] is False
        assert norm["total_per_kg"] == pytest.approx(8.0)


class TestRobustSlopesTimeDecay:
    def test_spanning_pair_uses_average_months_not_latest_only(self):
        """跨多年点对应用双端点平均 months_ago，权重应明显低于仅看近端。"""
        points = [("202001", 100.0, 10.0), ("202501", 120.0, 12.0)]
        _, weights = _robust_slopes_from_points(points)
        assert weights is not None and len(weights) == 1
        assert weights[0] < 0.15

    def test_recent_pair_weight_higher_than_spanning(self):
        points = [
            ("202001", 100.0, 10.0),
            ("202411", 110.0, 11.0),
            ("202501", 120.0, 12.0),
        ]
        _, weights = _robust_slopes_from_points(points)
        assert weights is not None and len(weights) >= 2
        assert min(weights) < max(weights) * 0.5


class TestLatestValidBomDate:
    def test_ignores_zero_placeholder(self):
        import pandas as pd

        s = pd.Series(["20240101", "00000000", "20230615"])
        assert _latest_valid_bom_date(s) == "20240101"

    def test_all_invalid_falls_back_to_max(self):
        import pandas as pd

        s = pd.Series(["00000000", "00000000"])
        assert _latest_valid_bom_date(s) == "00000000"


class TestPieceCostListNotMisreadAsKg:
    def test_mat_labor_total_26_piece_with_4kg_weight(self):
        norm = normalize_price_list_cost_fields(23.794, 2.857, 26.651, 4.0817)
        assert norm["costs_are_per_piece"] is True
        assert norm["total_per_kg"] == pytest.approx(26.651 / 4.0817, rel=1e-3)

    def test_bom_kg_hint_does_not_override_piece_price_list(self):
        """BOM 原料行多为 KG，不能覆盖成品清单的元/件口径。"""
        norm = normalize_price_list_cost_fields(
            23.794, 2.857, 26.651, 4.0817, unit_hint="kg"
        )
        assert norm["costs_are_per_piece"] is True
        assert norm["total_piece"] == pytest.approx(26.651)


class TestLatestBomSnapshot:
    def test_snapshot_lists_all_table_components(self):
        rows = [
            {"组件编码": "A", "MENGE合计": 10.0, "组件单价": 1.0},
            {"组件编码": "B", "MENGE合计": 7.5, "组件单价": 1.3},
        ]
        snap = latest_bom_snapshot_from_rows("P1", rows)
        assert len(snap) == 2

    def test_high_price_on_snapshot_row_moves_prediction(self):
        bom_df = pd.DataFrame(
            {
                COL_PRODUCT: ["P1", "P1"],
                COL_COMPONENT: ["PVC", "PVC"],
                COL_QTY: [2125.0, 2125.0],
                COL_UNIT_PRICE: [5.0, 5.0],
                COL_CREATEDATE: ["20250101", "20250201"],
            }
        )
        price_df = pd.DataFrame(
            {
                "产品编码": ["P1"],
                "报价月份": ["202501"],
                "重量": [4.0],
                "材料成本": [20.0],
                "工费": [3.0],
                "总成本": [23.0],
            }
        )
        snap = latest_bom_snapshot_from_rows(
            "P1",
            [
                {"组件编码": "PVC", "MENGE合计": 2125.0, "组件单价": 5.0},
                {"组件编码": "CA", "MENGE合计": 7.5, "组件单价": 1.3},
            ],
        )
        ref = {"PVC": 5.0, "CA": 1.3}
        base = predict_product_price(
            "P1", {}, bom_df, price_df, reference_prices=ref, latest_bom_snapshot=snap
        )
        high = predict_product_price(
            "P1",
            {"CA": 122.3},
            bom_df,
            price_df,
            reference_prices=ref,
            latest_bom_snapshot=snap,
        )
        assert "error" not in base and "error" not in high
        assert float(high["point_per_kg"]) > float(base["point_per_kg"])


class TestCredibilityRefBands:
    @pytest.mark.parametrize(
        "ref_mult,expected_cap",
        [
            (1.05, None),
            (1.1, None),
            (1.2, 95),
            (1.3, 90),
            (1.45, 82),
            (1.8, 65),
            (2.5, 48),
            (4.0, 28),
            (6.0, 15),
            (0.5, 55),
        ],
    )
    def test_piecewise_caps(self, ref_mult, expected_cap):
        cap, msg, _ = _credibility_cap_from_ref_multiple(ref_mult)
        if expected_cap is None:
            assert cap is None
            assert msg is None
        else:
            assert cap == expected_cap
            assert msg and "倍）" in msg

    def test_caps_monotone_with_r(self):
        mults = [1.0, 1.15, 1.35, 1.7, 2.2, 3.0, 4.5, 8.0]
        caps = []
        for m in mults:
            c, _, _ = _credibility_cap_from_ref_multiple(m)
            caps.append(100 if c is None else c)
        assert caps == sorted(caps, reverse=True)


class TestConfidenceOnExtremePriceEdit:
    def _calcium_pred(self, unit_price: float) -> dict:
        """主物料 PVC + 辅料碳酸钙，与页面 BOM 结构一致。"""
        bom_df = pd.DataFrame(
            {
                COL_PRODUCT: ["P1", "P1"],
                COL_COMPONENT: ["PVC", "CA"],
                COL_MAKTX: ["聚氯乙烯五型", "活性碳酸钙"],
                COL_QTY: [2125.0, 7.5],
                COL_UNIT_PRICE: [5.0824, 1.3],
                COL_CREATEDATE: ["20250115", "20250115"],
            }
        )
        price_df = pd.DataFrame(
            {
                "产品编码": ["P1"] * 12,
                "报价月份": [f"2024{m:02d}" for m in range(1, 13)],
                "重量": [4.0817] * 12,
                "材料成本": [23.794] * 12,
                "工费": [2.857] * 12,
                "总成本": [26.651] * 12,
            }
        )
        snap = latest_bom_snapshot_from_rows(
            "P1",
            [
                {
                    "组件编码": "PVC",
                    "MENGE合计": 2125.0,
                    "组件单价": 5.0824,
                    "组件名称": "聚氯乙烯五型",
                },
                {
                    "组件编码": "CA",
                    "MENGE合计": 7.5,
                    "组件单价": 1.3,
                    "组件名称": "活性碳酸钙",
                },
            ],
        )
        return predict_product_price(
            "P1",
            {"CA": unit_price},
            bom_df,
            price_df,
            reference_prices={"PVC": 5.0824, "CA": 1.3},
            latest_bom_snapshot=snap,
        )

    def test_business_ref_multiple_bands(self):
        """相对加载价：1.1 倍内满分；5 倍以上保底 15 分。"""
        s13 = self._calcium_pred(1.3)["confidence_score"]
        s143 = self._calcium_pred(1.43)["confidence_score"]
        s169 = self._calcium_pred(1.69)["confidence_score"]
        s26 = self._calcium_pred(2.6)["confidence_score"]
        s1223 = self._calcium_pred(122.3)["confidence_score"]
        assert s13 >= 95
        assert 85 <= s143 <= 100
        assert 70 <= s169 <= 95
        assert 30 <= s26 <= 60
        assert s1223 <= 20
        assert s143 > s169 > s26 > s1223

    def test_hundred_x_auxiliary_material_lowers_score(self):
        pred = self._calcium_pred(122.3)
        assert pred.get("confidence_score", 100) <= 20
        assert any("🚨" in w for w in (pred.get("warnings") or []))


class TestPredictScreenshotScenario:
    def test_pvc_price_up_yields_reasonable_per_kg(self):
        """回归：元/件清单 + BOM~1.32倍 时不应出现 130+ 元/kg。"""
        bom_df = pd.DataFrame(
            {
                COL_PRODUCT: ["P1"],
                COL_COMPONENT: ["1010010002"],
                COL_MAKTX: ["聚氯乙烯树脂5型"],
                COL_QTY: [2125.0],
                COL_UNIT_PRICE: [5.0824],
                COL_CREATEDATE: ["20250115"],
            }
        )
        price_df = pd.DataFrame(
            {
                "产品编码": ["P1"],
                "报价月份": ["202501"],
                "重量": [4.0817],
                "材料成本": [23.794],
                "工费": [2.857],
                "总成本": [26.651],
            }
        )
        ref = {"1010010002": 5.0824}
        pred = predict_product_price(
            "P1",
            {"1010010002": 7.0823529},
            bom_df,
            price_df,
            reference_prices=ref,
        )
        assert "error" not in pred
        ppk = float(pred["point_per_kg"])
        assert 6.0 < ppk < 12.0
        sens = pred.get("sensitivity") or []
        if sens:
            lo = sens[0].get("product_price_lo_per_kg")
            hi = sens[0].get("product_price_hi_per_kg")
            if lo is not None and hi is not None:
                assert lo <= ppk <= hi + 0.5


class TestNormalizeOneKgFallback:
    def test_one_kg_product_not_treated_as_per_piece(self):
        norm = normalize_price_list_cost_fields(8.0, 4.0, 12.0, 1.0)
        assert norm["costs_are_per_piece"] is False
        assert norm["total_per_kg"] == pytest.approx(12.0)


class TestBuildProductCostHistoryVersions:
    def test_same_day_uses_latest_quote_not_mean(self):
        raw = pl.DataFrame(
            {
                "产品编码": ["P1", "P1", "P1"],
                "组件编码": ["C1", "C1", "C1"],
                "MENGE": [1.0, 1.0, 1.0],
                "组件单价": [10.0, 20.0, 30.0],
                "CREATEDATE": ["2024-01-15", "2024-01-15", "2024-01-15"],
                "报价号": ["Q1", "Q2", "Q3"],
                "产品价格": [5.0, 5.0, 5.0],
            }
        )
        hist = build_product_cost_history(normalize_columns(raw))
        assert "P1" in hist
        assert hist["P1"][0]["bom_material"] == pytest.approx(30.0)


class TestLinearFitR2:
    def test_r2_upper_capped_at_one(self):
        import numpy as np

        x = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
        y = np.array([2.0, 4.0, 6.0, 8.0, 10.0])
        _, r2, _, _, _ = _linear_fit_predict(x, y, 6.0)
        assert r2 == pytest.approx(1.0)

    def test_grade_invalid_for_negative_r2(self):
        assert _regression_fit_grade(-0.2, 10) == "无效"


class TestBaselinePricesFromSnap:
    def test_shared_with_app_mapping(self):
        snap = {"材料成本": 30.0, "工费": 15.0, "总成本": 45.0, "重量": 2.5}
        bl = baseline_prices_from_price_snap(snap, 2.5)
        assert bl["基准产品价格"] == pytest.approx(45.0)
        assert bl["基准产品价格_每公斤"] == pytest.approx(18.0)


class TestConductionCoeffBounds:
    def test_heavy_product_tighter_max(self):
        lo, hi = _conduction_coeff_clip_bounds(20.0)
        assert hi == pytest.approx(0.1)
        assert lo < hi


class TestPoolSlopesReservoir:
    def test_caps_at_slope_pool_max(self):
        hist = {}
        for i in range(80):
            pts = []
            for m in range(6):
                pts.append({
                    "date": f"20240{m+1:02d}",
                    "bom_material_per_kg": 10.0 + i * 0.01 + m,
                    "product_price_per_kg": 20.0 + i * 0.02 + m * 2,
                    "重量_kg": 2.0,
                })
            hist[f"P{i:03d}"] = pts
        slopes, _, _, n_prod = _pool_slopes_from_history(hist)
        assert n_prod == 80
        assert len(slopes) <= SLOPE_POOL_MAX


class TestModelErrorInterval:
    def test_no_perturbation_matches_fixed_mae(self):
        raw = {"可用": True, "MAE": 2.0, "RMSE": 2.1}
        out = _model_error_interval(100.0, raw, bom_perturbation_ratio=0.0)
        assert out["区间半宽"] == pytest.approx(MODEL_ERROR_Z * 2.0)
        assert out["区间放大系数"] == pytest.approx(1.0)
        assert out["BOM扰动比例"] == pytest.approx(0.0)

    def test_large_perturbation_widens_interval(self):
        raw = {"可用": True, "MAE": 2.0}
        base_half = MODEL_ERROR_Z * 2.0
        out = _model_error_interval(100.0, raw, bom_perturbation_ratio=3.0)
        expected_scale = 1.0 + MODEL_ERROR_GAMMA * 3.0
        assert out["区间放大系数"] == pytest.approx(expected_scale)
        assert out["区间半宽"] == pytest.approx(base_half * expected_scale)
        assert out["预测区间_kg"][0] == pytest.approx(100.0 - out["区间半宽"])
