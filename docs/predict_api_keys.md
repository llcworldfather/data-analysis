# Cost predict API — internal English keys

`predict_product_price` and `_predict_product_price_legacy` return **English** keys internally.  
Flask [`_map_new_predict_result`](../web/app.py) / `map_legacy_predict_en_to_zh` map them to Chinese for the UI.

## Environment (optional)

| Variable | Default | Meaning |
|----------|---------|---------|
| `PER_PIECE_TOTAL_THRESHOLD_YUAN` | `12` | When **MEINS / 基本单位** is missing, treat material+labor sum above this (CNY) as **元/件** instead of **元/KG**. |
| `DEEPSEEK_API_KEY` | — | DeepSeek chat / purchase-advice (see `web/.env`). |
| `BAIDU_SEARCH_API_KEY` | — | Baidu Qianfan AI search for market tools. |

Cost unit detection order: **MEINS / 基本单位** (`PC` → 元/件, `KG` → 元/KG) → consistency of 材料+工费 vs 总成本 vs 重量 → threshold above.

## `predict_product_price` response

| Key | Meaning |
|-----|---------|
| `method` | Primary prediction method id |
| `point_estimate` | Product total price (CNY/piece) |
| `point_per_kg` | Product price per kg |
| `confidence_score` | 0–100 |
| `warnings` | List of warning strings |
| `sensitivity` | List of single-factor sensitivity items |
| `sensitivity_grid` | Top-2 modified components 3×3 grid |
| `model_error` | Historical backtest MAE interval |
| `detail` | Base BOM, weight, components |

### `sensitivity[]` item

- `material_code`, `material_name`, `user_price`
- `price_range`, `price_swing_pct`
- `product_price_kg_range`, `product_price_change_pct`

### `sensitivity_grid`

- `available`, `components[]`, `pct_steps`, `baseline_per_kg`, `matrix`

### `model_error` (Chinese keys in payload; passed through to UI)

Historical backtest supplies robust **MAE**; display interval half-width scales with simulated BOM drift:

`half = MODEL_ERROR_Z × MAE × (1 + MODEL_ERROR_GAMMA × |ΔBOM| / base_BOM)`

| Key | Meaning |
|-----|---------|
| `可用` | Whether backtest interval is available |
| `MAE` | Median absolute error (CNY/kg), from history only |
| `样本数` | Backtest months used |
| `预测区间_kg` | `[lo, hi]` around current `point_per_kg` |
| `区间半宽` | Half-width after BOM perturbation scaling |
| `BOM扰动比例` | `|sim_bom − base_bom| / max(base_bom, MIN)` |
| `区间放大系数` | `1 + γ × BOM扰动比例` |
| `区间口径` | Human-readable interval formula note |

### `regression_analysis` (legacy path only)

- `enabled`, `used_as_primary`, `valid_sample_count`, `r2`, `fit_grade`, …
