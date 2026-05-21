# -*- coding: utf-8 -*-
"""
日历日期键：将「生价日期」与 CREATEDATE 统一为 YYYYMMDD 字符串，便于跨表对齐合并。
"""
from __future__ import annotations

import math
from datetime import date, datetime

import pandas as pd


def calendar_date_key_yyyymmdd(x) -> str:
    """
    将 EXPORT 的「生价日期」、历史清单的 CREATEDATE 等统一为 8 位 YYYYMMDD。

    支持：整数/浮点 20260107、Excel 序列日期、pd.Timestamp、datetime、
    字符串 '2026-01-07' / '20260107' / '20260107.0' 等。
    无法解析时返回空字符串。
    """
    if x is None:
        return ""
    if isinstance(x, bool):
        return ""
    try:
        if pd.isna(x):
            return ""
    except (ValueError, TypeError):
        pass
    if isinstance(x, float) and not math.isfinite(x):
        return ""

    if isinstance(x, pd.Timestamp):
        return x.strftime("%Y%m%d")
    if isinstance(x, datetime):
        return x.strftime("%Y%m%d")
    if isinstance(x, date):
        return x.strftime("%Y%m%d")

    if isinstance(x, int) and not isinstance(x, bool):
        xi = int(x)
        if 19000101 <= xi <= 21001231:
            return f"{xi:08d}"

    if isinstance(x, float):
        fx = float(x)
        if 35000 < fx < 60000:
            try:
                return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=fx)).strftime("%Y%m%d")
            except Exception:
                pass
        if abs(fx - round(fx)) < 1e-9:
            xi = int(round(fx))
            if 19000101 <= xi <= 21001231:
                return f"{xi:08d}"

    s = str(x).strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        return calendar_date_key_yyyymmdd(s[:-2])

    head = s.replace("-", "").replace("/", "").replace(" ", "")
    if len(head) >= 8 and head[:8].isdigit():
        return head[:8]

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        return ""
    return pd.Timestamp(parsed).strftime("%Y%m%d")
