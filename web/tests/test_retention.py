# -*- coding: utf-8 -*-
"""Tests for upload/results file retention purge."""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import pytest

_WEB = Path(__file__).resolve().parent.parent
if str(_WEB) not in sys.path:
    sys.path.insert(0, str(_WEB))

from app import purge_stale_files  # noqa: E402


def test_purge_stale_files_removes_old_only(tmp_path: Path) -> None:
    old = tmp_path / "old.txt"
    new = tmp_path / "new.txt"
    old.write_text("x", encoding="utf-8")
    new.write_text("y", encoding="utf-8")
    old_ts = time.time() - 10 * 86400
    os.utime(old, (old_ts, old_ts))

    removed = purge_stale_files(tmp_path, max_age_days=7)
    assert removed == 1
    assert not old.exists()
    assert new.exists()


def test_purge_stale_files_empty_dir(tmp_path: Path) -> None:
    assert purge_stale_files(tmp_path, max_age_days=7) == 0
