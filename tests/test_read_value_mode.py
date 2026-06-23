"""Tests for read value_mode validation."""

from __future__ import annotations

import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.read_value_mode import validate_value_mode  # noqa: E402


@pytest.mark.parametrize("mode", ["value", "text"])
def test_validate_value_mode_accepts_known_modes(mode: str) -> None:
    assert validate_value_mode(mode) == mode


def test_validate_value_mode_rejects_unknown() -> None:
    with pytest.raises(ValueError, match="Invalid value_mode 'raw'"):
        validate_value_mode("raw")
