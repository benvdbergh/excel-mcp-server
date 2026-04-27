"""Guards for optional COM stack detection (Story 6-1)."""

from __future__ import annotations

import sys

from excel_mcp import com_support


def test_com_support_module_imports_without_error():
    assert isinstance(com_support.COM_STACK_AVAILABLE, bool)
    assert com_support.is_com_runtime_supported() == com_support.COM_STACK_AVAILABLE


def test_non_windows_never_reports_com_available():
    if sys.platform == "win32":
        return
    assert com_support.COM_STACK_AVAILABLE is False
    assert com_support.is_com_runtime_supported() is False
