"""Tests for ``ComWorkbookOpenInExcel`` (COM routing open detection)."""

from __future__ import annotations

import sys
import types
from unittest.mock import MagicMock, patch

import pytest

from excel_mcp.routing.com_workbook_open_detection import ComWorkbookOpenInExcel


class ImmediateExecutor:
    def submit(self, fn, /, *args, **kwargs):
        return fn(*args, **kwargs)


def _fake_win32(xl_app: MagicMock) -> dict[str, types.ModuleType]:
    client_mod = types.ModuleType("win32com.client")
    client_mod.GetActiveObject = lambda *_a, **_kw: xl_app
    pkg = types.ModuleType("win32com")
    pkg.client = client_mod
    return {"win32com": pkg, "win32com.client": client_mod}


@pytest.fixture
def book_path(tmp_path):
    p = tmp_path / "wb.xlsx"
    return str(p.resolve())


def test_no_excel_returns_false(book_path):
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0
    xl.Workbooks.Item = MagicMock(side_effect=RuntimeError("no app"))
    with patch.dict(sys.modules, _fake_win32(xl), clear=False):
        port = ComWorkbookOpenInExcel(ImmediateExecutor())
        assert port.is_workbook_open_in_excel(book_path) is False


def test_get_active_object_fails_returns_false(book_path):
    def boom(*_a, **_kw):
        raise RuntimeError("no Excel")

    client_mod = types.ModuleType("win32com.client")
    client_mod.GetActiveObject = boom
    pkg = types.ModuleType("win32com")
    pkg.client = client_mod
    with patch.dict(sys.modules, {"win32com": pkg, "win32com.client": client_mod}, clear=False):
        port = ComWorkbookOpenInExcel(ImmediateExecutor())
        assert port.is_workbook_open_in_excel(book_path) is False


def test_one_matching_workbook_true(book_path):
    wb = MagicMock()
    wb.FullName = book_path
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32(xl), clear=False):
        port = ComWorkbookOpenInExcel(ImmediateExecutor())
        assert port.is_workbook_open_in_excel(book_path) is True


def test_two_matching_workbooks_false(book_path):
    wb1 = MagicMock()
    wb1.FullName = book_path
    wb2 = MagicMock()
    wb2.FullName = book_path
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 2
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb1 if i == 1 else wb2)

    with patch.dict(sys.modules, _fake_win32(xl), clear=False):
        port = ComWorkbookOpenInExcel(ImmediateExecutor())
        assert port.is_workbook_open_in_excel(book_path) is False


def test_one_other_path_false(book_path, tmp_path):
    other = str((tmp_path / "other.xlsx").resolve())
    wb = MagicMock()
    wb.FullName = other
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32(xl), clear=False):
        port = ComWorkbookOpenInExcel(ImmediateExecutor())
        assert port.is_workbook_open_in_excel(book_path) is False
