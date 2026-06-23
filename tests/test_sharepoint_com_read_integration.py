"""BEN-139: integration tests — SharePoint FullName routing + COM read fidelity (mocked).

Exercises ``server.read_data_from_excel`` end-to-end with mocked win32 COM (no Excel
required; safe on Linux CI per NFR-6). Covers BEN-131 identity matching, formatted
currency (``value_mode=text``), and sparse null ``Value2`` fallback from MyMA sessions.

For live Excel + SharePoint verification see
``docs/plan/transport-routing/MANUAL-WINDOWS-RC-CHECKLIST.md``.
"""

from __future__ import annotations

import json
import os
import sys
from contextlib import contextmanager
from typing import Iterator
from unittest.mock import MagicMock, patch

import pytest

from excel_mcp.path_resolution import parse_cloud_workbook_locator
from excel_mcp.routing.com_workbook_open_detection import ComWorkbookOpenInExcel
from excel_mcp.routing.com_workbook_service import ComWorkbookService
from excel_mcp.routing.routing_backend import RoutingBackend

from test_com_workbook_service import (  # noqa: E402
    ImmediateExecutor,
    _fake_win32_modules,
    _workbook_mock,
)

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

_OPERATOR_URL = "HTTPS://tenant.SharePoint.com/sites/s/Shared%20Documents/book.xlsx"
_EXCEL_FULLNAME = "https://tenant.sharepoint.com/sites/s/Shared Documents/book.xlsx"
_SHEET = "Sheet"


def _xl_with_workbook(excel_fullname: str, ws: MagicMock) -> MagicMock:
    wb = _workbook_mock(excel_fullname, {_SHEET: ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)
    return xl


def _used_range(rows: int, cols: int) -> MagicMock:
    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = rows
    used.Columns.Count = cols
    return used


def _cell(value2, *, text: str | None = None) -> MagicMock:
    cell = MagicMock()
    cell.Value2 = value2
    cell.Text = text if text is not None else value2
    cell.Validation = MagicMock()
    cell.Validation.Type = 0
    return cell


@contextmanager
def _patched_com_server(monkeypatch: pytest.MonkeyPatch, xl: MagicMock) -> Iterator:
    import excel_mcp.server as srv

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        com_svc = ComWorkbookService(ImmediateExecutor())
        rb = RoutingBackend(
            ComWorkbookOpenInExcel(ImmediateExecutor()),
            com_execution_available=True,
            runtime_platform="win32",
        )
        monkeypatch.setitem(srv.__dict__, "_COM_WORKBOOK_SERVICE", com_svc)
        monkeypatch.setitem(srv.__dict__, "_ROUTING_BACKEND", rb)
        yield srv


@pytest.fixture
def cloud_allowlist(monkeypatch: pytest.MonkeyPatch) -> None:
    import excel_mcp.server as srv

    srv.EXCEL_FILES_PATH = None
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_PATHS", os.getcwd())
    monkeypatch.setenv(
        "EXCEL_MCP_ALLOWED_URL_PREFIXES",
        "https://tenant.sharepoint.com/sites/s/",
    )


def test_sharepoint_auto_routes_com_formatted_currency_text_mode(
    cloud_allowlist: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Operator locator differs from Excel FullName; auto → COM; text mode returns display."""
    rng = MagicMock()
    rng.Text = "19,900.00 €"
    rng.Value2 = 19900
    cell = _cell(19900, text="19,900.00 €")
    cell.Resize = MagicMock(return_value=rng)

    ws = MagicMock()
    ws.UsedRange = _used_range(1, 1)
    ws.Cells = MagicMock(return_value=cell)
    xl = _xl_with_workbook(_EXCEL_FULLNAME, ws)

    with _patched_com_server(monkeypatch, xl) as srv:
        out = srv.read_data_from_excel(
            _OPERATOR_URL,
            _SHEET,
            "A1",
            "A1",
            workbook_transport="auto",
            value_mode="text",
            include_routing_metadata=True,
        )

    envelope = json.loads(out)
    assert envelope["_meta"]["workbook_backend"] == "com"
    assert envelope["_meta"]["routing_reason"] == "full_name_match"
    assert envelope["result"]["value_mode"] == "text"
    assert envelope["result"]["cells"][0]["value"] == "19,900.00 €"


def test_sharepoint_forced_com_value_mode_returns_raw_value2(
    cloud_allowlist: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Default value_mode keeps Value2 (raw number), not formatted display text."""
    rng = MagicMock()
    rng.Value2 = 19900
    rng.Text = "19,900.00 €"
    cell = _cell(19900, text="19,900.00 €")
    cell.Resize = MagicMock(return_value=rng)

    ws = MagicMock()
    ws.UsedRange = _used_range(1, 1)
    ws.Cells = MagicMock(return_value=cell)
    xl = _xl_with_workbook(_EXCEL_FULLNAME, ws)

    canonical = parse_cloud_workbook_locator(_OPERATOR_URL)
    with _patched_com_server(monkeypatch, xl) as srv:
        out = srv.read_data_from_excel(
            canonical,
            _SHEET,
            "A1",
            "A1",
            workbook_transport="com",
        )

    data = json.loads(out)
    assert data["value_mode"] == "value"
    assert data["cells"][0]["value"] == 19900


def test_sharepoint_com_sparse_null_value2_fallback(
    cloud_allowlist: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Bulk Value2 all-null matrix triggers per-cell fallback (MyMA sparse read)."""
    rng = MagicMock()
    rng.Value2 = ((None, None), (None, None))

    direct_vals = {
        (1, 1): "A11",
        (1, 2): "B11",
        (2, 1): "A12",
        (2, 2): "B12",
    }
    cells: dict[tuple[int, int], MagicMock] = {}
    for coord, val in direct_vals.items():
        cells[coord] = _cell(val)
    cells[(1, 1)].Resize = MagicMock(return_value=rng)

    ws = MagicMock()
    ws.UsedRange = _used_range(2, 2)
    ws.Cells = MagicMock(side_effect=lambda r, c: cells[(r, c)])
    xl = _xl_with_workbook(_EXCEL_FULLNAME, ws)

    with _patched_com_server(monkeypatch, xl) as srv:
        out = srv.read_data_from_excel(
            _OPERATOR_URL,
            _SHEET,
            "A1",
            "B2",
            workbook_transport="auto",
            include_routing_metadata=True,
        )

    envelope = json.loads(out)
    assert envelope["_meta"]["workbook_backend"] == "com"
    values = [c["value"] for c in envelope["result"]["cells"]]
    assert values == ["A11", "B11", "A12", "B12"]


def test_sharepoint_text_mode_bulk_text_fallback(
    cloud_allowlist: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When bulk Range.Text is null, per-cell Text fallback still returns currency display."""
    rng = MagicMock()
    rng.Text = None
    rng.Value2 = 19900
    cell = _cell(19900, text="19,900.00 €")
    cell.Resize = MagicMock(return_value=rng)

    ws = MagicMock()
    ws.UsedRange = _used_range(1, 1)
    ws.Cells = MagicMock(return_value=cell)
    xl = _xl_with_workbook(_EXCEL_FULLNAME, ws)

    with _patched_com_server(monkeypatch, xl) as srv:
        out = srv.read_data_from_excel(
            _OPERATOR_URL,
            _SHEET,
            "A1",
            "A1",
            workbook_transport="auto",
            value_mode="text",
        )

    data = json.loads(out)
    assert data["cells"][0]["value"] == "19,900.00 €"


@pytest.mark.requires_excel
def test_manual_windows_sharepoint_formatted_cells() -> None:
    """Placeholder: run against live Excel per MANUAL-WINDOWS-RC-CHECKLIST.md."""
    pytest.skip(
        "Manual Windows RC only — see docs/plan/transport-routing/"
        "MANUAL-WINDOWS-RC-CHECKLIST.md"
    )
