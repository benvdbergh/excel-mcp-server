"""Tests for routed dispatch structured logging (STORY-4-3)."""

from __future__ import annotations

import json
import logging
import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.routed_dispatch import (  # noqa: E402
    execute_routed_workbook_operation,
    redact_workbook_path_for_logs,
)
from excel_mcp.routing.routing_backend import RoutingBackend  # noqa: E402
from excel_mcp.routing.routing_errors import ComExecutionNotImplementedError  # noqa: E402
from excel_mcp.routing.tool_inventory import ToolKind, get_tool_kind  # noqa: E402


class _FakeWorkbookOpen:
    def __init__(self, open_paths: frozenset[str]) -> None:
        self._open_paths = open_paths

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        return resolved_path in self._open_paths


class _DummyFileService:
    """Stand-in ``RoutedWorkbookOperations``; Epic 4 dispatch runs ``operation_callable`` only."""

    pass


_PATH = r"C:\Users\secret\subdir\book.xlsx"
_DUMMY = _DummyFileService()


def _last_json_record(caplog: pytest.LogCaptureFixture) -> dict:
    for rec in reversed(caplog.records):
        if rec.name == "excel-mcp.routing" and rec.levelno == logging.INFO:
            try:
                return json.loads(rec.getMessage())
            except json.JSONDecodeError:
                continue
    raise AssertionError("no JSON routing log on excel-mcp.routing")


def test_dispatch_logs_required_fields_and_redacts_path(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    rb = RoutingBackend(_FakeWorkbookOpen(frozenset()), runtime_platform="win32")
    out, backend = execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=_PATH,
        workbook_transport="file",
        tool_kind=ToolKind.READ,
        com_strict=True,
        operation_name="workbook_metadata",
        operation_callable=lambda: '{"ok": true}',
        mcp_tool_name="get_workbook_metadata",
    )
    assert out == '{"ok": true}'
    assert backend == "file"
    data = _last_json_record(caplog)
    assert data["workbook_transport"] == "file"
    assert data["workbook_backend"] == "file"
    assert data["routing_reason"] == "forced_file"
    assert data["operation_name"] == "workbook_metadata"
    assert data["mcp_tool_name"] == "get_workbook_metadata"
    assert data["workbook_path"] == "book.xlsx"
    assert "duration_ms" in data
    assert isinstance(data["duration_ms"], (int, float))
    assert data["duration_ms"] >= 0


def test_duration_includes_callable_time(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    rb = RoutingBackend(_FakeWorkbookOpen(frozenset()), runtime_platform="win32")

    def _slow() -> str:
        import time as _t

        _t.sleep(0.06)
        return "done"

    execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=_PATH,
        workbook_transport="file",
        tool_kind=ToolKind.READ,
        com_strict=True,
        operation_name="workbook_metadata",
        operation_callable=_slow,
    )
    data = _last_json_record(caplog)
    assert float(data["duration_ms"]) >= 50.0


def test_com_backend_logs_then_raises(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    called = {"n": 0}

    def _should_not_run() -> str:
        called["n"] += 1
        return "bad"

    with pytest.raises(ComExecutionNotImplementedError) as ei:
        execute_routed_workbook_operation(
            rb,
            _DUMMY,
            resolved_path=_PATH,
            workbook_transport="auto",
            tool_kind=ToolKind.WRITE,
            com_strict=False,
            operation_name="workbook_metadata",
            operation_callable=_should_not_run,
            mcp_tool_name="get_workbook_metadata",
        )
    assert ComExecutionNotImplementedError.STABLE_TOKEN in str(ei.value)
    assert called["n"] == 0
    data = _last_json_record(caplog)
    assert data["workbook_backend"] == "com"
    assert data["routing_reason"] == "full_name_match"
    assert data["workbook_transport"] == "auto"
    assert data["operation_name"] == "workbook_metadata"


def test_com_backend_invokes_callable_no_not_implemented_error(
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    sentinel = "SENTINEL_COM_OK"

    def _no_file() -> str:
        raise AssertionError("file operation_callable must not run for COM backend")

    out, backend = execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=_PATH,
        workbook_transport="auto",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
        operation_name="write_cell_grid",
        operation_callable=_no_file,
        com_operation_callable=lambda: sentinel,
        mcp_tool_name="write_data_to_excel",
    )
    assert out == sentinel
    assert backend == "com"
    data = _last_json_record(caplog)
    assert data["workbook_backend"] == "com"
    assert data["routing_reason"] == "full_name_match"


@pytest.mark.parametrize(
    ("mcp_tool_name", "operation_name"),
    [
        ("create_chart", "create_chart_in_sheet"),
        ("create_pivot_table", "create_pivot_table_in_sheet"),
    ],
)
def test_dispatch_adr0004_v1_file_forced_auto_open_com_logs_flag(
    caplog: pytest.LogCaptureFixture,
    mcp_tool_name: str,
    operation_name: str,
) -> None:
    """Open workbook + viable COM + transport auto still runs file; log marks ADR 0004."""
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    out, backend = execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=_PATH,
        workbook_transport="auto",
        tool_kind=get_tool_kind(mcp_tool_name),
        com_strict=False,
        operation_name=operation_name,
        operation_callable=lambda: '{"ok": true}',
        mcp_tool_name=mcp_tool_name,
    )
    assert out == '{"ok": true}'
    assert backend == "file"
    data = _last_json_record(caplog)
    assert data["workbook_backend"] == "file"
    assert data["routing_reason"] == "v1_file_forced"
    assert data["v1_file_forced"] is True
    assert data["workbook_transport"] == "auto"
    assert data["mcp_tool_name"] == mcp_tool_name


def test_redact_basename_default() -> None:
    assert redact_workbook_path_for_logs(r"E:\a\b\c.xlsx") == "c.xlsx"


def test_redact_full_path_when_env_set(monkeypatch: pytest.MonkeyPatch) -> None:
    p = r"E:\a\b\c.xlsx"
    monkeypatch.setenv("EXCEL_MCP_LOG_FULL_PATHS", "1")
    assert redact_workbook_path_for_logs(p) == p


def test_invalid_operation_name() -> None:
    rb = RoutingBackend(_FakeWorkbookOpen(frozenset()), runtime_platform="win32")
    with pytest.raises(ValueError, match="operation_name"):
        execute_routed_workbook_operation(
            rb,
            _DUMMY,
            resolved_path=_PATH,
            workbook_transport="file",
            tool_kind=ToolKind.READ,
            com_strict=True,
            operation_name="not_a_real_operation",
            operation_callable=lambda: "",
        )


def test_routing_package_exports_dispatch_helpers() -> None:
    from excel_mcp.routing import (  # noqa: E402
        execute_routed_workbook_operation as ex,
        redact_workbook_path_for_logs as red,
    )

    assert ex is execute_routed_workbook_operation
    assert red is redact_workbook_path_for_logs
