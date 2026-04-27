"""Tests for ``RoutingBackend`` / ``resolve_workbook_backend`` (STORY-4-2)."""

from __future__ import annotations

import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.routing_backend import (  # noqa: E402
    RoutingBackend,
    WorkbookBackendResolution,
)
from excel_mcp.routing.routing_errors import ComRoutingError  # noqa: E402
from excel_mcp.routing.tool_inventory import ToolKind, get_tool_kind  # noqa: E402


class _FakeWorkbookOpen:
    def __init__(self, open_paths: frozenset[str] | None = None, *, always: bool | None = None) -> None:
        self._open_paths = open_paths or frozenset()
        self._always = always

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        if self._always is not None:
            return self._always
        return resolved_path in self._open_paths


_PATH = r"C:\tmp\book.xlsx"


def test_auto_closed_workbook_uses_file() -> None:
    rb = RoutingBackend(_FakeWorkbookOpen(frozenset()), runtime_platform="win32")
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind=ToolKind.WRITE,
        com_strict=True,
    )
    assert r == WorkbookBackendResolution(
        backend="file",
        reason="auto_workbook_not_open_file",
        requested_transport="auto",
    )


def test_read_auto_open_workbook_stays_file_adr0003() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind=ToolKind.READ,
        com_strict=False,
    )
    assert r.backend == "file"
    assert r.reason == "read_class_file_backed"
    assert r.requested_transport == "auto"


def test_read_com_transport_stays_file_adr0003() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="com",
        tool_kind=ToolKind.READ,
        com_strict=True,
    )
    assert r.backend == "file"
    assert r.reason == "read_class_file_backed"


def test_auto_open_workbook_uses_com_when_viable() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
    )
    assert r.backend == "com"
    assert r.reason == "full_name_match"
    assert r.requested_transport == "auto"


def test_forced_file_transport() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="file",
        tool_kind=ToolKind.WRITE,
        com_strict=True,
    )
    assert r == WorkbookBackendResolution(
        backend="file",
        reason="forced_file",
        requested_transport="file",
    )


def test_com_strict_when_execution_unavailable_raises() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(always=True),
        com_execution_available=False,
        runtime_platform="win32",
    )
    with pytest.raises(ComRoutingError) as excinfo:
        rb.resolve_workbook_backend(
            resolved_path=_PATH,
            transport="com",
            tool_kind=ToolKind.WRITE,
            com_strict=True,
        )
    assert ComRoutingError.STABLE_TOKEN in str(excinfo.value)
    assert excinfo.value.reason_code == "com_execution_unavailable"


def test_com_non_strict_fallback_when_unavailable() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(always=True),
        com_execution_available=False,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="com",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
    )
    assert r.backend == "file"
    assert r.reason == "com_unavailable_file_fallback"


@pytest.mark.parametrize(
    "tool_kind",
    [
        ToolKind.V1_FILE_FORCED,
        get_tool_kind("create_chart"),
        get_tool_kind("create_pivot_table"),
    ],
    ids=["v1_file_forced_enum", "create_chart_inventory", "create_pivot_table_inventory"],
)
def test_adr0004_v1_file_forced_auto_open_com_viable_stays_file(
    tool_kind: ToolKind,
) -> None:
    """ADR 0004: chart/pivot must not drift to COM when auto + open + COM viable."""
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind=tool_kind,
        com_strict=False,
    )
    assert r == WorkbookBackendResolution(
        backend="file",
        reason="v1_file_forced",
        requested_transport="auto",
    )


def test_v1_file_forced_as_string() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind="v1_file_forced",
        com_strict=False,
    )
    assert r.reason == "v1_file_forced"


def test_non_windows_auto_open_strict_raises() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="linux",
    )
    with pytest.raises(ComRoutingError) as excinfo:
        rb.resolve_workbook_backend(
            resolved_path=_PATH,
            transport="auto",
            tool_kind=ToolKind.WRITE,
            com_strict=True,
        )
    assert ComRoutingError.STABLE_TOKEN in str(excinfo.value)
    assert excinfo.value.reason_code == "com_unsupported_non_windows"


def test_non_windows_auto_open_non_strict_file_fallback() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="linux",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="auto",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
    )
    assert r.backend == "file"
    assert r.reason == "com_unsupported_non_windows"


def test_com_transport_forced_com_when_viable_and_open() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset({_PATH})),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="com",
        tool_kind=ToolKind.WRITE,
        com_strict=True,
    )
    assert r.backend == "com"
    assert r.reason == "forced_com"


def test_com_transport_forced_com_when_viable_even_if_not_open() -> None:
    """Explicit ``transport=com`` selects COM whenever viable (open port ignored)."""
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset()),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="com",
        tool_kind=ToolKind.WRITE,
        com_strict=True,
    )
    assert r.backend == "com"
    assert r.reason == "forced_com"


def test_com_non_strict_still_com_when_viable_not_open() -> None:
    rb = RoutingBackend(
        _FakeWorkbookOpen(frozenset()),
        com_execution_available=True,
        runtime_platform="win32",
    )
    r = rb.resolve_workbook_backend(
        resolved_path=_PATH,
        transport="com",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
    )
    assert r.backend == "com"
    assert r.reason == "forced_com"


def test_routing_package_exports_routing_backend_symbols() -> None:
    from excel_mcp.routing import (  # noqa: E402
        ComRoutingError as CRE,
        RoutingBackend as RB,
        WorkbookBackendResolution as WBR,
    )

    assert RB is RoutingBackend
    assert WBR is WorkbookBackendResolution
    assert CRE is ComRoutingError
