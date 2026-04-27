"""STORY-11-3: read-class handlers pass ``com_do_op`` so COM routing is wired (not null)."""

from __future__ import annotations

import ast
import os
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing import ComWorkbookService, RoutingBackend, StubWorkbookOpenInExcel  # noqa: E402
from excel_mcp.routing.tool_inventory import (  # noqa: E402
    MCP_TOOL_INVENTORY,
    ToolKind,
)


class _ImmediateExecutor:
    def submit(self, fn, /, *args, **kwargs):
        return fn(*args, **kwargs)


def _read_mcp_tool_names() -> set[str]:
    return {
        name
        for name, entry in MCP_TOOL_INVENTORY.items()
        if entry.kind == ToolKind.READ
    }


def test_workbook_dispatch_read_tools_include_com_do_op_keyword() -> None:
    """Contract: every ToolKind.READ _workbook_dispatch call passes com_do_op=."""
    server_path = (
        Path(__file__).resolve().parent.parent
        / "src"
        / "excel_mcp"
        / "server.py"
    )
    tree = ast.parse(server_path.read_text(encoding="utf-8"))
    read_names = _read_mcp_tool_names()
    missing: list[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "_workbook_dispatch":
            continue
        if not node.args:
            continue
        first = node.args[0]
        if not isinstance(first, ast.Constant) or not isinstance(first.value, str):
            continue
        if first.value not in read_names:
            continue
        kw = {k.arg for k in node.keywords if k.arg is not None}
        if "com_do_op" not in kw:
            missing.append(first.value)
    assert not missing, f"READ tools missing com_do_op keyword: {sorted(missing)}"


def test_get_workbook_metadata_com_transport_wired_no_com_not_implemented_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """With COM backend selected, read handler must not raise missing-wiring error."""
    p = tmp_path / "com_wiring.xlsx"
    Workbook().save(p)
    path = str(p.resolve())

    import excel_mcp.server as srv

    monkeypatch.setitem(
        srv.__dict__,
        "_COM_WORKBOOK_SERVICE",
        ComWorkbookService(_ImmediateExecutor()),  # type: ignore[arg-type]
    )
    rb = RoutingBackend(
        StubWorkbookOpenInExcel(),
        com_execution_available=True,
        runtime_platform="win32",
    )
    monkeypatch.setitem(srv.__dict__, "_ROUTING_BACKEND", rb)

    out = srv.get_workbook_metadata(path, workbook_transport="com", include_ranges=False)
    # COM read is implemented; without Excel hosting the workbook, expect a stable open-path error.
    assert "COM path not implemented" not in out
    assert "Workbook not open in Excel" in out
