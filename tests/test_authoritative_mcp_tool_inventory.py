"""Authoritative MCP tool inventory (read/write/v1 exception)."""

import os
import re
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.tool_inventory import (  # noqa: E402
    MCP_TOOL_INVENTORY,
    ToolKind,
    get_tool_kind,
)

_EXPECTED_TOOL_NAMES = frozenset(
    {
        "apply_formula",
        "validate_formula_syntax",
        "format_range",
        "read_data_from_excel",
        "write_data_to_excel",
        "create_workbook",
        "save_workbook",
        "create_worksheet",
        "create_chart",
        "create_pivot_table",
        "create_table",
        "copy_worksheet",
        "delete_worksheet",
        "rename_worksheet",
        "get_workbook_metadata",
        "merge_cells",
        "unmerge_cells",
        "get_merged_cells",
        "copy_range",
        "delete_range",
        "validate_excel_range",
        "get_data_validation_info",
        "insert_rows",
        "insert_columns",
        "delete_sheet_rows",
        "delete_sheet_columns",
        "excel_open_workbook",
        "excel_close_workbook",
    }
)

_NAME_RE = re.compile(r"^[a-z][a-z0-9_]*$")


def test_inventory_has_exactly_28_keys() -> None:
    assert len(MCP_TOOL_INVENTORY) == 28


def test_every_key_matches_expected_set_or_pattern() -> None:
    keys = frozenset(MCP_TOOL_INVENTORY)
    assert keys == _EXPECTED_TOOL_NAMES
    for name in MCP_TOOL_INVENTORY:
        assert _NAME_RE.match(name), name


@pytest.mark.parametrize("tool_name", sorted(_EXPECTED_TOOL_NAMES))
def test_get_tool_kind_for_each_registered_tool(tool_name: str) -> None:
    kind = get_tool_kind(tool_name)
    assert kind is MCP_TOOL_INVENTORY[tool_name].kind


def test_get_tool_kind_unknown_raises() -> None:
    with pytest.raises(KeyError):
        get_tool_kind("not_a_real_excel_mcp_tool")


def test_chart_and_pivot_are_v1_file_forced() -> None:
    assert get_tool_kind("create_chart") is ToolKind.V1_FILE_FORCED
    assert get_tool_kind("create_pivot_table") is ToolKind.V1_FILE_FORCED


def test_at_least_one_read_tool_is_read() -> None:
    assert get_tool_kind("read_data_from_excel") is ToolKind.READ


def test_lifecycle_tools_are_session() -> None:
    assert get_tool_kind("excel_open_workbook") is ToolKind.SESSION
    assert get_tool_kind("excel_close_workbook") is ToolKind.SESSION
