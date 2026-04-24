"""STORY-1-2: shared workbook operation contract (Protocol) for routed backends."""

import os
import sys
from typing import Any, Dict, List, Optional

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.workbook_operation_contract import (  # noqa: E402
    ROUTED_WORKBOOK_OPERATION_NAMES,
    RoutedWorkbookOperations,
    WorkbookOperationMetadata,
    WorkbookReadOperations,
    WorkbookWriteOperations,
)


def test_module_exports_importable() -> None:
    assert WorkbookReadOperations is not None
    assert WorkbookWriteOperations is not None
    assert RoutedWorkbookOperations is not None
    assert WorkbookOperationMetadata is not None
    assert len(ROUTED_WORKBOOK_OPERATION_NAMES) > 0


class _AllRoutedOpsDummy:
    """Minimal stand-in for structural checks (no server, no I/O)."""

    def read_range_with_metadata(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        preview_only: bool = False,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def workbook_metadata(
        self,
        filepath: str,
        include_ranges: bool = False,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def read_merged_cell_ranges(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def read_worksheet_data_validation(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def validate_sheet_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def validate_formula_syntax(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def apply_formula(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def format_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: Optional[int] = None,
        font_color: Optional[str] = None,
        bg_color: Optional[str] = None,
        border_style: Optional[str] = None,
        border_color: Optional[str] = None,
        number_format: Optional[str] = None,
        alignment: Optional[str] = None,
        wrap_text: bool = False,
        merge_cells: bool = False,
        protection: Optional[Dict[str, Any]] = None,
        conditional_format: Optional[Dict[str, Any]] = None,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def write_cell_grid(
        self,
        filepath: str,
        sheet_name: str,
        data: List[List[Any]],
        start_cell: str = "A1",
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def create_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def create_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def create_chart_in_sheet(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        title: str = "",
        x_axis: str = "",
        y_axis: str = "",
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def create_pivot_table_in_sheet(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        rows: List[str],
        values: List[str],
        columns: Optional[List[str]] = None,
        agg_func: str = "mean",
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def create_excel_table(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        table_name: Optional[str] = None,
        table_style: str = "TableStyleMedium9",
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def copy_worksheet(
        self,
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def delete_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def rename_worksheet(
        self,
        filepath: str,
        old_name: str,
        new_name: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def merge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def unmerge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def copy_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: Optional[str] = None,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def delete_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        shift_direction: str = "up",
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def insert_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def insert_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def delete_sheet_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def delete_sheet_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""

    def save_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[dict[str, Any]] = None,
    ) -> str:
        return ""


def test_routed_operation_names_cover_dummy_surface() -> None:
    dummy = _AllRoutedOpsDummy()
    for name in ROUTED_WORKBOOK_OPERATION_NAMES:
        assert hasattr(dummy, name), name
        assert callable(getattr(dummy, name)), name


@pytest.mark.parametrize("method_name", ROUTED_WORKBOOK_OPERATION_NAMES)
def test_dummy_each_method_callable(method_name: str) -> None:
    dummy = _AllRoutedOpsDummy()
    assert callable(getattr(dummy, method_name))


def test_package_init_reexports_contract() -> None:
    from excel_mcp.routing import (  # noqa: E402
        ROUTED_WORKBOOK_OPERATION_NAMES as names,
        RoutedWorkbookOperations as RWO,
        WorkbookOperationMetadata as WOM,
    )

    assert RWO is RoutedWorkbookOperations
    assert names is ROUTED_WORKBOOK_OPERATION_NAMES
    assert WOM is WorkbookOperationMetadata
