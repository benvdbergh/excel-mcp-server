"""COM-backed workbook façade implementing ``RoutedWorkbookOperations``.

COM and pywin32 are used **only** inside callables passed to
:class:`excel_mcp.com.executor.ComThreadExecutor` (FR-10: no Excel start on
import or idle paths).
"""

from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

from excel_mcp.com import listobject as listobject_mod
from excel_mcp.com import read as read_mod
from excel_mcp.com import session as session_mod
from excel_mcp.com import sheet as sheet_mod
from excel_mcp.com import views as views_mod
from excel_mcp.com import write as write_mod
from excel_mcp.com.executor import ComThreadExecutor
from excel_mcp.query import DEFAULT_EXPORT_MAX_ROWS, DEFAULT_LIST_TABLES_DETAIL
from excel_mcp.value_mode import validate_metadata_mode, validate_value_mode

_COM_NOT_IMPLEMENTED = "Error: COM path not implemented for this operation yet"

class ComWorkbookService:
    """COM implementation of ``RoutedWorkbookOperations`` for routed write tools (Epic 7)."""

    def __init__(self, executor: ComThreadExecutor) -> None:
        self._executor = executor

    def _stub(self) -> str:
        return _COM_NOT_IMPLEMENTED

    @staticmethod
    def _collect_workbooks_matching_path(xl: Any, target: str) -> List[Any]:
        return session_mod.collect_workbooks_matching_path(xl, target)

    @staticmethod
    def _get_open_workbook_com(filepath: str, *, for_write: bool = False) -> Tuple[Any, Optional[str]]:
        return session_mod.get_open_workbook_com(filepath, for_write=for_write)

    def read_range_with_metadata(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        *,
        value_mode: str = "value",
        metadata_mode: str = "full",
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        try:
            validate_value_mode(value_mode)
            validate_metadata_mode(metadata_mode)
        except ValueError as e:
            return f"Error: {e}"
        return self._executor.submit(
            self._read_range_with_metadata_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
            value_mode,
            metadata_mode,
        )

    @staticmethod
    def _read_range_with_metadata_com(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str],
        value_mode: str = "value",
        metadata_mode: str = "full",
    ) -> str:
        return read_mod.read_range_with_metadata_com(filepath, sheet_name, start_cell, end_cell, value_mode, metadata_mode)

    def export_worksheet_table(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        max_rows: int = DEFAULT_EXPORT_MAX_ROWS,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._export_worksheet_table_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
            max_rows,
        )

    @staticmethod
    def _export_worksheet_table_com(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str],
        max_rows: int,
    ) -> str:
        return read_mod.export_worksheet_table_com(filepath, sheet_name, start_cell, end_cell, max_rows)

    def workbook_metadata(
        self,
        filepath: str,
        include_ranges: bool = False,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._workbook_metadata_com, filepath, include_ranges
        )

    @staticmethod
    def _workbook_metadata_com(filepath: str, include_ranges: bool) -> str:
        return read_mod.workbook_metadata_com(filepath, include_ranges)

    def list_tables(
        self,
        filepath: str,
        detail: str = DEFAULT_LIST_TABLES_DETAIL,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._list_tables_com, filepath, detail)

    @staticmethod
    def _list_tables_com(filepath: str, detail: str) -> str:
        return listobject_mod.list_tables_com(filepath, detail)

    def query_table(
        self,
        filepath: str,
        table: str,
        columns: Optional[List[str]] = None,
        where: Optional[List[Dict[str, Any]]] = None,
        limit: Optional[int] = None,
        offset: Optional[int] = None,
        omit_empty: bool = False,
        search: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._query_table_com,
            filepath,
            table,
            columns,
            where,
            limit,
            offset,
            omit_empty,
            search,
        )

    @staticmethod
    def _query_table_com(
        filepath: str,
        table: str,
        columns: Optional[List[str]],
        where: Optional[List[Dict[str, Any]]],
        limit: Optional[int],
        offset: Optional[int],
        omit_empty: bool = False,
        search: Optional[str] = None,
    ) -> str:
        return listobject_mod.query_table_com(
            filepath,
            table,
            columns,
            where,
            limit,
            offset,
            omit_empty=omit_empty,
            search=search,
        )

    def map_sheet_layout(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._map_sheet_layout_com, filepath, sheet_name
        )

    @staticmethod
    def _map_sheet_layout_com(filepath: str, sheet_name: str) -> str:
        return listobject_mod.map_sheet_layout_com(filepath, sheet_name)

    def query_region(
        self,
        filepath: str,
        sheet_name: str,
        region_id: Optional[str] = None,
        range: Optional[str] = None,
        columns: Optional[List[str]] = None,
        where: Optional[List[Dict[str, Any]]] = None,
        limit: Optional[int] = None,
        offset: Optional[int] = None,
        omit_empty: bool = False,
        search: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._query_region_com,
            filepath,
            sheet_name,
            region_id,
            range,
            columns,
            where,
            limit,
            offset,
            omit_empty,
            search,
        )

    @staticmethod
    def _query_region_com(
        filepath: str,
        sheet_name: str,
        region_id: Optional[str],
        range: Optional[str],
        columns: Optional[List[str]],
        where: Optional[List[Dict[str, Any]]],
        limit: Optional[int],
        offset: Optional[int],
        omit_empty: bool = False,
        search: Optional[str] = None,
    ) -> str:
        return listobject_mod.query_region_com(
            filepath,
            sheet_name,
            region_id,
            range,
            columns,
            where,
            limit,
            offset,
            omit_empty=omit_empty,
            search=search,
        )

    def read_merged_cell_ranges(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._read_merged_cell_ranges_com, filepath, sheet_name
        )

    @staticmethod
    def _read_merged_cell_ranges_com(filepath: str, sheet_name: str) -> str:
        return read_mod.read_merged_cell_ranges_com(filepath, sheet_name)

    def read_worksheet_data_validation(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._read_worksheet_data_validation_com, filepath, sheet_name
        )

    @staticmethod
    def _read_worksheet_data_validation_com(
        filepath: str, sheet_name: str
    ) -> str:
        return read_mod.read_worksheet_data_validation_com(filepath, sheet_name)

    def validate_sheet_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._validate_sheet_range_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
        )

    @staticmethod
    def _validate_sheet_range_com(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str],
    ) -> str:
        return read_mod.validate_sheet_range_com(filepath, sheet_name, start_cell, end_cell)

    def validate_formula_syntax(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._validate_formula_syntax_com, filepath, sheet_name, cell, formula
        )

    @staticmethod
    def _validate_formula_syntax_com(
        filepath: str, sheet_name: str, cell: str, formula: str
    ) -> str:
        return read_mod.validate_formula_syntax_com(filepath, sheet_name, cell, formula)

    def open_workbook_in_excel(self, filepath: str) -> str:
        """``Workbooks.Open`` on the COM thread (ADR 0008 lifecycle)."""
        return self._executor.submit(self._open_workbook_in_excel_com, filepath)

    def close_workbook_in_excel(self, filepath: str, *, save: bool = False) -> str:
        """Close a workbook in the Excel host; optional save (ADR 0008 lifecycle)."""
        return self._executor.submit(
            self._close_workbook_in_excel_com, filepath, save
        )

    @staticmethod
    def _open_workbook_in_excel_com(filepath: str) -> str:
        return session_mod.open_workbook_in_excel_com(filepath)

    @staticmethod
    def _close_workbook_in_excel_com(filepath: str, save: bool) -> str:
        return session_mod.close_workbook_in_excel_com(filepath, save)

    def evaluate_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: Optional[str] = None,
        end_cell: Optional[str] = None,
    ) -> str:
        """Force Excel to recalculate a worksheet or range (COM-only side effect)."""
        return self._executor.submit(
            self._evaluate_range_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
        )

    @staticmethod
    def _evaluate_range_com(
        filepath: str,
        sheet_name: str,
        start_cell: Optional[str],
        end_cell: Optional[str],
    ) -> str:
        return read_mod.evaluate_range_com(filepath, sheet_name, start_cell, end_cell)

    def list_open_workbooks(self, detail: Optional[str] = None) -> str:
        """Enumerate ``Application.Workbooks`` on the COM thread (ADR 0009).

        ``detail`` is ``minimal`` (default) or ``active_context`` (adds active
        workbook, sheet, and selection). Returns a JSON string.
        """
        return self._executor.submit(
            ComWorkbookService._list_open_workbooks_com, detail
        )

    @staticmethod
    def _list_open_workbooks_com(detail: Optional[str] = None) -> str:
        return session_mod.list_open_workbooks_com(detail)

    def apply_formula(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._apply_formula_com, filepath, sheet_name, cell, formula
        )

    @staticmethod
    def _apply_formula_com(
        filepath: str, sheet_name: str, cell: str, formula: str
    ) -> str:
        return write_mod.apply_formula_com(filepath, sheet_name, cell, formula)

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
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._format_range_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
            bold,
            italic,
            underline,
            font_size,
            font_color,
            bg_color,
            border_style,
            border_color,
            number_format,
            alignment,
            wrap_text,
            merge_cells,
            protection,
            conditional_format,
        )

    @staticmethod
    def _format_range_com(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str],
        bold: bool,
        italic: bool,
        underline: bool,
        font_size: Optional[int],
        font_color: Optional[str],
        bg_color: Optional[str],
        border_style: Optional[str],
        border_color: Optional[str],
        number_format: Optional[str],
        alignment: Optional[str],
        wrap_text: bool,
        merge_cells: bool,
        protection: Optional[Dict[str, Any]],
        conditional_format: Optional[Dict[str, Any]],
    ) -> str:
        return write_mod.format_range_com(filepath, sheet_name, start_cell, end_cell, bold, italic, underline, font_size, font_color, bg_color, border_style, border_color, number_format, alignment, wrap_text, merge_cells, protection, conditional_format)

    def write_cell_grid(
        self,
        filepath: str,
        sheet_name: str,
        data: List[List[Any]],
        start_cell: str = "A1",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._write_cell_grid_com, filepath, sheet_name, data, start_cell)

    @staticmethod
    def _write_cell_grid_com(
        filepath: str,
        sheet_name: str,
        data: List[List[Any]],
        start_cell: str,
    ) -> str:
        return write_mod.write_cell_grid_com(filepath, sheet_name, data, start_cell)

    def create_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._create_workbook_com, filepath)

    @staticmethod
    def _create_workbook_com(filepath: str) -> str:
        return write_mod.create_workbook_com(filepath)

    def create_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._create_worksheet_com, filepath, sheet_name)

    @staticmethod
    def _create_worksheet_com(filepath: str, sheet_name: str) -> str:
        return write_mod.create_worksheet_com(filepath, sheet_name)

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
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del (
            filepath,
            sheet_name,
            data_range,
            chart_type,
            target_cell,
            title,
            x_axis,
            y_axis,
            operation_metadata,
        )
        return self._stub()

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
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, data_range, rows, values, columns, agg_func, operation_metadata
        return self._stub()

    def create_excel_table(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        table_name: Optional[str] = None,
        table_style: str = "TableStyleMedium9",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._create_excel_table_com,
            filepath,
            sheet_name,
            data_range,
            table_name,
            table_style,
        )

    @staticmethod
    def _create_excel_table_com(
        filepath: str,
        sheet_name: str,
        data_range: str,
        table_name: Optional[str],
        table_style: str,
    ) -> str:
        return listobject_mod.create_excel_table_com(filepath, sheet_name, data_range, table_name, table_style)

    def apply_table_view(
        self,
        filepath: str,
        view_spec: dict[str, Any],
        mode: str = "in_place",
        view_applicability: Optional[dict[str, Any]] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._apply_table_view_com,
            filepath,
            view_spec,
            mode,
            view_applicability,
        )

    @staticmethod
    def _apply_table_view_com(
        filepath: str,
        view_spec: dict[str, Any],
        mode: str,
        view_applicability: Optional[dict[str, Any]],
    ) -> str:
        return views_mod.apply_table_view_com(filepath, view_spec, mode, view_applicability)

    @staticmethod
    def _apply_snapshot_view_com(
        wb_com: Any,
        view_spec: dict[str, Any],
        view_applicability: Optional[dict[str, Any]],
    ) -> str:
        return views_mod.apply_snapshot_view_com(wb_com, view_spec, view_applicability)

    @staticmethod
    def _apply_listobject_view_com(
        wb_com: Any,
        view_spec: dict[str, Any],
        mode: str,
        view_applicability: Optional[dict[str, Any]],
    ) -> str:
        return views_mod.apply_listobject_view_com(wb_com, view_spec, mode, view_applicability)

    @staticmethod
    def _apply_range_view_com(
        wb_com: Any,
        view_spec: dict[str, Any],
        mode: str,
        view_applicability: Optional[dict[str, Any]],
    ) -> str:
        return views_mod.apply_range_view_com(wb_com, view_spec, mode, view_applicability)

    def clear_table_view(
        self,
        filepath: str,
        restore_token: dict[str, Any],
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._clear_table_view_com, filepath, restore_token
        )

    @staticmethod
    def _clear_table_view_com(
        filepath: str, restore_token: dict[str, Any]
    ) -> str:
        return views_mod.clear_table_view_com(filepath, restore_token)

    @staticmethod
    def _clear_snapshot_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
        return views_mod.clear_snapshot_view_com(wb_com, token)

    @staticmethod
    def _clear_listobject_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
        return views_mod.clear_listobject_view_com(wb_com, token)

    @staticmethod
    def _clear_range_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
        return views_mod.clear_range_view_com(wb_com, token)

    def copy_worksheet(
        self,
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._copy_worksheet_com, filepath, source_sheet, target_sheet)

    @staticmethod
    def _copy_worksheet_com(filepath: str, source_sheet: str, target_sheet: str) -> str:
        return sheet_mod.copy_worksheet_com(filepath, source_sheet, target_sheet)

    def delete_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._delete_worksheet_com, filepath, sheet_name)

    @staticmethod
    def _delete_worksheet_com(filepath: str, sheet_name: str) -> str:
        return sheet_mod.delete_worksheet_com(filepath, sheet_name)

    def rename_worksheet(
        self,
        filepath: str,
        old_name: str,
        new_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._rename_worksheet_com, filepath, old_name, new_name)

    @staticmethod
    def _rename_worksheet_com(filepath: str, old_name: str, new_name: str) -> str:
        return sheet_mod.rename_worksheet_com(filepath, old_name, new_name)

    def merge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._merge_cells_com, filepath, sheet_name, start_cell, end_cell
        )

    @staticmethod
    def _merge_cells_com(
        filepath: str, sheet_name: str, start_cell: str, end_cell: str
    ) -> str:
        return sheet_mod.merge_cells_com(filepath, sheet_name, start_cell, end_cell)

    def unmerge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._unmerge_cells_com, filepath, sheet_name, start_cell, end_cell
        )

    @staticmethod
    def _unmerge_cells_com(
        filepath: str, sheet_name: str, start_cell: str, end_cell: str
    ) -> str:
        return sheet_mod.unmerge_cells_com(filepath, sheet_name, start_cell, end_cell)

    def copy_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._copy_cell_range_com,
            filepath,
            sheet_name,
            source_start,
            source_end,
            target_start,
            target_sheet or sheet_name,
        )

    @staticmethod
    def _copy_cell_range_com(
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: str,
    ) -> str:
        return sheet_mod.copy_cell_range_com(filepath, sheet_name, source_start, source_end, target_start, target_sheet)

    def delete_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        shift_direction: str = "up",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._delete_cell_range_com,
            filepath,
            sheet_name,
            start_cell,
            end_cell,
            shift_direction,
        )

    @staticmethod
    def _delete_cell_range_com(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        shift_direction: str,
    ) -> str:
        return sheet_mod.delete_cell_range_com(filepath, sheet_name, start_cell, end_cell, shift_direction)

    def insert_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._insert_rows_com, filepath, sheet_name, start_row, count
        )

    @staticmethod
    def _insert_rows_com(
        filepath: str, sheet_name: str, start_row: int, count: int
    ) -> str:
        return sheet_mod.insert_rows_com(filepath, sheet_name, start_row, count)

    def insert_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._insert_columns_com, filepath, sheet_name, start_col, count
        )

    @staticmethod
    def _insert_columns_com(
        filepath: str, sheet_name: str, start_col: int, count: int
    ) -> str:
        return sheet_mod.insert_columns_com(filepath, sheet_name, start_col, count)

    def delete_sheet_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._delete_sheet_rows_com, filepath, sheet_name, start_row, count
        )

    @staticmethod
    def _delete_sheet_rows_com(
        filepath: str, sheet_name: str, start_row: int, count: int
    ) -> str:
        return sheet_mod.delete_sheet_rows_com(filepath, sheet_name, start_row, count)

    def delete_sheet_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(
            self._delete_sheet_columns_com, filepath, sheet_name, start_col, count
        )

    @staticmethod
    def _delete_sheet_columns_com(
        filepath: str, sheet_name: str, start_col: int, count: int
    ) -> str:
        return sheet_mod.delete_sheet_columns_com(filepath, sheet_name, start_col, count)

    def save_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del operation_metadata
        return self._executor.submit(self._save_workbook_com, filepath)

    @staticmethod
    def _save_workbook_com(filepath: str) -> str:
        return write_mod.save_workbook_com(filepath)


__all__ = ["ComWorkbookService"]
