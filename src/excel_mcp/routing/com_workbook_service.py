"""COM-backed workbook façade implementing ``RoutedWorkbookOperations`` (Epic 6).

COM and pywin32 are used **only** inside callables passed to
:class:`excel_mcp.com_executor.ComThreadExecutor` (FR-10: no Excel start on
import or idle paths).
"""

from __future__ import annotations

import os
from typing import Any, Dict, List, Mapping, Optional

from excel_mcp.com_executor import ComThreadExecutor

_COM_NOT_IMPLEMENTED = "Error: COM path not implemented for this operation yet"


def _norm_workbook_path(path: str) -> str:
    return os.path.normcase(os.path.normpath(os.path.abspath(path)))


class ComWorkbookService:
    """Thin COM implementation of ``RoutedWorkbookOperations``; vertical slice: ``write_cell_grid``."""

    def __init__(self, executor: ComThreadExecutor) -> None:
        self._executor = executor

    def _stub(self) -> str:
        return _COM_NOT_IMPLEMENTED

    def read_range_with_metadata(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        preview_only: bool = False,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_cell, end_cell, preview_only, operation_metadata
        return self._stub()

    def workbook_metadata(
        self,
        filepath: str,
        include_ranges: bool = False,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, include_ranges, operation_metadata
        return self._stub()

    def read_merged_cell_ranges(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, operation_metadata
        return self._stub()

    def read_worksheet_data_validation(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, operation_metadata
        return self._stub()

    def validate_sheet_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_cell, end_cell, operation_metadata
        return self._stub()

    def validate_formula_syntax(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, cell, formula, operation_metadata
        return self._stub()

    def apply_formula(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, cell, formula, operation_metadata
        return self._stub()

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
        del (
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
            operation_metadata,
        )
        return self._stub()

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
        import win32com.client  # lazy: worker thread only

        target = _norm_workbook_path(filepath)
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return "Error: No running Excel application found"

        wb_com = None
        try:
            count = int(xl.Workbooks.Count)
        except Exception:
            count = 0
        for i in range(1, count + 1):
            try:
                wb = xl.Workbooks.Item(i)
                full = str(wb.FullName)
            except Exception:
                continue
            if _norm_workbook_path(full) == target:
                wb_com = wb
                break
        if wb_com is None:
            return "Error: Workbook not open in Excel (path does not match any open workbook FullName)"

        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        if not data:
            return f"Data written to {sheet_name}"

        nrows = len(data)
        ncols = max(len(row) for row in data)
        grid: list[list[Any]] = []
        for row in data:
            padded = list(row) + [None] * (ncols - len(row))
            grid.append(padded)

        try:
            start = ws.Range(start_cell)
            rng = start.Resize(nrows, ncols)
            if nrows == 1 and ncols == 1:
                rng.Value = grid[0][0]
            else:
                rng.Value = tuple(tuple(r) for r in grid)
        except Exception as exc:
            return f"Error: {exc}"

        return f"Data written to {sheet_name}"

    def create_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, operation_metadata
        return self._stub()

    def create_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, operation_metadata
        return self._stub()

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
        del filepath, sheet_name, data_range, table_name, table_style, operation_metadata
        return self._stub()

    def copy_worksheet(
        self,
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, source_sheet, target_sheet, operation_metadata
        return self._stub()

    def delete_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, operation_metadata
        return self._stub()

    def rename_worksheet(
        self,
        filepath: str,
        old_name: str,
        new_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, old_name, new_name, operation_metadata
        return self._stub()

    def merge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_cell, end_cell, operation_metadata
        return self._stub()

    def unmerge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_cell, end_cell, operation_metadata
        return self._stub()

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
        del filepath, sheet_name, source_start, source_end, target_start, target_sheet, operation_metadata
        return self._stub()

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
        del filepath, sheet_name, start_cell, end_cell, shift_direction, operation_metadata
        return self._stub()

    def insert_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_row, count, operation_metadata
        return self._stub()

    def insert_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_col, count, operation_metadata
        return self._stub()

    def delete_sheet_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_row, count, operation_metadata
        return self._stub()

    def delete_sheet_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        del filepath, sheet_name, start_col, count, operation_metadata
        return self._stub()

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
        import win32com.client  # lazy: worker thread only

        target = _norm_workbook_path(filepath)
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return "Error: No running Excel application found"

        wb_com = None
        try:
            count = int(xl.Workbooks.Count)
        except Exception:
            count = 0
        for i in range(1, count + 1):
            try:
                wb = xl.Workbooks.Item(i)
                full = str(wb.FullName)
            except Exception:
                continue
            if _norm_workbook_path(full) == target:
                wb_com = wb
                break
        if wb_com is None:
            return "Error: Workbook not open in Excel (path does not match any open workbook FullName)"
        try:
            wb_com.Save()
        except Exception as exc:
            return f"Error: {exc}"
        return f"Workbook saved: {filepath}"


__all__ = ["ComWorkbookService"]
