"""COM-backed workbook façade implementing ``RoutedWorkbookOperations`` (Epic 6).

COM and pywin32 are used **only** inside callables passed to
:class:`excel_mcp.com_executor.ComThreadExecutor` (FR-10: no Excel start on
import or idle paths).
"""

from __future__ import annotations

import numbers
import os
import re
import uuid
from typing import Any, Dict, List, Mapping, Optional, Tuple
from urllib.parse import urljoin

from excel_mcp.com_executor import ComThreadExecutor
from excel_mcp.path_resolution import normalize_workbook_target_for_com

_COM_NOT_IMPLEMENTED = "Error: COM path not implemented for this operation yet"

# Stable routing errors (FR-9 / ADR 0005) — wording fixed for automation clients.
_ERR_COM_MULTIPLE_MATCH = (
    "Error: Multiple Excel workbooks match this path; close duplicates or use a single instance."
)
_ERR_COM_UNSAVED_PATH = (
    "Error: Workbook has no saved path on disk; save the workbook to a known path before COM routing."
)
_ERR_COM_NOT_OPEN = (
    "Error: Workbook not open in Excel (path does not match any open workbook FullName)"
)
_ERR_COM_PROTECTED_VIEW = (
    "Error: Workbook is open in Protected View in Excel; enable editing before COM routing."
)
_ERR_COM_READ_ONLY = (
    "Error: Workbook is read-only in Excel; COM routing cannot modify this workbook."
)

# Excel COM constants (avoid importing win32com at module load)
_XL_SHIFT_UP = -4162
_XL_SHIFT_TO_LEFT = -4159
_XL_LINE_STYLE_CONTINUOUS = 1
_XL_THIN = 2
_XL_MEDIUM = -4138
_XL_THICK = 4
_XL_DOUBLE = -4119
_XL_UNDERLINE_SINGLE = 2
_XL_UNDERLINE_NONE = -4142
_XL_SRC_RANGE = 1
_XL_LIST_HAS_HEADERS_GUESS = 0
_XL_OPEN_XML_WORKBOOK = 51
_H_ALIGN = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
_BORDER_WEIGHT = {
    "thin": _XL_THIN,
    "medium": _XL_MEDIUM,
    "thick": _XL_THICK,
    "double": _XL_DOUBLE,
}


def _norm_workbook_path(path: str) -> str:
    """Canonical disk path for comparison with Excel ``Workbook.FullName`` (FR-1).

    Uses :func:`os.path.realpath` so junctions/symlinks align with the host path
    Excel reports, while staying consistent with ``resolve_target`` for absolute
    paths when the allowlist path is active.
    """
    expanded = os.path.expanduser(path)
    try:
        canonical = os.path.realpath(expanded)
    except OSError:
        canonical = os.path.abspath(expanded)
    return os.path.normcase(os.path.normpath(canonical))


def _coerce_com_count(val: Any, default: int = 0) -> int:
    """Coerce Excel COM ``Count`` to ``int``; reject non-numeric test doubles (e.g. ``MagicMock``)."""
    if isinstance(val, bool):
        return default
    if isinstance(val, numbers.Integral):
        return int(val)
    return default


def _com_bool_is_true(val: Any) -> bool:
    """True only for a real COM ``VARIANT_BOOL`` / Python ``bool`` (mocks are ignored)."""
    return isinstance(val, bool) and val


def _workbook_fullname_norm(wb: Any) -> Optional[str]:
    try:
        return normalize_workbook_target_for_com(str(wb.FullName))
    except Exception:
        return None


def _protected_view_candidate_paths(pv: Any) -> list[str]:
    """Paths to compare to ``target`` for a Protected View window (COM object)."""
    out: list[str] = []
    try:
        wb = pv.Workbook
        fn = _workbook_fullname_norm(wb)
        if fn:
            out.append(fn)
    except Exception:
        pass
    try:
        sp, sn = str(pv.SourcePath), str(pv.SourceName)
        if sp and sn:
            sps = sp.strip()
            if sps.lower().startswith("https://"):
                base = sp if sp.endswith("/") else sp + "/"
                combined = urljoin(
                    base, str(sn).replace("\\", "/").lstrip("/")
                )
                out.append(normalize_workbook_target_for_com(combined))
            else:
                out.append(
                    normalize_workbook_target_for_com(os.path.join(sp, sn))
                )
        elif sp:
            out.append(normalize_workbook_target_for_com(sp))
    except Exception:
        pass
    return out


def _hex_to_bgr_int(color: str) -> int:
    h = color.strip().lstrip("#").upper()
    if h.startswith("FF") and len(h) == 8:
        h = h[2:]
    if len(h) != 6 or not re.fullmatch(r"[0-9A-F]{6}", h):
        raise ValueError(f"Invalid color: {color}")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return int(r | (g << 8) | (b << 16))


class ComWorkbookService:
    """COM implementation of ``RoutedWorkbookOperations`` for routed write tools (Epic 7)."""

    def __init__(self, executor: ComThreadExecutor) -> None:
        self._executor = executor

    def _stub(self) -> str:
        return _COM_NOT_IMPLEMENTED

    @staticmethod
    def _workbook_in_protected_view(xl: Any, wb: Any) -> bool:
        """True if ``wb`` is the workbook shown in a :class:`ProtectedViewWindow` (Excel COM)."""
        want = _workbook_fullname_norm(wb)
        try:
            pvw = xl.ProtectedViewWindows
            n = _coerce_com_count(getattr(pvw, "Count", 0))
        except Exception:
            return False
        for i in range(1, n + 1):
            try:
                pv = pvw.Item(i)
                pw = pv.Workbook
                got = _workbook_fullname_norm(pw)
                if want and got and want == got:
                    return True
            except Exception:
                continue
        return False

    @staticmethod
    def _collect_workbooks_matching_path(xl: Any, target: str) -> List[Any]:
        """All open COM workbooks whose on-disk path equals ``target`` (normalized).

        Includes workbooks open only in **Protected View** — those are not members of
        ``Application.Workbooks``; match via ``Application.ProtectedViewWindows`` and
        ``ProtectedViewWindow.Workbook`` / ``SourcePath`` + ``SourceName`` (Excel COM).
        """
        matches: List[Any] = []
        try:
            count = _coerce_com_count(getattr(xl.Workbooks, "Count", 0))
        except Exception:
            count = 0
        for i in range(1, count + 1):
            try:
                wb = xl.Workbooks.Item(i)
                if _workbook_fullname_norm(wb) == target:
                    matches.append(wb)
            except Exception:
                continue

        try:
            pvw = xl.ProtectedViewWindows
            n_pv = _coerce_com_count(getattr(pvw, "Count", 0))
        except Exception:
            n_pv = 0
        for i in range(1, n_pv + 1):
            try:
                pv = pvw.Item(i)
                for cand in _protected_view_candidate_paths(pv):
                    if cand == target:
                        matches.append(pv.Workbook)
                        break
            except Exception:
                continue

        # Same COM instance must not be counted twice if ever visible from both paths.
        seen: set[int] = set()
        unique: List[Any] = []
        for wb in matches:
            k = id(wb)
            if k not in seen:
                seen.add(k)
                unique.append(wb)
        return unique

    @staticmethod
    def _get_open_workbook_com(filepath: str) -> Tuple[Any, Optional[str]]:
        """Return ``(workbook_com, None)`` or ``(None, error_message)``.

        Never-saved workbooks (e.g. ``Book1``) expose a **non-disk** ``FullName`` and an
        empty ``Workbook.Path``; they cannot be matched to a caller-supplied absolute
        file path. When exactly one such workbook is open and nothing matches ``target``,
        return :data:`_ERR_COM_UNSAVED_PATH` instead of a generic not-open message.
        """
        import win32com.client  # lazy: worker thread only

        target = normalize_workbook_target_for_com(filepath)
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return None, "Error: No running Excel application found"

        matches = ComWorkbookService._collect_workbooks_matching_path(xl, target)
        if len(matches) > 1:
            return None, _ERR_COM_MULTIPLE_MATCH

        if len(matches) == 0:
            # Fail-closed: only infer "unsaved path" when a single workbook is open and
            # it has no saved directory (``Path`` is empty per Excel COM).
            try:
                wb_count = _coerce_com_count(getattr(xl.Workbooks, "Count", 0))
            except Exception:
                wb_count = 0
            never_saved_open = 0
            for i in range(1, wb_count + 1):
                try:
                    wb = xl.Workbooks.Item(i)
                    p = str(getattr(wb, "Path", "") or "").strip()
                    if p == "":
                        never_saved_open += 1
                except Exception:
                    continue
            if wb_count == 1 and never_saved_open == 1:
                return None, _ERR_COM_UNSAVED_PATH
            return None, _ERR_COM_NOT_OPEN

        wb_com = matches[0]
        if ComWorkbookService._workbook_in_protected_view(xl, wb_com):
            return None, _ERR_COM_PROTECTED_VIEW
        try:
            if _com_bool_is_true(getattr(wb_com, "ReadOnly", False)):
                return None, _ERR_COM_READ_ONLY
        except Exception:
            pass
        return wb_com, None

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
        del operation_metadata
        return self._executor.submit(
            self._apply_formula_com, filepath, sheet_name, cell, formula
        )

    @staticmethod
    def _apply_formula_com(
        filepath: str, sheet_name: str, cell: str, formula: str
    ) -> str:
        from excel_mcp.cell_utils import validate_cell_reference
        from excel_mcp.validation import validate_formula

        if not validate_cell_reference(cell):
            return f"Error: Invalid cell reference: {cell}"

        ftext = formula if formula.startswith("=") else f"={formula}"
        is_valid, vmsg = validate_formula(ftext)
        if not is_valid:
            return f"Error: Invalid formula syntax: {vmsg}"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Range(cell).Formula = ftext
        except Exception as exc:
            return f"Error: {exc}"

        return f"Applied formula '{ftext}' to cell {cell}"

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
        if conditional_format is not None:
            return (
                "Error: conditional_format is not supported on the COM path "
                "(use file transport or omit conditional_format)"
            )

        from excel_mcp.cell_utils import parse_cell_range, validate_cell_reference

        if not validate_cell_reference(start_cell):
            return f"Error: Invalid start cell reference: {start_cell}"
        if end_cell and not validate_cell_reference(end_cell):
            return f"Error: Invalid end cell reference: {end_cell}"
        try:
            parse_cell_range(start_cell, end_cell)
        except ValueError as e:
            return f"Error: Invalid cell range: {e}"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        if end_cell:
            rng = ws.Range(start_cell, end_cell)
        else:
            rng = ws.Range(start_cell)

        try:
            fnt = rng.Font
            fnt.Bold = bold
            fnt.Italic = italic
            fnt.Underline = _XL_UNDERLINE_SINGLE if underline else _XL_UNDERLINE_NONE
            if font_size is not None:
                fnt.Size = font_size
            if font_color is not None:
                fnt.Color = _hex_to_bgr_int(font_color)
            if bg_color is not None:
                rng.Interior.Color = _hex_to_bgr_int(bg_color)
            if number_format is not None:
                rng.NumberFormat = number_format

            if alignment is not None or wrap_text:
                hkey = (alignment or "general").lower()
                hconst = _H_ALIGN.get(hkey)
                if hconst is not None:
                    rng.HorizontalAlignment = hconst
                rng.WrapText = wrap_text

            if border_style is not None:
                key = border_style.lower()
                if key not in _BORDER_WEIGHT and key != "double":
                    return f"Error: Unsupported border_style for COM: {border_style}"
                bcol = border_color or "000000"
                clr = _hex_to_bgr_int(bcol)
                if key == "double":
                    line_style = _XL_DOUBLE
                    weight = _XL_THICK
                else:
                    line_style = _XL_LINE_STYLE_CONTINUOUS
                    weight = _BORDER_WEIGHT[key]
                for edge in (7, 8, 9, 10, 11, 12):  # xlEdgeLeft..DiagonalDown
                    b = rng.Borders(edge)
                    b.LineStyle = line_style
                    b.Weight = weight
                    b.Color = clr

            if protection is not None:
                if "locked" in protection:
                    rng.Locked = bool(protection["locked"])
                if "hidden" in protection:
                    rng.FormulaHidden = bool(protection["hidden"])

            if merge_cells:
                if not end_cell:
                    return "Error: merge_cells requires end_cell"
                rng.Merge()
        except ValueError as exc:
            return f"Error: {exc}"
        except Exception as exc:
            return f"Error: {exc}"

        return "Range formatted successfully"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err

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
        del operation_metadata
        return self._executor.submit(self._create_workbook_com, filepath)

    @staticmethod
    def _create_workbook_com(filepath: str) -> str:
        import win32com.client  # lazy: worker thread only

        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return "Error: No running Excel application found"

        path = os.path.abspath(filepath)
        parent = os.path.dirname(path)
        if parent:
            try:
                os.makedirs(parent, exist_ok=True)
            except OSError as exc:
                return f"Error: {exc}"

        try:
            wb = xl.Workbooks.Add()
            if path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
                wb.SaveAs(path, FileFormat=_XL_OPEN_XML_WORKBOOK)
            else:
                wb.SaveAs(path)
        except Exception as exc:
            return f"Error: {exc}"

        return f"Created workbook at {filepath}"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err

        try:
            wb_com.Worksheets(sheet_name)
            return f"Error: Sheet {sheet_name} already exists"
        except Exception:
            pass

        try:
            ws = wb_com.Worksheets.Add()
            ws.Name = sheet_name
        except Exception as exc:
            return f"Error: {exc}"

        return f"Sheet {sheet_name} created successfully"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        tname = table_name or f"Table_{uuid.uuid4().hex[:8]}"

        try:
            src = ws.Range(data_range)
            lo = ws.ListObjects.Add(
                _XL_SRC_RANGE,
                src,
                None,
                _XL_LIST_HAS_HEADERS_GUESS,
            )
            lo.Name = tname
            lo.TableStyle = table_style
        except Exception as exc:
            return f"Error: {exc}"

        return f"Successfully created table '{tname}' in sheet '{sheet_name}'."

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err

        app = wb_com.Application
        prev = app.DisplayAlerts
        app.DisplayAlerts = False
        try:
            src = wb_com.Worksheets(source_sheet)
            n = int(wb_com.Worksheets.Count)
            last = wb_com.Worksheets(n)
            src.Copy(After=last)
            wb_com.ActiveSheet.Name = target_sheet
        except Exception as exc:
            return f"Error: {exc}"
        finally:
            app.DisplayAlerts = prev

        return f"Sheet '{source_sheet}' copied to '{target_sheet}'"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err

        app = wb_com.Application
        prev = app.DisplayAlerts
        app.DisplayAlerts = False
        try:
            wb_com.Worksheets(sheet_name).Delete()
        except Exception as exc:
            return f"Error: {exc}"
        finally:
            app.DisplayAlerts = prev

        return f"Sheet '{sheet_name}' deleted"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err

        try:
            wb_com.Worksheets(old_name).Name = new_name
        except Exception as exc:
            return f"Error: {exc}"

        return f"Sheet renamed from '{old_name}' to '{new_name}'"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        range_string = f"{start_cell}:{end_cell}"
        try:
            ws.Range(start_cell, end_cell).Merge()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Range '{range_string}' merged in sheet '{sheet_name}'"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Range(start_cell, end_cell).UnMerge()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Range '{start_cell}:{end_cell}' unmerged successfully"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            src_ws = wb_com.Worksheets(sheet_name)
            dst_ws = wb_com.Worksheets(target_sheet)
        except Exception:
            return f"Error: Sheet not found"

        try:
            src_ws.Range(source_start, source_end).Copy(Destination=dst_ws.Range(target_start))
        except Exception as exc:
            return f"Error: {exc}"

        return "Range copied successfully"

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
        sd = shift_direction.lower()
        if sd not in ("up", "left"):
            return (
                f"Error: Invalid shift direction: {shift_direction}. Must be 'up' or 'left'"
            )

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        shift = _XL_SHIFT_UP if sd == "up" else _XL_SHIFT_TO_LEFT
        range_string = f"{start_cell}:{end_cell}"
        try:
            ws.Range(start_cell, end_cell).Delete(Shift=shift)
        except Exception as exc:
            return f"Error: {exc}"

        return f"Range {range_string} deleted successfully"

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
        if start_row < 1:
            return "Error: Start row must be 1 or greater"
        if count < 1:
            return "Error: Count must be 1 or greater"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Rows(start_row).Resize(count).EntireRow.Insert()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Inserted {count} row(s) starting at row {start_row} in sheet '{sheet_name}'"

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
        if start_col < 1:
            return "Error: Start column must be 1 or greater"
        if count < 1:
            return "Error: Count must be 1 or greater"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Columns(start_col).Resize(count).EntireColumn.Insert()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Inserted {count} column(s) starting at column {start_col} in sheet '{sheet_name}'"

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
        if start_row < 1:
            return "Error: Start row must be 1 or greater"
        if count < 1:
            return "Error: Count must be 1 or greater"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Rows(start_row).Resize(count).EntireRow.Delete()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Deleted {count} row(s) starting at row {start_row} in sheet '{sheet_name}'"

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
        if start_col < 1:
            return "Error: Start column must be 1 or greater"
        if count < 1:
            return "Error: Count must be 1 or greater"

        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            ws = wb_com.Worksheets(sheet_name)
        except Exception:
            return f"Error: Sheet '{sheet_name}' not found"

        try:
            ws.Columns(start_col).Resize(count).EntireColumn.Delete()
        except Exception as exc:
            return f"Error: {exc}"

        return f"Deleted {count} column(s) starting at column {start_col} in sheet '{sheet_name}'"

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
        wb_com, err = ComWorkbookService._get_open_workbook_com(filepath)
        if err:
            return err
        try:
            wb_com.Save()
        except Exception as exc:
            return f"Error: {exc}"
        return f"Workbook saved: {filepath}"


__all__ = ["ComWorkbookService"]
