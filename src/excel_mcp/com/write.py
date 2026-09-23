"""COM write helpers: formula, format, grid, create workbook/sheet, save."""

from __future__ import annotations

import os
import re
from typing import Any, List, Mapping, Optional, Sequence

from excel_mcp.com.session import get_open_workbook_com

# Excel COM constants (avoid importing win32com at module load)
_XL_LINE_STYLE_CONTINUOUS = 1
_XL_THIN = 2
_XL_MEDIUM = -4138
_XL_THICK = 4
_XL_DOUBLE = -4119
_XL_UNDERLINE_SINGLE = 2
_XL_UNDERLINE_NONE = -4142
_XL_OPEN_XML_WORKBOOK = 51
_H_ALIGN = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
_BORDER_WEIGHT = {
    "thin": _XL_THIN,
    "medium": _XL_MEDIUM,
    "thick": _XL_THICK,
    "double": _XL_DOUBLE,
}

def _hex_to_bgr_int(color: str) -> int:
    h = color.strip().lstrip("#").upper()
    if h.startswith("FF") and len(h) == 8:
        h = h[2:]
    if len(h) != 6 or not re.fullmatch(r"[0-9A-F]{6}", h):
        raise ValueError(f"Invalid color: {color}")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return int(r | (g << 8) | (b << 16))

def apply_formula_com(
    filepath: str, sheet_name: str, cell: str, formula: str
) -> str:
    from excel_mcp.cells import validate_cell_reference
    from excel_mcp.formula_syntax import validate_formula

    if not validate_cell_reference(cell):
        return f"Error: Invalid cell reference: {cell}"

    ftext = formula if formula.startswith("=") else f"={formula}"
    is_valid, vmsg = validate_formula(ftext)
    if not is_valid:
        return f"Error: Invalid formula syntax: {vmsg}"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def format_range_com(
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

    from excel_mcp.cells import parse_cell_range, validate_cell_reference

    if not validate_cell_reference(start_cell):
        return f"Error: Invalid start cell reference: {start_cell}"
    if end_cell and not validate_cell_reference(end_cell):
        return f"Error: Invalid end cell reference: {end_cell}"
    try:
        parse_cell_range(start_cell, end_cell)
    except ValueError as e:
        return f"Error: Invalid cell range: {e}"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def write_cell_grid_com(
    filepath: str,
    sheet_name: str,
    data: List[List[Any]],
    start_cell: str,
) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def create_workbook_com(filepath: str) -> str:
    try:
        import win32com.client  # lazy: worker thread only
    except ModuleNotFoundError:
        return (
            "Error: COM workbook automation requires Windows with "
            "optional dependency excel-com-mcp[com] (pywin32)."
        )

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

    wb = None
    try:
        wb = xl.Workbooks.Add()
        if path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
            wb.SaveAs(path, FileFormat=_XL_OPEN_XML_WORKBOOK)
        else:
            wb.SaveAs(path)
    except Exception as exc:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        return f"Error: {exc}"

    return f"Created workbook at {filepath}"

def create_worksheet_com(filepath: str, sheet_name: str) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def save_workbook_com(filepath: str) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
    if err:
        return err
    try:
        wb_com.Save()
    except Exception as exc:
        return f"Error: {exc}"
    return f"Workbook saved: {filepath}"

