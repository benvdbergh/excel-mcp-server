"""COM worksheet and range structure mutations (copy/delete/rename/merge/insert)."""

from __future__ import annotations

from typing import Optional

from excel_mcp.com.session import get_open_workbook_com

# Excel COM constants (avoid importing win32com at module load)
_XL_SHIFT_UP = -4162
_XL_SHIFT_TO_LEFT = -4159

def copy_worksheet_com(filepath: str, source_sheet: str, target_sheet: str) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def delete_worksheet_com(filepath: str, sheet_name: str) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def rename_worksheet_com(filepath: str, old_name: str, new_name: str) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
    if err:
        return err

    try:
        wb_com.Worksheets(old_name).Name = new_name
    except Exception as exc:
        return f"Error: {exc}"

    return f"Sheet renamed from '{old_name}' to '{new_name}'"

def merge_cells_com(
    filepath: str, sheet_name: str, start_cell: str, end_cell: str
) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def unmerge_cells_com(
    filepath: str, sheet_name: str, start_cell: str, end_cell: str
) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def copy_cell_range_com(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str,
) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def delete_cell_range_com(
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

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def insert_rows_com(
    filepath: str, sheet_name: str, start_row: int, count: int
) -> str:
    if start_row < 1:
        return "Error: Start row must be 1 or greater"
    if count < 1:
        return "Error: Count must be 1 or greater"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def insert_columns_com(
    filepath: str, sheet_name: str, start_col: int, count: int
) -> str:
    if start_col < 1:
        return "Error: Start column must be 1 or greater"
    if count < 1:
        return "Error: Count must be 1 or greater"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def delete_sheet_rows_com(
    filepath: str, sheet_name: str, start_row: int, count: int
) -> str:
    if start_row < 1:
        return "Error: Start row must be 1 or greater"
    if count < 1:
        return "Error: Count must be 1 or greater"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

def delete_sheet_columns_com(
    filepath: str, sheet_name: str, start_col: int, count: int
) -> str:
    if start_col < 1:
        return "Error: Start column must be 1 or greater"
    if count < 1:
        return "Error: Count must be 1 or greater"

    wb_com, err = get_open_workbook_com(filepath, for_write=True)
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

