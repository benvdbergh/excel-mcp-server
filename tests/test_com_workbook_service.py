"""ComWorkbookService: executor wiring and COM workers without Excel (Story 7-1)."""

from __future__ import annotations

import json
import os
import sys
import types
from unittest.mock import MagicMock, patch

import pytest

from excel_mcp.routing.com_workbook_service import ComWorkbookService


class ImmediateExecutor:
    """Runs submitted callables on the calling thread (no background worker)."""

    def submit(self, fn, /, *args, **kwargs):
        return fn(*args, **kwargs)


def _fake_win32_modules(
    xl_app: MagicMock,
    *,
    protected_view_windows: MagicMock | None = None,
) -> dict[str, types.ModuleType]:
    client_mod = types.ModuleType("win32com.client")
    client_mod.GetActiveObject = lambda *_a, **_kw: xl_app
    pkg = types.ModuleType("win32com")
    pkg.client = client_mod
    # Avoid MagicMock Count coercing to a non-zero integer in COM helpers.
    if protected_view_windows is not None:
        xl_app.ProtectedViewWindows = protected_view_windows
    else:
        pv = MagicMock()
        pv.Count = 0
        xl_app.ProtectedViewWindows = pv
    return {"win32com": pkg, "win32com.client": client_mod}


def _workbook_mock(path: str, worksheets: MagicMock | dict) -> MagicMock:
    wb = MagicMock()
    wb.FullName = path
    wb.Application = MagicMock()
    wb.Application.DisplayAlerts = True
    if isinstance(worksheets, dict):

        def _ws(name):
            if name in worksheets:
                return worksheets[name]
            raise RuntimeError(f"Sheet '{name}' not found")

        wb.Worksheets = _ws
    else:
        wb.Worksheets = worksheets
    return wb


@pytest.fixture
def book_path(tmp_path):
    p = tmp_path / "wb.xlsx"
    return str(p.resolve())


def test_write_cell_grid_no_excel_returns_error(book_path):
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.write_cell_grid(book_path, "Sheet1", [[1]], "A1")
    assert msg.startswith("Error:") and "Workbook" in msg


def test_write_cell_grid_success_message(book_path):
    ws = MagicMock()
    start = MagicMock()
    ws.Range = MagicMock(return_value=start)
    rng = MagicMock()
    start.Resize = MagicMock(return_value=rng)

    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.write_cell_grid(book_path, "Sheet1", [[1, 2]], "A1")

    assert "Data written to Sheet1" in msg
    assert rng.Value == ((1, 2),)


def test_read_range_with_metadata_fallback_reads_direct_cells(book_path):
    ws = MagicMock()
    rng = MagicMock()
    rng.Value2 = ((None, None), (None, None))

    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = 2
    used.Columns.Count = 2
    ws.UsedRange = used

    direct_vals = {
        (1, 1): "A11",
        (1, 2): "B11",
        (2, 1): "A12",
        (2, 2): "B12",
    }
    cells: dict[tuple[int, int], MagicMock] = {}
    for coord, val in direct_vals.items():
        cell = MagicMock()
        cell.Value2 = val
        cell.Validation = MagicMock()
        cell.Validation.Type = 0
        cells[coord] = cell
    cells[(1, 1)].Resize = MagicMock(return_value=rng)

    ws.Cells = MagicMock(side_effect=lambda r, c: cells[(r, c)])
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.read_range_with_metadata(book_path, "Sheet1", "A1", "B2")

    data = json.loads(raw)
    values = [c["value"] for c in data["cells"]]
    assert values == ["A11", "B11", "A12", "B12"]


def _read_range_fixture(
    book_path: str,
    *,
    rng: MagicMock,
    cells: dict[tuple[int, int], MagicMock],
    used_rows: int = 1,
    used_cols: int = 1,
    value_mode: str = "value",
    end_cell: str = "A1",
) -> str:
    ws = MagicMock()
    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = used_rows
    used.Columns.Count = used_cols
    ws.UsedRange = used
    cells[(1, 1)].Resize = MagicMock(return_value=rng)
    ws.Cells = MagicMock(side_effect=lambda r, c: cells[(r, c)])
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        return svc.read_range_with_metadata(
            book_path, "Sheet1", "A1", end_cell, value_mode=value_mode
        )


def test_read_range_with_metadata_text_mode_returns_display_text(book_path):
    """Formatted currency cells return Range.Text, not raw Value2."""
    rng = MagicMock()
    rng.Text = "19,900.00 €"
    rng.Value2 = 19900

    cell = MagicMock()
    cell.Text = "19,900.00 €"
    cell.Value2 = 19900
    cell.Validation = MagicMock()
    cell.Validation.Type = 0

    raw = _read_range_fixture(
        book_path,
        rng=rng,
        cells={(1, 1): cell},
        value_mode="text",
    )
    data = json.loads(raw)
    assert data["value_mode"] == "text"
    assert data["cells"][0]["value"] == "19,900.00 €"


def test_read_range_with_metadata_invalid_value_mode_returns_error(book_path):
    with patch.dict(sys.modules, _fake_win32_modules(MagicMock()), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.read_range_with_metadata(
            book_path, "Sheet1", "A1", value_mode="formatted"
        )
    assert msg.startswith("Error:")
    assert "Invalid value_mode" in msg


def test_read_range_with_metadata_text_mode_multi_cell(book_path):
    ws = MagicMock()
    rng = MagicMock()
    rng.Text = "19,900.00 €"

    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = 1
    used.Columns.Count = 2
    ws.UsedRange = used

    cells: dict[tuple[int, int], MagicMock] = {}
    for coord, text, raw in [
        ((1, 1), "19,900.00 €", 19900),
        ((1, 2), "1,250.50 €", 1250.5),
    ]:
        cell = MagicMock()
        cell.Text = text
        cell.Value2 = raw
        cell.Validation = MagicMock()
        cell.Validation.Type = 0
        cells[coord] = cell
    cells[(1, 1)].Resize = MagicMock(return_value=rng)
    ws.Cells = MagicMock(side_effect=lambda r, c: cells[(r, c)])

    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.read_range_with_metadata(
            book_path, "Sheet1", "A1", "B1", value_mode="text"
        )

    data = json.loads(raw)
    values = [c["value"] for c in data["cells"]]
    assert values == ["19,900.00 €", "1,250.50 €"]


def test_read_range_with_metadata_value_mode_uses_value2(book_path):
    """Default value_mode keeps Value2 reads (raw numbers, not display text)."""
    rng = MagicMock()
    rng.Value2 = 19900
    rng.Text = "19,900.00 €"

    cell = MagicMock()
    cell.Value2 = 19900
    cell.Text = "19,900.00 €"
    cell.Validation = MagicMock()
    cell.Validation.Type = 0

    raw = _read_range_fixture(
        book_path,
        rng=rng,
        cells={(1, 1): cell},
    )
    data = json.loads(raw)
    assert data["value_mode"] == "value"
    assert data["cells"][0]["value"] == 19900


def test_read_range_with_metadata_text_mode_bulk_fallback(book_path):
    """Bulk Range.Text sparsity on 1x1 triggers per-cell Text fallback."""
    rng = MagicMock()
    rng.Text = None

    cell = MagicMock()
    cell.Text = "19,900.00 €"
    cell.Value2 = 19900
    cell.Validation = MagicMock()
    cell.Validation.Type = 0

    raw = _read_range_fixture(
        book_path,
        rng=rng,
        cells={(1, 1): cell},
        value_mode="text",
    )
    data = json.loads(raw)
    assert data["cells"][0]["value"] == "19,900.00 €"


def test_save_workbook_invokes_save(book_path):
    wb = _workbook_mock(book_path, {})
    wb.Save = MagicMock()
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.save_workbook(book_path)

    assert "Workbook saved" in msg
    wb.Save.assert_called_once()


def test_apply_formula_sets_formula(book_path):
    cell_rng = MagicMock()
    ws = MagicMock()
    ws.Range = MagicMock(return_value=cell_rng)
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.apply_formula(book_path, "Sheet1", "B2", "1+1")

    assert "Applied formula" in msg and "B2" in msg
    assert cell_rng.Formula == "=1+1"


def test_format_range_rejects_conditional_format(book_path):
    ws = MagicMock()
    wb = _workbook_mock(book_path, {"S": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.format_range(
            book_path,
            "S",
            "A1",
            end_cell="B2",
            conditional_format={"type": "cell_is"},
        )
    assert msg.startswith("Error:") and "conditional_format" in msg


def test_create_workbook_add_and_saveas(book_path):
    wb_new = MagicMock()
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Add = MagicMock(return_value=wb_new)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.create_workbook(book_path)

    assert "Created workbook at" in msg
    wb_new.SaveAs.assert_called_once()


def test_create_worksheet_adds_sheet(book_path):
    ws_new = MagicMock()
    existing = {"Sheet1": MagicMock()}

    def worksheets(name):
        if name in existing:
            return existing[name]
        raise RuntimeError("missing")

    wb = _workbook_mock(book_path, worksheets)
    wb.Worksheets.Add = MagicMock(return_value=ws_new)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.create_worksheet(book_path, "NewTab")

    assert "created successfully" in msg
    assert ws_new.Name == "NewTab"


def test_create_excel_table_list_object(book_path):
    lo = MagicMock()
    ws = MagicMock()
    rng = MagicMock()
    ws.Range = MagicMock(return_value=rng)
    ws.ListObjects = MagicMock()
    ws.ListObjects.Add = MagicMock(return_value=lo)
    wb = _workbook_mock(book_path, {"S": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.create_excel_table(book_path, "S", "A1:D4", table_name="T1")

    assert "Successfully created table 'T1'" in msg
    ws.ListObjects.Add.assert_called_once()
    assert lo.TableStyle == "TableStyleMedium9"


def test_copy_worksheet_copy_and_rename_active(book_path):
    src = MagicMock()
    last = MagicMock()
    active = MagicMock()
    wb = MagicMock()
    wb.FullName = book_path
    wb.Application = MagicMock()
    wb.Application.DisplayAlerts = True
    wb.ActiveSheet = active

    def worksheets_call(*args, **_kwargs):
        if not args:
            return MagicMock()
        x = args[0]
        if x == "Src":
            return src
        if x == 2:
            return last
        raise RuntimeError("unexpected Worksheets index")

    ws_coll = MagicMock(side_effect=worksheets_call)
    ws_coll.Count = 2
    wb.Worksheets = ws_coll

    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.copy_worksheet(book_path, "Src", "Dst")

    assert "copied to" in msg
    src.Copy.assert_called_once_with(After=last)
    assert active.Name == "Dst"


def test_delete_cell_range_shift(book_path):
    rng = MagicMock()
    ws = MagicMock()
    ws.Range = MagicMock(return_value=rng)
    wb = _workbook_mock(book_path, {"S": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.delete_cell_range(book_path, "S", "A1", "B2", shift_direction="left")

    assert "deleted successfully" in msg
    rng.Delete.assert_called_once()


def test_chart_and_pivot_remain_stubbed():
    svc = ComWorkbookService(ImmediateExecutor())
    assert "not implemented" in svc.create_chart_in_sheet(
        "x", "s", "A1:B2", "line", "D1"
    ).lower()
    assert "not implemented" in svc.create_pivot_table_in_sheet(
        "x", "s", "A1:B2", [], []
    ).lower()


def test_com_thread_executor_still_serializes_com_workbook_workers(book_path):
    """Regression: production uses ComThreadExecutor; ensure submit API works."""
    from excel_mcp.com_executor import ComThreadExecutor

    ws = MagicMock()
    start = MagicMock()
    ws.Range = MagicMock(return_value=start)
    rng = MagicMock()
    start.Resize = MagicMock(return_value=rng)

    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    ex = ComThreadExecutor()
    try:
        with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
            svc = ComWorkbookService(ex)
            msg = svc.write_cell_grid(book_path, "Sheet1", [[3]], "A1")
        assert "Data written" in msg
    finally:
        ex.shutdown(wait=True)


def test_get_open_workbook_com_duplicate_paths_fail_closed(book_path):
    wb1 = _workbook_mock(book_path, {})
    wb2 = _workbook_mock(book_path, {})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 2
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb1 if i == 1 else wb2)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        wb, err = ComWorkbookService._get_open_workbook_com(book_path)

    assert wb is None
    assert (
        err
        == "Error: Multiple Excel workbooks match this path; close duplicates or use a single instance."
    )


def test_get_open_workbook_com_read_only(book_path):
    wb = _workbook_mock(book_path, {})
    wb.ReadOnly = True
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        wb_out, err = ComWorkbookService._get_open_workbook_com(book_path)

    assert wb_out is None
    assert err == "Error: Workbook is read-only in Excel; COM routing cannot modify this workbook."


def test_get_open_workbook_com_protected_view_only(book_path):
    wb_pv = MagicMock()
    wb_pv.FullName = book_path
    wb_pv.ReadOnly = False
    pv = MagicMock()
    pv.Workbook = wb_pv
    pv.SourcePath = os.path.dirname(book_path)
    pv.SourceName = os.path.basename(book_path)

    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0
    xl.Workbooks.Item = MagicMock(side_effect=RuntimeError("no workbooks"))
    pvw = MagicMock()
    pvw.Count = 1
    pvw.Item = MagicMock(side_effect=lambda i: pv if i == 1 else None)
    xl.ProtectedViewWindows = pvw

    with patch.dict(
        sys.modules,
        _fake_win32_modules(xl, protected_view_windows=pvw),
        clear=False,
    ):
        wb_out, err = ComWorkbookService._get_open_workbook_com(book_path)

    assert wb_out is None
    assert (
        err
        == "Error: Workbook is open in Protected View in Excel; enable editing before COM routing."
    )


def test_get_open_workbook_com_protected_view_https_source_path() -> None:
    u = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx"
    wb_pv = MagicMock()
    wb_pv.FullName = u
    wb_pv.ReadOnly = False
    pv = MagicMock()
    pv.Workbook = wb_pv
    pv.SourcePath = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/"
    pv.SourceName = "book.xlsx"

    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0
    xl.Workbooks.Item = MagicMock(side_effect=RuntimeError("no workbooks"))
    pvw = MagicMock()
    pvw.Count = 1
    pvw.Item = MagicMock(side_effect=lambda i: pv if i == 1 else None)
    xl.ProtectedViewWindows = pvw

    with patch.dict(
        sys.modules,
        _fake_win32_modules(xl, protected_view_windows=pvw),
        clear=False,
    ):
        wb_out, err = ComWorkbookService._get_open_workbook_com(u)

    assert wb_out is None
    assert "Protected View" in (err or "")


def test_get_open_workbook_com_unsaved_single_workbook(book_path):
    wb = MagicMock()
    wb.FullName = "Book1"
    wb.Path = ""
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        wb_out, err = ComWorkbookService._get_open_workbook_com(book_path)

    assert wb_out is None
    assert (
        err
        == "Error: Workbook has no saved path on disk; save the workbook to a known path before COM routing."
    )


def test_write_cell_grid_surfaces_read_only_error(book_path):
    wb = _workbook_mock(book_path, {"Sheet1": MagicMock()})
    wb.ReadOnly = True
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.write_cell_grid(book_path, "Sheet1", [[1]], "A1")

    assert msg == "Error: Workbook is read-only in Excel; COM routing cannot modify this workbook."


def test_list_open_workbooks_empty_collection():
    xl = MagicMock()
    xl.ActiveWorkbook = None
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_open_workbooks()

    assert json.loads(raw) == {"workbooks": []}


def test_list_open_workbooks_order_and_active_flag(tmp_path):
    p1 = str((tmp_path / "a.xlsx").resolve())
    p2 = str((tmp_path / "b.xlsx").resolve())
    wb1 = MagicMock()
    wb1.FullName = p1
    wb1.Name = "a.xlsx"
    wb2 = MagicMock()
    wb2.FullName = p2
    wb2.Name = "b.xlsx"

    xl = MagicMock()
    xl.ActiveWorkbook = wb2
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 2
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb1 if i == 1 else wb2)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_open_workbooks()

    data = json.loads(raw)["workbooks"]
    assert len(data) == 2
    assert data[0]["full_name"] == p1
    assert data[0]["name"] == "a.xlsx"
    assert data[0]["is_active"] is False
    assert data[1]["full_name"] == p2
    assert data[1]["name"] == "b.xlsx"
    assert data[1]["is_active"] is True


def test_list_open_workbooks_no_running_excel():
    client_mod = types.ModuleType("win32com.client")
    client_mod.GetActiveObject = MagicMock(side_effect=RuntimeError("RPC"))
    pkg = types.ModuleType("win32com")
    pkg.client = client_mod
    modules = {"win32com": pkg, "win32com.client": client_mod}

    with patch.dict(sys.modules, modules, clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.list_open_workbooks()

    assert msg == "Error: No running Excel application found"


def test_evaluate_range_whole_sheet(book_path):
    ws = MagicMock()
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.evaluate_range(book_path, "Sheet1")

    ws.Calculate.assert_called_once()
    assert "Recalculated sheet 'Sheet1'" in msg
    assert "save_workbook" in msg


def test_evaluate_range_cell_range(book_path):
    ws = MagicMock()
    rng = MagicMock()
    ws.Range = MagicMock(return_value=rng)
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.evaluate_range(book_path, "Sheet1", "A1", "B2")

    ws.Range.assert_called_once_with("A1", "B2")
    rng.Calculate.assert_called_once()
    assert "range A1:B2" in msg


def test_evaluate_range_workbook_not_open(book_path):
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.evaluate_range(book_path, "Sheet1")

    assert msg.startswith("Error:") and "Workbook" in msg


def test_evaluate_range_no_running_excel():
    client_mod = types.ModuleType("win32com.client")
    client_mod.GetActiveObject = MagicMock(side_effect=RuntimeError("RPC"))
    pkg = types.ModuleType("win32com")
    pkg.client = client_mod
    modules = {"win32com": pkg, "win32com.client": client_mod}

    with patch.dict(sys.modules, modules, clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.evaluate_range("C:\\fake\\book.xlsx", "Sheet1")

    assert msg == "Error: No running Excel application found"


def test_export_worksheet_table_com_backend(book_path):
    ws = MagicMock()
    rng = MagicMock()
    rng.Value2 = (("Col1", "Col2"), ("a", 1), ("b", 2))

    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = 3
    used.Columns.Count = 2
    ws.UsedRange = used

    cell = MagicMock()
    cell.Resize = MagicMock(return_value=rng)
    ws.Cells = MagicMock(return_value=cell)

    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.export_worksheet_table(book_path, "Sheet1")

    data = json.loads(raw)
    assert data["headers"] == ["Col1", "Col2"]
    assert data["rows"] == [["a", 1], ["b", 2]]
    assert data["row_count"] == 2
    assert data["truncated"] is False


def test_export_worksheet_table_com_truncates(book_path):
    ws = MagicMock()
    rng = MagicMock()
    rng.Value2 = (("H",), (0,), (1,), (2,))

    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = 4
    used.Columns.Count = 1
    ws.UsedRange = used

    cell = MagicMock()
    cell.Resize = MagicMock(return_value=rng)
    ws.Cells = MagicMock(return_value=cell)

    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.export_worksheet_table(book_path, "Sheet1", max_rows=2)

    data = json.loads(raw)
    assert data["headers"] == ["H"]
    assert len(data["rows"]) == 2
    assert data["row_count"] == 3
    assert data["truncated"] is True
