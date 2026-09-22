"""ComWorkbookService: executor wiring and COM workers without Excel (Story 7-1)."""

from __future__ import annotations

import json
import os
import sys
import types
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from excel_mcp.routing.com_workbook_service import (
    ComWorkbookService,
    _should_fallback_to_direct_read_com,
)


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


def test_should_fallback_stratified_sampling_large_sparse_matrix():
    """Large ranges probe across rows, not only the first blank cells."""
    nrows, ncols = 100, 10
    matrix = [[None] * ncols for _ in range(nrows)]
    ws = MagicMock()

    def _cell(r, c):
        m = MagicMock()
        m.Value2 = "x" if r >= 51 else None
        return m

    ws.Cells = MagicMock(side_effect=_cell)
    assert _should_fallback_to_direct_read_com(ws, matrix, 1, 1, nrows, ncols) is True


def test_should_fallback_no_trigger_when_direct_also_blank():
    nrows, ncols = 100, 10
    matrix = [[None] * ncols for _ in range(nrows)]
    ws = MagicMock()
    ws.Cells = MagicMock(return_value=MagicMock(Value2=None))
    assert _should_fallback_to_direct_read_com(ws, matrix, 1, 1, nrows, ncols) is False


def test_should_fallback_small_range_row_major():
    matrix = [[None, None], [None, None]]
    ws = MagicMock()
    vals = {(1, 1): "a", (1, 2): "b", (2, 1): "c", (2, 2): "d"}
    ws.Cells = MagicMock(side_effect=lambda r, c: MagicMock(Value2=vals[(r, c)]))
    assert _should_fallback_to_direct_read_com(ws, matrix, 1, 1, 2, 2) is True


def test_read_range_large_sparse_matrix_fallback(book_path):
    """Integration: bulk Value2 all-null on a large range falls back to direct reads."""
    nrows, ncols = 80, 5
    ws = MagicMock()
    rng = MagicMock()
    rng.Value2 = tuple([tuple([None] * ncols) for _ in range(nrows)])

    used = MagicMock()
    used.Row = 1
    used.Column = 1
    used.Rows.Count = nrows
    used.Columns.Count = ncols
    ws.UsedRange = used

    cells: dict[tuple[int, int], MagicMock] = {}
    for ir in range(nrows):
        for ic in range(ncols):
            r, c = ir + 1, ic + 1
            cell = MagicMock()
            cell.Validation = MagicMock()
            cell.Validation.Type = 0
            cell.Value2 = f"R{r}C{c}" if ir >= nrows // 2 else None
            cells[(r, c)] = cell
    cells[(1, 1)].Resize = MagicMock(return_value=rng)
    ws.Cells = MagicMock(side_effect=lambda r, c: cells[(r, c)])
    wb = _workbook_mock(book_path, {"Sheet1": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.read_range_with_metadata(book_path, "Sheet1", "A1", "E80")

    data = json.loads(raw)
    populated = [c for c in data["cells"] if c["value"] is not None]
    assert len(populated) == (nrows // 2) * ncols
    assert populated[0]["value"] == "R41C1"


def _read_range_fixture(
    book_path: str,
    *,
    rng: MagicMock,
    cells: dict[tuple[int, int], MagicMock],
    used_rows: int = 1,
    used_cols: int = 1,
    value_mode: str = "value",
    metadata_mode: str = "full",
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
            book_path,
            "Sheet1",
            "A1",
            end_cell,
            value_mode=value_mode,
            metadata_mode=metadata_mode,
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


def test_read_range_with_metadata_invalid_metadata_mode_returns_error(book_path):
    with patch.dict(sys.modules, _fake_win32_modules(MagicMock()), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.read_range_with_metadata(
            book_path, "Sheet1", "A1", metadata_mode="sparse"
        )
    assert msg.startswith("Error:")
    assert "Invalid metadata_mode" in msg


def test_read_range_with_metadata_compact_omits_validation(book_path):
    rng = MagicMock()
    rng.Value2 = "yes"

    cell = MagicMock()
    cell.Value2 = "yes"
    cell.Text = "yes"
    cell.Validation = MagicMock()
    cell.Validation.Type = 3
    cell.Validation.IgnoreBlank = True
    cell.Validation.Formula1 = '"yes,no"'

    raw = _read_range_fixture(
        book_path,
        rng=rng,
        cells={(1, 1): cell},
        metadata_mode="compact",
    )
    data = json.loads(raw)
    assert data["metadata_mode"] == "compact"
    assert "validation" not in data["cells"][0]


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


def test_create_workbook_saveas_failure_closes_orphan(book_path):
    wb_new = MagicMock()
    wb_new.SaveAs.side_effect = RuntimeError("save failed")
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Add = MagicMock(return_value=wb_new)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.create_workbook(book_path)

    assert msg.startswith("Error:")
    assert "save failed" in msg
    wb_new.Close.assert_called_once_with(SaveChanges=False)


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
        wb_write, err_write = ComWorkbookService._get_open_workbook_com(
            book_path, for_write=True
        )
        close_msg = ComWorkbookService._close_workbook_in_excel_com(book_path, False)

    assert wb_out is wb
    assert err is None
    assert wb_write is None
    assert err_write == "Error: Workbook is read-only in Excel; COM routing cannot modify this workbook."
    assert close_msg.startswith("Workbook closed")


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


def test_list_open_workbooks_active_context(tmp_path):
    p1 = str((tmp_path / "a.xlsx").resolve())
    wb1 = MagicMock()
    wb1.FullName = p1
    wb1.Name = "a.xlsx"
    active = MagicMock()
    active.Name = "Sheet1"

    xl = MagicMock()
    xl.ActiveWorkbook = wb1
    xl.ActiveSheet = active
    xl.Selection = MagicMock()
    xl.Selection.Address = "$B$2:$D$5"
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb1)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_open_workbooks(detail="active_context")

    data = json.loads(raw)
    assert data["active_workbook"] == {"full_name": p1, "name": "a.xlsx"}
    assert data["active_sheet"] == "Sheet1"
    assert data["selection"] == "B2:D5"
    assert len(data["workbooks"]) == 1
    assert data["workbooks"][0]["is_active"] is True


def test_list_open_workbooks_active_context_no_active_workbook():
    xl = MagicMock()
    xl.ActiveWorkbook = None
    xl.ActiveSheet = None
    xl.Selection = None
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_open_workbooks(detail="active_context")

    data = json.loads(raw)
    assert data["active_workbook"] is None
    assert data["active_sheet"] is None
    assert data["selection"] is None
    assert data["workbooks"] == []


def test_list_open_workbooks_invalid_detail():
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 0

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        msg = svc.list_open_workbooks(detail="full")

    assert msg.startswith("Error:")
    assert "active_context" in msg


def test_list_open_workbooks_default_detail_is_minimal(tmp_path):
    p1 = str((tmp_path / "a.xlsx").resolve())
    wb1 = MagicMock()
    wb1.FullName = p1
    wb1.Name = "a.xlsx"

    xl = MagicMock()
    xl.ActiveWorkbook = wb1
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb1)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_open_workbooks()

    data = json.loads(raw)
    assert set(data.keys()) == {"workbooks"}


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
    cell.Resize.assert_called_once_with(3, 1)


def test_export_worksheet_table_com_invalid_max_rows(book_path):
    with patch.dict(sys.modules, _fake_win32_modules(MagicMock()), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.export_worksheet_table(book_path, "Sheet1", max_rows=0)

    assert raw == "Error: max_rows must be a positive integer"


def _list_object_mock(
    *,
    name: str,
    range_addr: str,
    header_addr: str,
    data_addr: str | None,
    data_rows: int,
    column_names: list[str],
    filter_mode: bool,
) -> MagicMock:
    lo = MagicMock()
    lo.Name = name
    lo.Range.Address = range_addr
    lo.HeaderRowRange.Address = header_addr
    if data_addr is None:
        lo.DataBodyRange = None
    else:
        body = MagicMock()
        body.Address = data_addr
        body.Rows.Count = data_rows
        lo.DataBodyRange = body
    lo.ListRows.Count = data_rows
    cols = MagicMock()
    cols.Count = len(column_names)

    def _col_item(i: int) -> MagicMock:
        c = MagicMock()
        c.Name = column_names[i - 1]
        return c

    cols.Item = MagicMock(side_effect=_col_item)
    lo.ListColumns = cols
    af = MagicMock()
    af.FilterMode = filter_mode
    # Sentinel attributes that must never be invoked on a catalog read.
    af.ShowAllData = MagicMock()
    lo.AutoFilter = af
    lo.Sort = MagicMock()
    return lo


def test_list_tables_com_schema_and_empty(book_path):
    lo = _list_object_mock(
        name="ap_sw_components",
        range_addr="$B$2:$D$4",
        header_addr="$B$2:$D$2",
        data_addr="$B$3:$D$4",
        data_rows=2,
        column_names=["ColA", "ColB", "ColC"],
        filter_mode=False,
    )
    ws_tables = MagicMock()
    ws_tables.Name = "software components"
    ws_tables.ListObjects.Count = 1
    ws_tables.ListObjects.Item = MagicMock(return_value=lo)
    # Must not use worksheet AutoFilterMode as the filter signal.
    ws_tables.AutoFilterMode = True
    ws_tables.AutoFilter = MagicMock()
    ws_tables.AutoFilter.ShowAllData = MagicMock()

    ws_plain = MagicMock()
    ws_plain.Name = "stackfleet concept"
    ws_plain.ListObjects.Count = 0

    sheets = MagicMock()
    sheets.Count = 2
    sheets.Item = MagicMock(side_effect=lambda i: ws_tables if i == 1 else ws_plain)

    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        schema = json.loads(svc.list_tables(book_path, detail="schema"))
        minimal = json.loads(svc.list_tables(book_path, detail="minimal"))
        empty_ws = MagicMock()
        empty_ws.Name = "Only"
        empty_ws.ListObjects.Count = 0
        sheets.Count = 1
        sheets.Item = MagicMock(return_value=empty_ws)
        empty = json.loads(svc.list_tables(book_path))

    assert len(schema["tables"]) == 1
    entry = schema["tables"][0]
    assert entry["sheet"] == "software components"
    assert entry["name"] == "ap_sw_components"
    assert entry["range"] == "B2:D4"
    assert entry["header_range"] == "B2:D2"
    assert entry["data_range"] == "B3:D4"
    assert entry["row_count"] == 2
    assert entry["filter_applied"] is False
    assert entry["columns"] == ["ColA", "ColB", "ColC"]
    assert "columns" not in minimal["tables"][0]
    assert empty == {"tables": []}
    lo.AutoFilter.ShowAllData.assert_not_called()
    ws_tables.AutoFilter.ShowAllData.assert_not_called()
    lo.Sort.assert_not_called()


def test_list_tables_com_filter_mode_from_listobject(book_path):
    lo = _list_object_mock(
        name="T1",
        range_addr="$A$1:$B$3",
        header_addr="$A$1:$B$1",
        data_addr="$A$2:$B$3",
        data_rows=2,
        column_names=["H1", "H2"],
        filter_mode=True,
    )
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.ListObjects.Count = 1
    ws.ListObjects.Item = MagicMock(return_value=lo)
    ws.AutoFilterMode = False

    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.list_tables(book_path, detail="minimal"))

    assert data["tables"][0]["filter_applied"] is True


def test_list_tables_com_invalid_detail(book_path):
    with patch.dict(sys.modules, _fake_win32_modules(MagicMock()), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.list_tables(book_path, detail="full")
    assert raw.startswith("Error:")
    assert "detail" in raw.lower()


def _query_list_object_mock(
    *,
    name: str,
    column_data: dict[str, list],
) -> MagicMock:
    """ListObject mock that exposes per-column DataBodyRange only."""
    column_names = list(column_data.keys())
    lo = MagicMock()
    lo.Name = name
    lo.Range.Address = "$A$1:$C$4"
    lo.HeaderRowRange.Address = "$A$1:$C$1"
    # Whole-table body must not be required for query_table.
    lo.DataBodyRange = MagicMock()
    lo.DataBodyRange.Value2 = MagicMock(
        side_effect=AssertionError("must not read full DataBodyRange")
    )
    cols = MagicMock()
    cols.Count = len(column_names)
    data_reads: list[str] = []

    def _col_item(key):
        if isinstance(key, int):
            col_name = column_names[key - 1]
        else:
            col_name = str(key)
            if col_name not in column_data:
                raise KeyError(col_name)
        col = MagicMock()
        col.Name = col_name
        values = column_data[col_name]

        class _Body:
            @property
            def Value2(self):
                data_reads.append(col_name)
                if not values:
                    return None
                if len(values) == 1:
                    return values[0]
                return tuple((v,) for v in values)

        col.DataBodyRange = _Body()
        return col

    cols.Item = MagicMock(side_effect=_col_item)
    lo.ListColumns = cols
    af = MagicMock()
    af.FilterMode = True
    af.ShowAllData = MagicMock()
    lo.AutoFilter = af
    lo.Sort = MagicMock()
    lo._data_reads = data_reads  # type: ignore[attr-defined]
    return lo


def test_query_table_com_selected_columns_only(book_path):
    lo = _query_list_object_mock(
        name="Items",
        column_data={
            "Name": ["a", "b", "c"],
            "Qty": [1, 2, 3],
            "Tag": ["x", "y", "x"],
            "Wide": list(range(3)),
        },
    )
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.ListObjects.Count = 1
    ws.ListObjects.Item = MagicMock(return_value=lo)
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(
            svc.query_table(
                book_path,
                "Items",
                columns=["Name", "Qty"],
                where=[{"column": "Tag", "op": "eq", "value": "x"}],
            )
        )

    assert data["rows"] == [{"Name": "a", "Qty": 1}, {"Name": "c", "Qty": 3}]
    assert data["row_count"] == 2
    assert data["truncated"] is False
    # Only projection + filter columns were read (not Wide).
    assert set(lo._data_reads) == {"Name", "Qty", "Tag"}
    lo.AutoFilter.ShowAllData.assert_not_called()
    lo.Sort.assert_not_called()
    assert lo.AutoFilter.FilterMode is True


def test_query_table_com_unknown_table(book_path):
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.ListObjects.Count = 0
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.query_table(book_path, "Missing")
    assert raw.startswith("Error:")
    assert "Table" in raw


def test_map_sheet_layout_and_query_region_com_no_mutation(book_path):
    # UsedRange starts at B4 (header not on worksheet row 1).
    matrix = (
        ("Name", "Qty"),
        ("alpha", 1),
        ("beta", 2),
    )
    lo = _list_object_mock(
        name="SideTable",
        range_addr="$E$1:$E$2",
        header_addr="$E$1",
        data_addr="$E$2",
        data_rows=1,
        column_names=["TCol"],
        filter_mode=False,
    )
    ws = MagicMock()
    ws.Name = "stackfleet concept"
    used = MagicMock()
    used.Row = 4
    used.Column = 2
    used.Value2 = matrix
    ws.UsedRange = used
    ws.ListObjects.Count = 1
    ws.ListObjects.Item = MagicMock(return_value=lo)
    ws.ListObjects.Add = MagicMock(
        side_effect=AssertionError("must not call ListObjects.Add")
    )
    region_rng = MagicMock()
    region_rng.Value2 = matrix
    ws.Range = MagicMock(return_value=region_rng)
    wb = _workbook_mock(book_path, {"stackfleet concept": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        layout = json.loads(svc.map_sheet_layout(book_path, "stackfleet concept"))
        assert layout["tables"][0]["name"] == "SideTable"
        assert layout["tables"][0]["kind"] == "table"
        assert len(layout["regions"]) == 1
        region = layout["regions"][0]
        assert region["header_row"] == 4
        assert region["header_guess"] is True
        assert region["id"] == "stackfleet concept!B4:C6"

        data = json.loads(
            svc.query_region(
                book_path,
                "stackfleet concept",
                region_id=region["id"],
                where=[{"column": "Name", "op": "eq", "value": "alpha"}],
            )
        )
    assert data["rows"] == [{"Name": "alpha", "Qty": 1}]
    assert data["view_spec"]["target"]["kind"] == "region"
    ws.ListObjects.Add.assert_not_called()
    lo.AutoFilter.ShowAllData.assert_not_called()
    lo.Sort.assert_not_called()


def _view_list_object_mock(
    *,
    name: str,
    header_addr: str,
    column_names: list[str],
    sheet_start_col: int,
    filter_mode: bool = False,
    prior_filter_on_field: int | None = None,
    sort_capturable: bool = True,
    sort_fields: list[tuple[str, str]] | None = None,
    hidden_sheet_cols: set[int] | None = None,
) -> tuple[MagicMock, MagicMock]:
    """ListObject + worksheet mocks for apply/clear table view tests."""
    hidden_sheet_cols = set(hidden_sheet_cols or ())
    lo = MagicMock()
    lo.Name = name
    lo.HeaderRowRange.Address = header_addr
    lo.Range = MagicMock()
    lo.Range.AutoFilter = MagicMock()

    cols = MagicMock()
    cols.Count = len(column_names)

    def _col_item(key):
        if isinstance(key, int):
            idx = key
            col_name = column_names[key - 1]
        else:
            col_name = str(key)
            idx = column_names.index(col_name) + 1
        col = MagicMock()
        col.Name = col_name
        sheet_col = sheet_start_col + idx - 1
        col.Range.Column = sheet_col
        return col

    cols.Item = MagicMock(side_effect=_col_item)
    lo.ListColumns = cols

    af = MagicMock()
    af.FilterMode = filter_mode
    af.ShowAllData = MagicMock()
    filters = MagicMock()
    if prior_filter_on_field is not None:
        filters.Count = len(column_names)

        def _filt_item(i: int):
            f = MagicMock()
            f.On = i == prior_filter_on_field
            f.Criteria1 = "old" if f.On else None
            f.Criteria2 = None
            f.Operator = 0
            return f

        filters.Item = MagicMock(side_effect=_filt_item)
    else:
        filters.Count = 0
        filters.Item = MagicMock(side_effect=lambda i: MagicMock(On=False))
    af.Filters = filters
    lo.AutoFilter = af

    sort_obj = MagicMock()
    sort_fields_coll = MagicMock()
    recorded_sort: list = []

    if sort_fields and sort_capturable:
        sort_fields_coll.Count = len(sort_fields)

        def _sf_item(i: int):
            col_name, order = sort_fields[i - 1]
            field_idx = column_names.index(col_name) + 1
            sf = MagicMock()
            sf.CustomOrder = None
            sf.SortOn = 0
            sf.Key.Column = sheet_start_col + field_idx - 1
            sf.Order = 2 if order == "desc" else 1
            return sf

        sort_fields_coll.Item = MagicMock(side_effect=_sf_item)
    elif not sort_capturable:
        sort_fields_coll.Count = 1
        sf = MagicMock()
        sf.CustomOrder = "CustomList"
        sf.SortOn = 0
        sf.Key.Column = sheet_start_col
        sf.Order = 1
        sort_fields_coll.Item = MagicMock(return_value=sf)
    else:
        sort_fields_coll.Count = 0

    def _sort_add(**kwargs):
        recorded_sort.append(kwargs)
        return MagicMock()

    sort_fields_coll.Clear = MagicMock()
    sort_fields_coll.Add = MagicMock(side_effect=_sort_add)
    sort_obj.SortFields = sort_fields_coll
    sort_obj.Apply = MagicMock()
    sort_obj.Header = None
    lo.Sort = sort_obj
    lo._recorded_sort = recorded_sort  # type: ignore[attr-defined]

    ws = MagicMock()
    ws.Name = "software components"
    ws.ListObjects.Count = 1
    ws.ListObjects.Item = MagicMock(return_value=lo)
    ws.ListObjects.Add = MagicMock(
        side_effect=AssertionError("must not call ListObjects.Add")
    )
    ws.AutoFilterMode = False
    ws.AutoFilter = MagicMock()
    ws.AutoFilter.ShowAllData = MagicMock(
        side_effect=AssertionError("must not clear via Worksheet.AutoFilter")
    )
    ws.Copy = MagicMock(side_effect=AssertionError("must not call Worksheet.Copy"))

    col_state: dict[int, bool] = {c: True for c in hidden_sheet_cols}

    def _columns(sheet_col: int):
        col = MagicMock()

        class _Entire:
            @property
            def Hidden(self):
                return col_state.get(int(sheet_col), False)

            @Hidden.setter
            def Hidden(self, value):
                col_state[int(sheet_col)] = bool(value)

        col.EntireColumn = _Entire()
        return col

    ws.Columns = MagicMock(side_effect=_columns)
    ws._col_state = col_state  # type: ignore[attr-defined]
    lo.Parent = ws
    return lo, ws


def test_apply_table_view_com_filter_sort_focus_and_field_index(book_path):
    # Table starts at column B → field 1 is sheet col 2, not column A.
    lo, ws = _view_list_object_mock(
        name="ap_sw_components",
        header_addr="$B$2:$D$2",
        column_names=["Name", "Qty", "Tag"],
        sheet_start_col=2,
        filter_mode=False,
        sort_fields=[("Name", "asc")],
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "ap_sw_components"},
        "columns": ["Name", "Qty"],
        "where": [
            {"column": "Tag", "op": "eq", "value": "x"},
            {"column": "Qty", "op": "gte", "value": 2},
        ],
        "sort": {"by": [{"column": "Qty", "order": "asc"}]},
    }

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(
            svc.apply_table_view(book_path, view_spec, mode="in_place")
        )

    assert data["restore_token"]["table"] == "ap_sw_components"
    assert data["restore_token"]["sort_applied"] is True
    # Tag is field 3 (table-relative), not sheet column index.
    applied_fields = [
        (c.kwargs or {}).get("Field") for c in lo.Range.AutoFilter.call_args_list
    ]
    assert 3 in applied_fields  # Tag
    assert 2 in applied_fields  # Qty
    # Column focus hid Tag (sheet col 4) only.
    assert data["restore_token"]["columns_hidden"] == [4]
    assert ws._col_state.get(4) is True
    assert lo._recorded_sort  # sort applied
    ws.ListObjects.Add.assert_not_called()
    ws.Copy.assert_not_called()


def test_apply_table_view_com_two_value_xlor(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$B$1",
        column_names=["Name", "Tag"],
        sheet_start_col=1,
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name", "Tag"],
        "where": [{"column": "Tag", "op": "in", "value": ["x", "y"]}],
        "sort": None,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        json.loads(svc.apply_table_view(book_path, view_spec))

    kwargs = lo.Range.AutoFilter.call_args.kwargs
    assert kwargs["Field"] == 2
    assert kwargs["Criteria1"] == "x"
    assert kwargs["Criteria2"] == "y"
    assert kwargs["Operator"] == 2  # xlOr


def test_apply_table_view_com_rejects_without_mutation(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$A$1",
        column_names=["Name"],
        sheet_start_col=1,
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    bad_specs = [
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [],
                "sort": None,
                "limit": 5,
            },
            "in_place",
            None,
        ),
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [],
                "sort": None,
                "offset": 1,
            },
            "in_place",
            None,
        ),
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [{"column": "Name", "op": "is_empty"}],
                "sort": None,
            },
            "in_place",
            None,
        ),
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [{"column": "Name", "op": "is_empty"}],
                "sort": None,
            },
            "snapshot",
            None,
        ),
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [
                    {
                        "op": "or",
                        "clauses": [
                            {"column": "Name", "op": "eq", "value": "a"},
                            {"column": "Name", "op": "eq", "value": "b"},
                        ],
                    }
                ],
                "sort": None,
            },
            "in_place",
            None,
        ),
        (
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [
                    {
                        "op": "or",
                        "clauses": [
                            {"column": "Name", "op": "eq", "value": "a"},
                            {"column": "Name", "op": "eq", "value": "b"},
                        ],
                    }
                ],
                "sort": None,
            },
            "snapshot",
            None,
        ),
    ]

    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        for spec, mode, va in bad_specs:
            raw = svc.apply_table_view(
                book_path, spec, mode=mode, view_applicability=va
            )
            assert raw.startswith("Error:"), raw
            lo.Range.AutoFilter.assert_not_called()
            lo.AutoFilter.ShowAllData.assert_not_called()
            lo.Sort.SortFields.Clear.assert_not_called()
            lo.Sort.Apply.assert_not_called()

        # view_applicability.viewable false
        raw = svc.apply_table_view(
            book_path,
            {
                "target": {"kind": "table", "name": "Items"},
                "columns": ["Name"],
                "where": [{"column": "Name", "op": "eq", "value": "a"}],
                "sort": None,
            },
            view_applicability={
                "viewable": False,
                "reason": "limit truncates matching rows",
            },
        )
        assert raw.startswith("Error:")
        assert "limit" in raw.lower() or "viewable" in raw.lower()
        lo.Range.AutoFilter.assert_not_called()


def test_clear_table_view_com_showalldata_and_selective_unhide(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$B$1:$D$1",
        column_names=["A", "B", "C"],
        sheet_start_col=2,
        filter_mode=True,
        hidden_sheet_cols={4},  # already hidden by user (sheet col D)
    )
    # Pretend apply hid sheet col 3 (field B).
    ws._col_state[3] = True
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    token = {
        "v": 1,
        "kind": "listobject",
        "table": "Items",
        "sheet": "software components",
        "prior_filters": [],
        "prior_sort": [],
        "sort_applied": False,
        "columns_hidden": [3],
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.clear_table_view(book_path, token))

    assert data["table"] == "Items"
    lo.AutoFilter.ShowAllData.assert_called_once()
    ws.AutoFilter.ShowAllData.assert_not_called()
    # Unhid only column this call hid (3); user-hidden 4 stays hidden.
    assert ws._col_state.get(3) is False
    assert ws._col_state.get(4) is True


def test_apply_table_view_com_sort_refusal_when_uncapturable(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$B$1",
        column_names=["Name", "Qty"],
        sheet_start_col=1,
        sort_capturable=False,
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name", "Qty"],
        "where": [{"column": "Name", "op": "eq", "value": "a"}],
        "sort": {"column": "Qty", "order": "asc"},
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.apply_table_view(book_path, view_spec))

    assert data["restore_token"]["sort_applied"] is False
    assert any("sort" in w.lower() for w in data.get("warnings", []))
    lo.Sort.SortFields.Clear.assert_not_called()
    lo.Sort.Apply.assert_not_called()
    # Filter still applied.
    assert lo.Range.AutoFilter.called


def test_apply_table_view_com_sort_refusal_when_no_prior_sort(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$B$1",
        column_names=["Name", "Qty"],
        sheet_start_col=1,
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name", "Qty"],
        "where": [{"column": "Name", "op": "eq", "value": "a"}],
        "sort": {"by": [{"column": "Qty", "order": "asc"}]},
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.apply_table_view(book_path, view_spec))

    assert data["restore_token"]["sort_applied"] is False
    assert data["restore_token"]["prior_sort"] == []
    assert any("sort" in w.lower() for w in data.get("warnings", []))
    lo.Sort.SortFields.Clear.assert_not_called()
    lo.Sort.Apply.assert_not_called()
    assert lo.Range.AutoFilter.called


def test_apply_and_clear_table_view_com_sort_restore(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$B$1",
        column_names=["Name", "Qty"],
        sheet_start_col=1,
        sort_fields=[("Name", "desc")],
    )
    sheets = MagicMock()
    sheets.Count = 1
    sheets.Item = MagicMock(return_value=ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name", "Qty"],
        "where": [],
        "sort": {"by": [{"column": "Qty", "order": "asc"}]},
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        applied = json.loads(svc.apply_table_view(book_path, view_spec))
        token = applied["restore_token"]
        assert token["sort_applied"] is True
        assert token["prior_sort"] == [{"column": "Name", "order": "desc"}]
        lo._recorded_sort.clear()
        json.loads(svc.clear_table_view(book_path, token))

    # Clear restored prior sort (Name desc).
    assert lo.Sort.SortFields.Clear.called
    assert any(
        c.get("Order") == 2 for c in lo._recorded_sort
    )  # xlDescending restored


def _view_plain_range_mock(
    *,
    sheet_name: str,
    range_a1: str,
    column_names: list[str],
    sheet_start_col: int = 1,
    prior_autofilter_mode: bool = False,
    filter_mode: bool = False,
    hidden_sheet_cols: set[int] | None = None,
    sort_capturable: bool = True,
    sort_fields: list[tuple[str, str]] | None = None,
) -> tuple[MagicMock, MagicMock]:
    """Worksheet + Range mocks for plain-region apply/clear view tests."""
    hidden_sheet_cols = set(hidden_sheet_cols or ())
    min_col = sheet_start_col
    # Header + one data row so Value2 looks like a real region.
    header = list(column_names)
    data = [f"v{i}" for i in range(len(column_names))]
    matrix = [header, data]

    rng = MagicMock()
    rng.Value2 = matrix
    recorded_af: list = []

    def _af(**kwargs):
        recorded_af.append(kwargs)
        ws.AutoFilterMode = True
        ws.FilterMode = True
        return True

    rng.AutoFilter = MagicMock(side_effect=_af)

    def _rng_columns(field: int):
        col = MagicMock()
        col.Column = min_col + int(field) - 1
        return col

    rng.Columns = MagicMock(side_effect=_rng_columns)

    ws = MagicMock()
    ws.Name = sheet_name
    ws.AutoFilterMode = prior_autofilter_mode
    ws.FilterMode = filter_mode
    ws.ListObjects.Count = 0
    ws.ListObjects.Item = MagicMock(
        side_effect=AssertionError("must not touch ListObjects")
    )
    ws.ListObjects.Add = MagicMock(
        side_effect=AssertionError("must not call ListObjects.Add")
    )
    ws.Copy = MagicMock(side_effect=AssertionError("must not call Worksheet.Copy"))

    af = MagicMock()
    af.FilterMode = filter_mode
    af.ShowAllData = MagicMock()
    filters = MagicMock()
    filters.Count = 0
    filters.Item = MagicMock(side_effect=lambda i: MagicMock(On=False))
    af.Filters = filters
    ws.AutoFilter = af
    ws.ShowAllData = MagicMock()

    sort_obj = MagicMock()
    sort_fields_coll = MagicMock()
    recorded_sort: list = []

    if sort_fields and sort_capturable:
        sort_fields_coll.Count = len(sort_fields)

        def _sf_item(i: int):
            col_name, order = sort_fields[i - 1]
            field_idx = column_names.index(col_name) + 1
            sf = MagicMock()
            sf.CustomOrder = None
            sf.SortOn = 0
            sf.Key.Column = min_col + field_idx - 1
            sf.Order = 2 if order == "desc" else 1
            return sf

        sort_fields_coll.Item = MagicMock(side_effect=_sf_item)
    elif not sort_capturable:
        sort_fields_coll.Count = 1
        sf = MagicMock()
        sf.CustomOrder = "CustomList"
        sf.SortOn = 0
        sf.Key.Column = min_col
        sf.Order = 1
        sort_fields_coll.Item = MagicMock(return_value=sf)
    else:
        sort_fields_coll.Count = 0

    def _sort_add(**kwargs):
        recorded_sort.append(kwargs)
        return MagicMock()

    sort_fields_coll.Clear = MagicMock()
    sort_fields_coll.Add = MagicMock(side_effect=_sort_add)
    sort_obj.SortFields = sort_fields_coll
    sort_obj.SetRange = MagicMock()
    sort_obj.Apply = MagicMock()
    sort_obj.Header = None
    ws.Sort = sort_obj
    ws._recorded_sort = recorded_sort  # type: ignore[attr-defined]
    ws._recorded_af = recorded_af  # type: ignore[attr-defined]

    col_state: dict[int, bool] = {c: True for c in hidden_sheet_cols}

    def _columns(sheet_col: int):
        col = MagicMock()

        class _Entire:
            @property
            def Hidden(self):
                return col_state.get(int(sheet_col), False)

            @Hidden.setter
            def Hidden(self, value):
                col_state[int(sheet_col)] = bool(value)

        col.EntireColumn = _Entire()
        return col

    ws.Columns = MagicMock(side_effect=_columns)
    ws._col_state = col_state  # type: ignore[attr-defined]
    ws.Range = MagicMock(return_value=rng)
    return ws, rng


def test_apply_table_view_com_region_autofilter_no_listobjects_add(book_path):
    ws, rng = _view_plain_range_mock(
        sheet_name="stackfleet concept",
        range_a1="B4:D6",
        column_names=["Name", "Qty", "Tag"],
        sheet_start_col=2,
        prior_autofilter_mode=False,
    )
    wb = _workbook_mock(book_path, {"stackfleet concept": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {
            "kind": "region",
            "sheet": "stackfleet concept",
            "range": "B4:D6",
        },
        "columns": ["Name", "Qty"],
        "where": [{"column": "Tag", "op": "eq", "value": "x"}],
        "sort": None,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.apply_table_view(book_path, view_spec, mode="in_place"))

    token = data["restore_token"]
    assert token["kind"] == "range"
    assert token["range"] == "B4:D6"
    assert token["sheet"] == "stackfleet concept"
    assert token["prior_autofilter_mode"] is False
    assert token["columns_hidden"] == [4]  # Tag at sheet col 4
    assert ws.AutoFilterMode is True
    assert ws.FilterMode is True
    assert rng.AutoFilter.called
    applied_fields = [(c or {}).get("Field") for c in ws._recorded_af if c]
    assert 3 in applied_fields  # Tag is field 3 in B:D
    ws.ListObjects.Add.assert_not_called()
    ws.Copy.assert_not_called()


def test_clear_table_view_com_region_restores_autofilter_mode(book_path):
    ws, rng = _view_plain_range_mock(
        sheet_name="stackfleet concept",
        range_a1="A1:C3",
        column_names=["Name", "Qty", "Tag"],
        sheet_start_col=1,
        prior_autofilter_mode=False,
        filter_mode=True,
        hidden_sheet_cols={3},  # user-hidden Tag
    )
    # Pretend apply hid Qty (col 2).
    ws._col_state[2] = True
    ws.AutoFilterMode = True
    ws.FilterMode = True
    wb = _workbook_mock(book_path, {"stackfleet concept": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    token = {
        "v": 1,
        "kind": "range",
        "sheet": "stackfleet concept",
        "range": "A1:C3",
        "prior_autofilter_mode": False,
        "prior_filters": [],
        "prior_sort": None,
        "sort_applied": False,
        "columns_hidden": [2],
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(svc.clear_table_view(book_path, token))

    assert data["range"] == "A1:C3"
    assert ws.AutoFilterMode is False
    # Unhid only column this call hid; user-hidden 3 stays hidden.
    assert ws._col_state.get(2) is False
    assert ws._col_state.get(3) is True
    ws.ListObjects.Add.assert_not_called()


def test_apply_and_clear_table_view_com_region_no_table_created(book_path):
    ws, rng = _view_plain_range_mock(
        sheet_name="plain",
        range_a1="A1:B3",
        column_names=["Name", "Qty"],
        sheet_start_col=1,
        prior_autofilter_mode=True,
        filter_mode=False,
    )
    wb = _workbook_mock(book_path, {"plain": ws})
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "region", "sheet": "plain", "range": "A1:B3"},
        "columns": ["Name", "Qty"],
        "where": [{"column": "Name", "op": "eq", "value": "a"}],
        "sort": None,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        applied = json.loads(svc.apply_table_view(book_path, view_spec))
        token = applied["restore_token"]
        assert token["prior_autofilter_mode"] is True
        assert ws.ListObjects.Count == 0
        json.loads(svc.clear_table_view(book_path, token))

    # Prior AutoFilterMode was True → arrows restored (not forced off).
    assert ws.AutoFilterMode is True
    ws.ListObjects.Add.assert_not_called()
    ws.Copy.assert_not_called()


def _attach_list_column_bodies(
    lo: MagicMock, column_data: dict[str, list]
) -> None:
    """Add per-column DataBodyRange.Value2 onto an existing ListColumns mock."""
    column_names = list(column_data.keys())
    original_item = lo.ListColumns.Item

    def _col_item(key):
        col = original_item(key)
        if isinstance(key, int):
            col_name = column_names[key - 1]
        else:
            col_name = str(key)
        values = column_data[col_name]

        class _Body:
            @property
            def Value2(self):
                if not values:
                    return None
                if len(values) == 1:
                    return values[0]
                return tuple((v,) for v in values)

        col.DataBodyRange = _Body()
        return col

    lo.ListColumns.Item = MagicMock(side_effect=_col_item)


class _SnapshotSheets:
    """Minimal Worksheets collection: Count/Item/Add/call-by-name/Delete."""

    def __init__(self, source_ws: MagicMock):
        self._sheets: list[Any] = [source_ws]
        self._by_name = {str(source_ws.Name): source_ws}
        self.added: list[Any] = []
        self.deleted: list[str] = []
        self.fail_next_rename = False

    @property
    def Count(self) -> int:
        return len(self._sheets)

    def Item(self, index: int) -> Any:
        return self._sheets[index - 1]

    def __call__(self, name: str) -> Any:
        key = str(name)
        if key not in self._by_name:
            raise RuntimeError(f"Sheet '{name}' not found")
        return self._by_name[key]

    def _rename(self, sheet: Any, old: str, new: str) -> None:
        if old in self._by_name and self._by_name[old] is sheet:
            del self._by_name[old]
        self._by_name[new] = sheet

    def Add(self) -> Any:
        store: dict = {}
        coll = self

        class _ValueRange:
            def Resize(self, nrows, ncols):
                return self

            @property
            def Value(self):
                return store.get("v")

            @Value.setter
            def Value(self, v):
                store["v"] = v

        class _NewSheet:
            def __init__(self) -> None:
                self._name = f"Sheet{len(coll._sheets) + 1}"
                self._value_store = store
                self.Range = lambda _addr: _ValueRange()
                self.ListObjects = MagicMock()
                self.ListObjects.Count = 0
                self.ListObjects.Add = MagicMock(
                    side_effect=AssertionError(
                        "must not call ListObjects.Add on snapshot"
                    )
                )
                self.Copy = MagicMock(
                    side_effect=AssertionError("must not call Worksheet.Copy")
                )
                self.Delete = MagicMock(side_effect=self._delete)

            @property
            def Name(self) -> str:
                return self._name

            @Name.setter
            def Name(self, value: str) -> None:
                if coll.fail_next_rename:
                    coll.fail_next_rename = False
                    raise RuntimeError("rename failed")
                new = str(value)
                coll._rename(self, self._name, new)
                self._name = new

            def _delete(self) -> None:
                coll.deleted.append(self._name)
                coll._sheets = [s for s in coll._sheets if s is not self]
                coll._by_name.pop(self._name, None)

        new_ws = _NewSheet()
        self._sheets.append(new_ws)
        self._by_name[new_ws.Name] = new_ws
        self.added.append(new_ws)
        return new_ws


def test_apply_table_view_com_snapshot_writes_values_honors_limit_offset(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$C$1",
        column_names=["Name", "Qty", "Tag"],
        sheet_start_col=1,
        filter_mode=True,
        hidden_sheet_cols={3},
    )
    _attach_list_column_bodies(
        lo,
        {
            "Name": ["a", "b", "c", "d"],
            "Qty": [4, 1, 3, 2],
            "Tag": ["x", "y", "x", "x"],
        },
    )
    ws.Name = "software components"
    ws.FilterMode = True
    prior_hidden = dict(ws._col_state)

    sheets = _SnapshotSheets(ws)
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name", "Qty"],
        "where": [{"column": "Tag", "op": "eq", "value": "x"}],
        "sort": {"by": [{"column": "Qty", "order": "asc"}]},
        "limit": 2,
        "offset": 1,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        data = json.loads(
            svc.apply_table_view(
                book_path,
                view_spec,
                mode="snapshot",
                view_applicability={
                    "viewable": False,
                    "reason": "limit truncates matching rows",
                },
            )
        )

    assert data["mode"] == "snapshot"
    assert data["sheet"] == "mcp_view"
    token = data["restore_token"]
    assert token == {
        "v": 1,
        "kind": "snapshot",
        "sheet": "mcp_view",
        "created_by_tool": True,
    }
    assert len(sheets.added) == 1
    written = sheets.added[0]._value_store["v"]
    # Matching Tag=x sorted by Qty asc: (d,2), (c,3), (a,4) → offset 1 limit 2 → c,a
    assert written == (("Name", "Qty"), ("c", 3), ("a", 4))

    lo.Range.AutoFilter.assert_not_called()
    lo.AutoFilter.ShowAllData.assert_not_called()
    lo.Sort.Apply.assert_not_called()
    ws.Copy.assert_not_called()
    ws.ListObjects.Add.assert_not_called()
    assert ws.FilterMode is True
    assert dict(ws._col_state) == prior_hidden
    sheets.added[0].ListObjects.Add.assert_not_called()
    sheets.added[0].Copy.assert_not_called()


def test_apply_table_view_com_snapshot_unique_name_and_clear_deletes_sheet(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$A$1",
        column_names=["Name"],
        sheet_start_col=1,
    )
    _attach_list_column_bodies(lo, {"Name": ["a", "b"]})
    ws.Name = "Sheet1"
    ws.FilterMode = False

    sheets = _SnapshotSheets(ws)
    # Seed a colliding name with no ListObjects so find still hits Items.
    seed = MagicMock()
    seed.Name = "mcp_view"
    seed.ListObjects = MagicMock()
    seed.ListObjects.Count = 0
    sheets._sheets.insert(0, seed)
    sheets._by_name["mcp_view"] = seed

    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name"],
        "where": [],
        "sort": None,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        first = json.loads(
            svc.apply_table_view(book_path, view_spec, mode="snapshot")
        )
        assert first["sheet"] == "mcp_view_2"
        second = json.loads(
            svc.apply_table_view(book_path, view_spec, mode="snapshot")
        )
        assert second["sheet"] == "mcp_view_3"
        clear = json.loads(
            svc.clear_table_view(book_path, first["restore_token"])
        )

    assert clear["sheet"] == "mcp_view_2"
    assert "mcp_view_2" in sheets.deleted
    assert "mcp_view_3" not in sheets.deleted
    assert "mcp_view" not in sheets.deleted
    assert "Sheet1" not in sheets.deleted
    assert ws.FilterMode is False
    lo.Range.AutoFilter.assert_not_called()
    ws.Copy.assert_not_called()


def test_apply_table_view_com_snapshot_deletes_sheet_when_rename_fails(book_path):
    lo, ws = _view_list_object_mock(
        name="Items",
        header_addr="$A$1:$A$1",
        column_names=["Name"],
        sheet_start_col=1,
    )
    _attach_list_column_bodies(lo, {"Name": ["a"]})
    ws.Name = "Sheet1"
    sheets = _SnapshotSheets(ws)
    sheets.fail_next_rename = True
    wb = _workbook_mock(book_path, sheets)
    xl = MagicMock()
    xl.Workbooks = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)
    view_spec = {
        "target": {"kind": "table", "name": "Items"},
        "columns": ["Name"],
        "where": [],
        "sort": None,
    }
    with patch.dict(sys.modules, _fake_win32_modules(xl), clear=False):
        svc = ComWorkbookService(ImmediateExecutor())
        raw = svc.apply_table_view(book_path, view_spec, mode="snapshot")

    assert raw.startswith("Error:")
    assert sheets.deleted
    assert ws.Name == "Sheet1"
    lo.Range.AutoFilter.assert_not_called()

