"""ComWorkbookService: executor wiring and COM workers without Excel (Story 7-1)."""

from __future__ import annotations

import json
import os
import sys
import types
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
