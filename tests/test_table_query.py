"""Shared ``query_table`` / region predicate and layout helpers (STORY-13-2/13-3)."""

import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.exceptions import DataError  # noqa: E402
from excel_mcp.tables import (  # noqa: E402
    MAX_QUERY_TABLE_COLUMNS_WITHOUT_EXPLICIT,
    evaluate_row_clause,
    filter_table_rows,
    guess_header_from_row,
    layout_regions_from_grid,
    query_region_rows,
    query_table_rows,
)


def test_numeric_coercion_eq_and_compare() -> None:
    row = {"N": "10", "T": "abc"}
    assert evaluate_row_clause(row, {"column": "N", "op": "eq", "value": 10})
    assert evaluate_row_clause(row, {"column": "N", "op": "gte", "value": 10})
    assert evaluate_row_clause(row, {"column": "N", "op": "lt", "value": 11})
    assert not evaluate_row_clause(row, {"column": "T", "op": "eq", "value": 10})


def test_contains_case_insensitive_and_in_and_is_empty() -> None:
    row = {"A": "Hello World", "B": None, "C": "  "}
    assert evaluate_row_clause(
        row, {"column": "A", "op": "contains", "value": "hello"}
    )
    assert evaluate_row_clause(
        row, {"column": "A", "op": "in", "value": ["x", "Hello World"]}
    )
    assert evaluate_row_clause(row, {"column": "B", "op": "is_empty"})
    assert evaluate_row_clause(row, {"column": "C", "op": "is_empty"})


def test_query_table_rows_filter_page_and_viewability() -> None:
    cols = ["Name", "Qty", "Tag"]
    rows = [
        {"Name": "a", "Qty": 1, "Tag": "x"},
        {"Name": "b", "Qty": 2, "Tag": "y"},
        {"Name": "c", "Qty": 3, "Tag": "x"},
        {"Name": "d", "Qty": 4, "Tag": "z"},
    ]
    out = query_table_rows(
        table_name="T1",
        table_columns=cols,
        rows=rows,
        columns=["Name", "Qty"],
        where=[{"column": "Tag", "op": "in", "value": ["x", "y"]}],
        limit=2,
        offset=0,
    )
    assert out["headers"] == ["Name", "Qty"]
    assert out["row_count"] == 2
    assert out["truncated"] is True
    assert out["rows"] == [
        {"Name": "a", "Qty": 1},
        {"Name": "b", "Qty": 2},
    ]
    assert out["view_spec"]["target"] == {"kind": "table", "name": "T1"}
    assert out["view_spec"]["where"] == [
        {"column": "Tag", "op": "in", "value": ["x", "y"]}
    ]
    assert out["view_spec"]["sort"] is None
    va = out["view_applicability"]
    assert va["limit_viewable"] is False
    assert va["offset_viewable"] is False
    assert va["clauses"][0]["viewable"] is True
    assert va["viewable"] is False  # truncated by limit


def test_query_table_rows_full_page_keeps_viewable_true() -> None:
    cols = ["A"]
    rows = [{"A": 1}, {"A": 2}]
    out = query_table_rows(
        table_name="T1",
        table_columns=cols,
        rows=rows,
        where=[{"column": "A", "op": "eq", "value": 1}],
        limit=100,
        offset=0,
    )
    assert out["truncated"] is False
    assert out["row_count"] == 1
    assert out["view_applicability"]["viewable"] is True


def test_is_empty_clause_not_viewable() -> None:
    out = query_table_rows(
        table_name="T1",
        table_columns=["A"],
        rows=[{"A": None}, {"A": 1}],
        where=[{"column": "A", "op": "is_empty"}],
    )
    assert out["row_count"] == 1
    assert out["view_applicability"]["viewable"] is False
    assert out["view_applicability"]["clauses"][0]["viewable"] is False
    # Non-viewable clauses stay in the spec so a later view can refuse them
    # instead of showing every row.
    assert out["view_spec"]["where"] == [{"column": "A", "op": "is_empty"}]


def test_wide_table_requires_columns() -> None:
    wide = [f"C{i}" for i in range(MAX_QUERY_TABLE_COLUMNS_WITHOUT_EXPLICIT + 1)]
    with pytest.raises(DataError, match="explicit columns"):
        query_table_rows(
            table_name="Wide",
            table_columns=wide,
            rows=[{c: 1 for c in wide}],
            columns=None,
        )


def test_unknown_column_and_invalid_op() -> None:
    with pytest.raises(DataError, match="Unknown column"):
        query_table_rows(
            table_name="T1",
            table_columns=["A"],
            rows=[{"A": 1}],
            columns=["B"],
        )
    with pytest.raises(DataError, match="Invalid where operator"):
        query_table_rows(
            table_name="T1",
            table_columns=["A"],
            rows=[{"A": 1}],
            where=[{"column": "A", "op": "regex", "value": "x"}],
        )


def test_query_region_rows_reuses_predicate_and_region_target() -> None:
    cols = ["Name", "Tag"]
    rows = [
        {"Name": "a", "Tag": "x"},
        {"Name": "b", "Tag": "y"},
    ]
    # Same predicate path as query_table (filter_table_rows).
    matching = filter_table_rows(
        rows, [{"column": "Tag", "op": "eq", "value": "x"}]
    )
    assert matching == [{"Name": "a", "Tag": "x"}]
    out = query_region_rows(
        sheet="Sheet1",
        range_a1="B4:C6",
        region_columns=cols,
        rows=rows,
        where=[{"column": "Tag", "op": "eq", "value": "x"}],
    )
    assert out["rows"] == [{"Name": "a", "Tag": "x"}]
    assert out["view_spec"]["target"] == {
        "kind": "region",
        "sheet": "Sheet1",
        "range": "B4:C6",
    }


def test_guess_header_mostly_text_vs_numeric() -> None:
    ok, cols = guess_header_from_row(["Name", "Qty", "Tag"])
    assert ok is True
    assert cols == ["Name", "Qty", "Tag"]
    bad, empty = guess_header_from_row([10, 20, 30])
    assert bad is False
    assert empty == []


def test_layout_regions_exclude_table_and_skip_blank_row1() -> None:
    # Row 1 blank in used origin; header at sheet row 4 (origin_row=1 → index 3).
    values = [
        [None, None, None],  # sheet row 1
        [None, None, None],  # row 2
        [None, None, None],  # row 3
        ["Name", "Qty", None],  # row 4 header
        ["a", 1, None],
        ["b", 2, None],
        [None, None, None],  # blank separator
        ["X", "Y", None],  # second island inside a table range → excluded
        [1, 2, None],
    ]
    regions = layout_regions_from_grid(
        sheet="stackfleet concept",
        values=values,
        origin_row=1,
        origin_col=1,
        excluded_ranges=["A8:B9"],
    )
    assert len(regions) == 1
    r0 = regions[0]
    assert r0["kind"] == "region"
    assert r0["header_row"] == 4
    assert r0["header_guess"] is True
    assert r0["columns"] == ["Name", "Qty"]
    assert r0["range"] == "A4:B6"
    assert r0["id"] == "stackfleet concept!A4:B6"
    assert r0["header_range"] == "A4:B4"


def test_layout_regions_do_not_cover_table_cells() -> None:
    """A ListObject inside the used range must not be part of an island bbox."""
    # A1:B2 values, C1:D2 is a table, A3:D3 values under both.
    values = [
        ["Name", "Qty", "T1", "T2"],
        ["a", 1, "t", "u"],
        ["b", 2, "c", "d"],
    ]
    regions = layout_regions_from_grid(
        sheet="Sheet1",
        values=values,
        origin_row=1,
        origin_col=1,
        excluded_ranges=["C1:D2"],
    )
    ranges = {r["range"] for r in regions}
    assert ranges == {"A1:B2", "A3:D3"}
    for region in regions:
        assert not region["range"].startswith("C1")
        assert "C1:D2" not in region["range"]
