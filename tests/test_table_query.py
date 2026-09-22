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
    XL_FILTER_OR,
    XL_FILTER_VALUES,
    allocate_unique_sheet_name,
    build_snapshot_value_matrix,
    compile_autofilter_field_steps,
    criteria_text_for_clause,
    evaluate_row_clause,
    filter_table_rows,
    guess_header_from_row,
    layout_regions_from_grid,
    normalize_view_sort,
    query_region_rows,
    query_table_rows,
    sort_table_rows,
    validate_view_spec_for_apply,
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


def test_validate_view_spec_refuses_limit_offset_inplace_and_is_empty() -> None:
    base = {
        "target": {"kind": "table", "name": "T1"},
        "columns": ["A"],
        "where": [{"column": "A", "op": "eq", "value": 1}],
        "sort": None,
    }
    with pytest.raises(DataError, match="limit"):
        validate_view_spec_for_apply({**base, "limit": 10}, mode="in_place")
    with pytest.raises(DataError, match="offset"):
        validate_view_spec_for_apply({**base, "offset": 1}, mode="in_place")
    with pytest.raises(DataError, match="not viewable"):
        validate_view_spec_for_apply(
            {
                **base,
                "where": [{"column": "A", "op": "is_empty"}],
            },
            mode="in_place",
            table_columns=["A"],
        )
    with pytest.raises(DataError, match="not viewable"):
        validate_view_spec_for_apply(
            {
                **base,
                "where": [{"column": "A", "op": "is_empty"}],
            },
            mode="snapshot",
            table_columns=["A"],
        )
    with pytest.raises(DataError, match="limit truncates|viewable"):
        validate_view_spec_for_apply(
            base,
            mode="in_place",
            view_applicability={"viewable": False, "reason": "limit truncates"},
        )


def test_validate_view_spec_snapshot_honors_limit_offset() -> None:
    base = {
        "target": {"kind": "table", "name": "T1"},
        "columns": ["A"],
        "where": [{"column": "A", "op": "eq", "value": 1}],
        "sort": None,
    }
    out = validate_view_spec_for_apply(
        {**base, "limit": 2, "offset": 1},
        mode="snapshot",
        table_columns=["A"],
        view_applicability={
            "viewable": False,
            "reason": "limit truncates matching rows",
        },
    )
    assert out["mode"] == "snapshot"
    assert out["limit"] == 2
    assert out["offset"] == 1
    # Absent limit/offset → full result (no default page size of 100).
    full = validate_view_spec_for_apply(base, mode="snapshot", table_columns=["A"])
    assert full["limit"] is None
    assert full["offset"] == 0
    with pytest.raises(DataError, match="no limit or offset"):
        validate_view_spec_for_apply(
            base,
            mode="snapshot",
            table_columns=["A"],
            view_applicability={
                "viewable": False,
                "reason": "limit truncates matching rows",
            },
        )


def test_allocate_unique_sheet_name_and_snapshot_matrix() -> None:
    assert allocate_unique_sheet_name([]) == "mcp_view"
    assert allocate_unique_sheet_name(["mcp_view"]) == "mcp_view_2"
    assert allocate_unique_sheet_name(["MCP_VIEW", "mcp_view_2"]) == "mcp_view_3"
    rows = [
        {"A": 1, "B": "x"},
        {"A": 2, "B": "y"},
        {"A": 3, "B": "z"},
    ]
    sorted_rows = sort_table_rows(
        rows, {"by": [{"column": "A", "order": "desc"}]}
    )
    assert [r["A"] for r in sorted_rows] == [3, 2, 1]
    matrix = build_snapshot_value_matrix(
        headers=["A", "B"],
        matching_rows=sorted_rows,
        limit=2,
        offset=1,
    )
    assert matrix == [["A", "B"], [2, "y"], [1, "x"]]
    full = build_snapshot_value_matrix(
        headers=["A"], matching_rows=rows, limit=None, offset=0
    )
    assert full == [["A"], [1], [2], [3]]


def test_validate_view_spec_accepts_region_target() -> None:
    out = validate_view_spec_for_apply(
        {
            "target": {"kind": "region", "sheet": "S", "range": "$A$1:$B$2"},
            "columns": ["A"],
            "where": [{"column": "A", "op": "eq", "value": 1}],
            "sort": None,
        },
        mode="in_place",
        table_columns=["A", "B"],
    )
    assert out["target_kind"] == "region"
    assert out["target_sheet"] == "S"
    assert out["target_range"] == "A1:B2"
    assert out["where"][0]["op"] == "eq"


def test_validate_view_spec_refuses_cross_column_or() -> None:
    with pytest.raises(DataError, match="Cross-column OR|Nested OR"):
        validate_view_spec_for_apply(
            {
                "target": {"kind": "table", "name": "T1"},
                "columns": ["A", "B"],
                "where": [
                    {
                        "op": "or",
                        "clauses": [
                            {"column": "A", "op": "eq", "value": 1},
                            {"column": "B", "op": "eq", "value": 2},
                        ],
                    }
                ],
                "sort": None,
            },
            mode="in_place",
            table_columns=["A", "B"],
        )


def test_normalize_view_sort_and_autofilter_compile() -> None:
    sort = normalize_view_sort(
        {"by": [{"column": "Qty", "order": "desc"}]},
        known_columns=["Name", "Qty"],
    )
    assert sort == {"by": [{"column": "Qty", "order": "desc"}]}
    sugar = normalize_view_sort(
        {"column": "Name", "order": "asc"}, known_columns=["Name", "Qty"]
    )
    assert sugar == {"by": [{"column": "Name", "order": "asc"}]}

    two = criteria_text_for_clause(
        {"column": "Tag", "op": "in", "value": ["x", "y"]}
    )
    assert two["operator"] == XL_FILTER_OR
    assert two["criteria1"] == "x"
    assert two["criteria2"] == "y"
    many = criteria_text_for_clause(
        {"column": "Tag", "op": "in", "value": ["a", "b", "c"]}
    )
    assert many["operator"] == XL_FILTER_VALUES
    contains = criteria_text_for_clause(
        {"column": "Tag", "op": "contains", "value": "a*b?c~d"}
    )
    assert contains["criteria1"] == "*a~*b~?c~~d*"

    # Field index is table-relative (Name=1 even if sheet column is B).
    steps = compile_autofilter_field_steps(
        [{"column": "Qty", "op": "gt", "value": 1}],
        column_to_field={"Name": 1, "Qty": 2},
    )
    assert steps == [{"field": 2, "column": "Qty", "criteria1": ">1"}]
