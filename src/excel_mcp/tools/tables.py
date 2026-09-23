from __future__ import annotations

import logging
from typing import Callable, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.exceptions import (
    ChartError,
    DataError,
    PivotError,
    ValidationError,
    WorkbookError,
)
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)

logger = logging.getLogger("excel-mcp")


def register_tables(
    mcp: FastMCP,
    *,
    file_service: FileWorkbookService,
    workbook_dispatch: Callable[..., str],
    com_dispatch: Callable[..., Callable[[str], str] | None],
) -> Dict[str, Callable]:
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Table",
            destructiveHint=True,
        ),
    )
    def create_table(
        filepath: str,
        sheet_name: str,
        data_range: str,
        table_name: Optional[str] = None,
        table_style: str = "TableStyleMedium9",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Creates a native Excel table from a specified range of data."""
        try:
            return workbook_dispatch(
                "create_table",
                filepath,
                workbook_transport,
                lambda fp: file_service.create_excel_table(
                    fp,
                    sheet_name,
                    data_range,
                    table_name=table_name,
                    table_style=table_style,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.create_excel_table(
                        fp,
                        sheet_name,
                        data_range,
                        table_name=table_name,
                        table_style=table_style,
                    )
                ),
            )
        except DataError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error creating table: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Chart",
            destructiveHint=True,
        ),
    )
    def create_chart(
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        title: str = "",
        x_axis: str = "",
        y_axis: str = "",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Create chart in worksheet."""
        try:
            return workbook_dispatch(
                "create_chart",
                filepath,
                workbook_transport,
                lambda fp: file_service.create_chart_in_sheet(
                    fp,
                    sheet_name,
                    data_range,
                    chart_type,
                    target_cell,
                    title=title,
                    x_axis=x_axis,
                    y_axis=y_axis,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.create_chart_in_sheet(
                        fp,
                        sheet_name,
                        data_range,
                        chart_type,
                        target_cell,
                        title=title,
                        x_axis=x_axis,
                        y_axis=y_axis,
                    )
                ),
            )
        except (ValidationError, ChartError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error creating chart: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Pivot Table",
            destructiveHint=True,
        ),
    )
    def create_pivot_table(
        filepath: str,
        sheet_name: str,
        data_range: str,
        rows: List[str],
        values: List[str],
        columns: Optional[List[str]] = None,
        agg_func: str = "mean",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Create pivot table in worksheet."""
        try:
            return workbook_dispatch(
                "create_pivot_table",
                filepath,
                workbook_transport,
                lambda fp: file_service.create_pivot_table_in_sheet(
                    fp,
                    sheet_name,
                    data_range,
                    rows,
                    values,
                    columns=columns,
                    agg_func=agg_func,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.create_pivot_table_in_sheet(
                        fp,
                        sheet_name,
                        data_range,
                        rows,
                        values,
                        columns=columns,
                        agg_func=agg_func,
                    )
                ),
            )
        except (ValidationError, PivotError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error creating pivot table: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Tables",
            readOnlyHint=True,
        ),
    )
    def list_tables(
        filepath: str,
        detail: str = "schema",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """List native Excel tables (ListObjects) in a workbook.

        Returns JSON ``{\"tables\": [...]}`` with sheet, name, range, header_range,
        data_range, row_count, and filter_applied. ``detail=schema`` (default) includes
        column names; ``detail=minimal`` omits them. No cell values. Does not change
        filters, sort order, or hidden columns.
        """
        try:
            return workbook_dispatch(
                "list_tables",
                filepath,
                workbook_transport,
                lambda fp: file_service.list_tables(fp, detail=detail),
                com_do_op=com_dispatch(
                    lambda c, fp: c.list_tables(fp, detail=detail)
                ),
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error listing tables: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Query Table",
            readOnlyHint=True,
        ),
    )
    def query_table(
        filepath: str,
        table: str,
        columns: Optional[list[str]] = None,
        where: Optional[list[dict]] = None,
        limit: Optional[int] = None,
        offset: Optional[int] = None,
        omit_empty: bool = False,
        search: Optional[str] = None,
        workbook_transport: Optional[str] = None,
        include_routing_metadata: bool = False,
    ) -> str:
        """Query rows from a native Excel table (ListObject) by column and filter.

        Args:
            filepath: Workbook path or cloud locator (see TOOLS.md filepath section).
            table: ListObject name (unique in the workbook).
            columns: Projected column names. Required when the table has more than 32
                columns unless ``omit_empty`` is true (then up to 128); otherwise
                defaults to all columns.
            where: AND-combined clauses ``{column, op, value}``. Operators: eq, neq,
                contains, in, gt, gte, lt, lte, is_empty. Empty list returns all data
                rows (subject to limit/offset).
            limit: Max rows to return (default 100). Marked not viewable.
            offset: Rows to skip from matching set (default 0). Marked not viewable.
            omit_empty: When true, drop projected columns that are empty (None or "")
                for every returned row; also allows omitting ``columns`` on tables
                with up to 128 columns.
            search: Optional case-insensitive substring matched against any loaded
                string cell; AND-combined with ``where``.
            workbook_transport: Optional ``auto`` | ``file`` | ``com``.
            include_routing_metadata: When true, wrap JSON in ADR 0010 envelope.

        Returns:
            JSON with ``headers``, row objects, ``row_count``, ``truncated``,
            ``view_spec``, and ``view_applicability``. Does not change filters, sort,
            or hidden columns. Never calls AutoFilter/Sort.
        """
        try:
            return workbook_dispatch(
                "query_table",
                filepath,
                workbook_transport,
                lambda fp: file_service.query_table(
                    fp,
                    table,
                    columns=columns,
                    where=where,
                    limit=limit,
                    offset=offset,
                    omit_empty=omit_empty,
                    search=search,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.query_table(
                        fp,
                        table,
                        columns=columns,
                        where=where,
                        limit=limit,
                        offset=offset,
                        omit_empty=omit_empty,
                        search=search,
                    )
                ),
                include_routing_metadata=include_routing_metadata,
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error querying table: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Map Sheet Layout",
            readOnlyHint=True,
        ),
    )
    def map_sheet_layout(
        filepath: str,
        sheet_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Map native Excel tables and non-table occupied islands on one sheet.

        Returns JSON ``{\"tables\": [...], \"regions\": [...]}``. Tables reuse the
        ``list_tables`` fields plus ``kind: \"table\"``. Regions include ``id``,
        bounds, and a header guess. Does not create ListObjects or change filters,
        sort, or hidden columns.
        """
        try:
            return workbook_dispatch(
                "map_sheet_layout",
                filepath,
                workbook_transport,
                lambda fp: file_service.map_sheet_layout(fp, sheet_name),
                com_do_op=com_dispatch(
                    lambda c, fp: c.map_sheet_layout(fp, sheet_name)
                ),
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error mapping sheet layout: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Query Region",
            readOnlyHint=True,
        ),
    )
    def query_region(
        filepath: str,
        sheet_name: str,
        region_id: Optional[str] = None,
        range: Optional[str] = None,
        columns: Optional[list[str]] = None,
        where: Optional[list[dict]] = None,
        limit: Optional[int] = None,
        offset: Optional[int] = None,
        omit_empty: bool = False,
        search: Optional[str] = None,
        workbook_transport: Optional[str] = None,
        include_routing_metadata: bool = False,
    ) -> str:
        """Query rows from a non-table rectangular region by id or explicit A1 range.

        Args:
            filepath: Workbook path or cloud locator (see TOOLS.md filepath section).
            sheet_name: Worksheet name.
            region_id: Stable id from ``map_sheet_layout`` (``sheet!A1:D10``). Exactly
                one of ``region_id`` or ``range`` is required.
            range: Explicit A1 range whose first row is the header.
            columns: Projected column names. Required when the region has more than 32
                columns unless ``omit_empty`` is true (then up to 128); otherwise
                defaults to all columns.
            where: AND-combined clauses ``{column, op, value}``. Same operators as
                ``query_table``.
            limit: Max rows to return (default 100). Marked not viewable.
            offset: Rows to skip from matching set (default 0). Marked not viewable.
            omit_empty: Same as ``query_table``: drop all-empty projected columns.
            search: Same as ``query_table``: substring across loaded string cells.
            workbook_transport: Optional ``auto`` | ``file`` | ``com``.
            include_routing_metadata: When true, wrap JSON in ADR 0010 envelope.

        Returns:
            JSON with ``headers``, row objects, ``row_count``, ``truncated``,
            ``view_spec`` (target kind ``region``), and ``view_applicability``. Does
            not create a ListObject or change filters, sort, or hidden columns.
        """
        try:
            return workbook_dispatch(
                "query_region",
                filepath,
                workbook_transport,
                lambda fp: file_service.query_region(
                    fp,
                    sheet_name,
                    region_id=region_id,
                    range=range,
                    columns=columns,
                    where=where,
                    limit=limit,
                    offset=offset,
                    omit_empty=omit_empty,
                    search=search,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.query_region(
                        fp,
                        sheet_name,
                        region_id=region_id,
                        range=range,
                        columns=columns,
                        where=where,
                        limit=limit,
                        offset=offset,
                        omit_empty=omit_empty,
                        search=search,
                    )
                ),
                include_routing_metadata=include_routing_metadata,
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error querying region: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Apply Table View",
            destructiveHint=True,
        ),
    )
    def apply_table_view(
        filepath: str,
        view_spec: dict,
        mode: str = "in_place",
        view_applicability: Optional[dict] = None,
        workbook_transport: Optional[str] = None,
        include_routing_metadata: bool = False,
    ) -> str:
        """Apply a ``view_spec`` as an Excel view (in_place or snapshot).

        COM-only WRITE. ``in_place`` applies AutoFilter / optional Sort / column
        focus on a ListObject or plain region (shared desktop view). ``snapshot``
        writes the projected, filtered, sorted result to a new values-only sheet and
        leaves the source unchanged. File transport returns an error without mutating
        the xlsx.

        Returns JSON including ``restore_token`` for ``clear_table_view``.
        """
        try:
            return workbook_dispatch(
                "apply_table_view",
                filepath,
                workbook_transport,
                lambda fp: file_service.apply_table_view(
                    fp,
                    view_spec,
                    mode=mode,
                    view_applicability=view_applicability,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.apply_table_view(
                        fp,
                        view_spec,
                        mode=mode,
                        view_applicability=view_applicability,
                    )
                ),
                include_routing_metadata=include_routing_metadata,
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error applying table view: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Clear Table View",
            destructiveHint=True,
        ),
    )
    def clear_table_view(
        filepath: str,
        restore_token: dict,
        workbook_transport: Optional[str] = None,
        include_routing_metadata: bool = False,
    ) -> str:
        """Clear a view using a ``restore_token`` from ``apply_table_view``.

        COM-only WRITE. ListObject tokens use ``ListObject.AutoFilter.ShowAllData()``
        (not ``Worksheet.AutoFilterMode``). Plain-range tokens restore prior
        ``AutoFilterMode``. Snapshot tokens delete only the sheet this tool created
        (``kind: snapshot`` with ``created_by_tool: true``). Unhides only columns
        this tool hid for in-place tokens. File transport returns an error.
        """
        try:
            return workbook_dispatch(
                "clear_table_view",
                filepath,
                workbook_transport,
                lambda fp: file_service.clear_table_view(fp, restore_token),
                com_do_op=com_dispatch(
                    lambda c, fp: c.clear_table_view(fp, restore_token)
                ),
                include_routing_metadata=include_routing_metadata,
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error clearing table view: {e}")
            raise


    return {
        "create_table": create_table,
        "create_chart": create_chart,
        "create_pivot_table": create_pivot_table,
        "list_tables": list_tables,
        "query_table": query_table,
        "map_sheet_layout": map_sheet_layout,
        "query_region": query_region,
        "apply_table_view": apply_table_view,
        "clear_table_view": clear_table_view,
    }
