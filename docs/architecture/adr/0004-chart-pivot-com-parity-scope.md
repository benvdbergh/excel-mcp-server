# ADR 0004: Chart and pivot COM parity scope (v1)

## Status

Accepted (v1 scope)

## Context

Pre-fork implementation uses **openpyxl** heavily:

- **Charts** — `chart.py` uses openpyxl chart APIs and internal drawing/chart hooks on worksheets. Mapping this 1:1 to Excel COM (`ChartObject` / `Chart`) is non-trivial.
- **Pivot** — `pivot.py` implements aggregation as **in-memory logic + a new sheet + Excel table**, not native Excel `PivotTable` / cache XML. COM “parity” is behavioral (similar sheet output), not identical Excel pivot objects.

Forcing these tools through COM in v1 without a full adapter risks regressions and long delivery time.

## Decision

**v1:** One of the following explicit policies (pick one at implementation time and keep README in sync):

1. **Tool-forced file** — `create_chart` and `create_pivot_table` always use **FileWorkbookService** regardless of `workbook_transport=auto`, until COM adapters exist; OR
2. **Route with degraded parity** — If COM is selected, implement minimal COM behavior (e.g. pivot as value grid + table) and document differences from openpyxl file output.

**Recommendation:** Policy **1** for v1 to preserve deterministic behavior and ship routing for high-value grid/formula writes first.

## Consequences

- PRD traceability: document in README / tool descriptions that chart/pivot may ignore `auto`→COM for v1 if policy 1 is chosen.
- `RoutingBackend` or tool layer documents exceptions to the default matrix.
- Future ADR or revision of this ADR when COM chart support lands.
