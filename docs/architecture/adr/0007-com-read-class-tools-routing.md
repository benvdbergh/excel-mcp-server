# ADR 0007: COM routing and execution for read-class MCP tools (draft)

## Status

Superseded by [ADR 0008 — COM-first default and file lifecycle tools](0008-com-first-default-and-file-lifecycle-tools.md)

## Context

[ADR 0003](0003-read-path-com-parity.md) accepted **file-backed reads** for all read-class tools, with an optional later phase for **COM-based reads**. The codebase reflects that decision:

- [`RoutingBackend.resolve_workbook_backend`](../../../src/excel_mcp/routing/routing_backend.py) returns **`backend="file"`** with reason **`read_class_file_backed`** for every `ToolKind.READ` invocation, regardless of `workbook_transport`.
- MCP handlers for read tools do not pass **`com_do_op`** into [`_workbook_dispatch`](../../../src/excel_mcp/server.py); [`ComWorkbookService`](../../../src/excel_mcp/routing/com_workbook_service.py) **read** methods are **stubs**.

Operators increasingly need:

1. **Cloud workbook locators** ([ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md)): HTTPS `filepath` values cannot use the openpyxl backend; **COM reads** are the practical way to read data when Excel hosts the workbook.
2. **Consistency with the live grid** when the workbook is open in Excel: file reads may **diverge** from on-screen state until **`save_workbook`** is used.

The design note [COM read-class tools: design note](../com-read-class-tools-design.md) analyzes routing options, `ComWorkbookService` parity with [`FileWorkbookService`](../../../src/excel_mcp/routing/file_workbook_service.py), risks, and backward compatibility.

## Decision (historical — superseded)

The **file-first + opt-in COM reads** direction below was **not** adopted. See **[ADR 0008](0008-com-first-default-and-file-lifecycle-tools.md)** for **COM-first default reads/writes**, explicit lifecycle tools, and **`save_after_write` removal**.

~~**Candidate direction:**~~

1. ~~**Default remains file-backed reads** for backward compatibility~~ — superseded.
2. ~~**Explicit opt-in** for COM-backed reads~~ — superseded by COM-first routing ([ADR 0008](0008-com-first-default-and-file-lifecycle-tools.md)).
3. ~~Extend `resolve_workbook_backend`~~ — superseded; READ participates in shared COM-first rules.
4. **`com_do_op` + `ComWorkbookService`** — still required technically; tracked under ADR 0008 and [com-read-class-tools-design.md](../com-read-class-tools-design.md).

## Consequences (historical)

- Preserved as rationale for why **READ** was forced to **file** in early routing docs (`read_class_file_backed`).
- Superseded ADR: [0008](0008-com-first-default-and-file-lifecycle-tools.md).

## Links

- **Current:** [ADR 0008 — COM-first default and file lifecycle](0008-com-first-default-and-file-lifecycle-tools.md) · [COM-first workbook session design](../com-first-workbook-session-design.md)
- [ADR 0003 — Read-path COM parity](0003-read-path-com-parity.md)
- [ADR 0006 — Cloud workbook locators](0006-cloud-workbook-locator-sharepoint-urls.md)
- [Design note: COM read-class tools](../com-read-class-tools-design.md)
