# ADR 0008: COM-first default routing and explicit file lifecycle tools

## Status

Accepted

## Context

[ADR 0007](0007-com-read-class-tools-routing.md) (now superseded) assumed **default file-backed reads** with **opt-in COM reads** to preserve backward compatibility. There are **no production users yet**; the product can pivot without preserving that default.

Operators expect the MCP to mirror **Excel as the system of record** when Excel hosts the workbook: **live grid**, **cloud identities** ([ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md)), and **consistent read/write routing**. Today’s [RoutingBackend](../target-architecture.md) effectively implements **read → file always** ([ADR 0003](0003-read-path-com-parity.md) Phase 1) and **auto → file when workbook not open** for writes—a **file-first** bias that fights COM-first operations.

Separately, **`save_after_write`** (and env-driven defaults in [`routing_env.py`](../../../src/excel_mcp/routing/routing_env.py)) couples **persistence** to **every mutating tool**, which obscures when state is flushed to disk—especially when Excel holds dirty state. An explicit **save** tool exists ([ADR 0003](0003-read-path-com-parity.md) §2) but coexists with per-call save flags—duplicate control surfaces.

Session semantics are implicit (path resolution “opens” file-backed workbooks via openpyxl; COM matches **open** workbooks by identity). There is no first-class **open this path in Excel** or **close workbook in Excel** for agent-driven session management.

See [COM-first workbook session design](../com-first-workbook-session-design.md) for routing matrix, session model, SSE/jail, threading, security, and inventory notes.

## Decision

### 1. Default routing: COM first when viable; file/openpyxl fallback

For **both reads and writes** (same philosophy unless a documented exception applies):

1. **`transport="auto"`:** Prefer **`com`** when **COM runtime is viable** ([ADR 0002](0002-com-automation-stack.md): Windows + `[com]` + executor) **and** the target workbook **matches an open Excel workbook** (identity / `FullName` per [ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md) for HTTPS locators).
2. **Otherwise** use **`file`** (openpyxl / disk path): non-Windows, COM unavailable, workbook not open in Excel, headless CI, etc.
3. **`transport="com"`:** Still forces COM when viable (existing intent); failure modes per [ADR 0005](0005-com-strict-and-fallback-controls.md).
4. **`transport="file"`:** Always file—unchanged.

**Documented exceptions** (carry forward unless revised by a later ADR):

- **`ToolKind.V1_FILE_FORCED`** ([ADR 0004](0004-chart-pivot-com-parity-scope.md)): chart/pivot remain file-forced where applicable.
- **RoutingBackend / `resolve_workbook_backend`:** **Invert** the current READ branch: **`ToolKind.READ` participates in the same COM-first / fallback rules** as writes instead of unconditional `read_class_file_backed` → file ([ADR 0007](0007-com-read-class-tools-routing.md) superseded).
- **COM read implementation** ([ADR 0003](0003-read-path-com-parity.md) Phase 2 intent): **no separate “opt-in only”** product gate—COM reads are **on the default path** when routing selects COM. (Phase 2 “opt-in” language in ADR 0003 is **superseded** for defaults; the ADR remains historical for the explicit `save_workbook` decision.)

### 2. Explicit file lifecycle tools (MCP)

Introduce or sharpen **first-class** lifecycle operations. Exact names are TBD; roles:

| Role | Semantics (product) |
|------|---------------------|
| **File create** | Create a new workbook **on disk** at an allowed path. **Clarify vs `create_workbook`:** either alias/rename for clarity (`create_workbook` = file-first creation) or **deprecate** one name; optional **“open in Excel after create”** flag so the next **auto** operations target the COM session without a separate open call. |
| **File open (in Excel host)** | **Explicit** `Workbooks.Open` (or equivalent) so subsequent **COM-first** tools target that **Excel.Application** session. **Distinct from** implicit open: today, passing a path does not guarantee Excel has opened the file; **auto** only matches if already open. This tool **binds the host** to a path the operator chose. |
| **Save** | **Only** via dedicated **`save_workbook`** (or renamed equivalent). **Remove** `save_after_write` from **all** mutating tool signatures and from **env default** hooks (`EXCEL_MCP_SAVE_AFTER_WRITE` / `effective_save_after_write` pattern). Agents call save when they want persistence. |
| **File close** | Close the workbook in Excel (COM), with **`save: bool` (or enum)** for **save-then-close** vs **discard** (no save). Does not delete the file. |

[ADR 0001](0001-workbook-transport-vs-mcp-wire-transport.md): lifecycle tools affect **workbook transport** (host vs file), not MCP stdio/SSE **wire** transport.

### 3. Read-class COM execution

All **read-class** tools ([`tool_inventory.py`](../../../src/excel_mcp/routing/tool_inventory.py)) **may** execute on COM when routing selects COM—**same as writes**. Handlers must supply **both** `do_op` and `com_do_op` (see superseded [ADR 0007](0007-com-read-class-tools-routing.md) analysis); `ComWorkbookService` read methods must be real implementations, not stubs.

## Consequences

- **RoutingBackend** is the primary locus of change: remove READ short-circuit; apply COM-first resolution for `ToolKind.READ` analogously to write tools, subject to `V1_FILE_FORCED` and `com_strict` / fallback ([ADR 0005](0005-com-strict-and-fallback-controls.md)).
- **Cloud HTTPS locators** ([ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md)): COM-first default **aligns** with “Excel hosts the document”; file backend remains **invalid** for pure HTTPS locators—operators rely on COM or must use a local synced path with matching identity.
- **Breaking changes (acceptable):** default read/write behavior may **differ** from file snapshots when Excel has unsaved state; removal of `save_after_write` **breaks** existing agent prompts and tool schemas.
- **Documentation / manifest / TOOLS.md** must be updated in a follow-up implementation pass (out of scope for this ADR-only task).
- **Supersedes** [ADR 0007](0007-com-read-class-tools-routing.md) for product direction; **supersedes** [ADR 0003](0003-read-path-com-parity.md) **§1 default** (file-only reads) and **§3 Phase 2 opt-in** framing only—**§2 explicit `save_workbook`** remains the anchor for “explicit save,” extended here by **removing** per-mutation save parameters.

## Links

- Related: [ADR 0009 — Open workbook discovery tool](0009-open-workbook-discovery-tool.md) (host enumeration; **`get_workbook_metadata`** stays single-book)
- Supersedes: [ADR 0007 — COM read-class tools routing (draft)](0007-com-read-class-tools-routing.md)
- Partially supersedes: [ADR 0003 — Read-path COM parity](0003-read-path-com-parity.md) (defaults and Phase 2 opt-in only)
- [ADR 0006 — Cloud workbook locators](0006-cloud-workbook-locator-sharepoint-urls.md)
- [ADR 0002 — COM automation stack](0002-com-automation-stack.md)
- [Design: COM-first workbook session](../com-first-workbook-session-design.md)
- [Design: COM read-class tools (updated note)](../com-read-class-tools-design.md)
