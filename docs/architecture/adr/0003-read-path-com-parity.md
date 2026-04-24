# ADR 0003: Read-path COM parity (read tools when workbook is open in Excel)

## Status

Accepted

## Context

Write-class tools benefit most from COM routing: the user sees live grid updates and avoids file-vs-Excel divergence. Read-class tools (`read_data_from_excel`, `get_workbook_metadata`, validation reads, etc.) stay **file-backed** for simplicity and performance: openpyxl reads the on-disk snapshot.

If the workbook is open in Excel, the **on-disk file** may lag the in-memory Excel state (unsaved changes, recalc), so file-first reads can disagree with what the user sees. Alternatives considered included COM reads, auto-save-before-read, and an explicit **save** tool.

## Decision

1. **Read tools (Phase 1 and ongoing default):** All read-class tools remain **file-backed** only. No COM read path in the first delivery; no implicit save before read.

2. **Explicit `save_workbook` tool (new MCP tool):** Add a dedicated save operation so the **agent** can persist when it chooses—especially when using **COM** with `save_after_write=false` (FR-8): after mutating via COM, the agent can call `save_workbook` before `read_data_from_excel` so file reads reflect what Excel has flushed to disk.
   - **COM path:** `Workbook.Save` (or `SaveAs` when never saved); same routing and threading as other COM writes; clear errors on read-only, Protected View, etc. (FR-9).
   - **File path:** Today’s stack already saves after most mutations; `save_workbook` still has a clear contract (“persist now”) for agents and for a future session/dirty-buffer model. Until then, the file implementation may be a safe **reload–save** or documented **no-op when already persisted**—exact behavior is an implementation detail documented on the tool.
   - **Parameters:** Same path resolution, allowlist, and optional `workbook_transport` as other workbook tools.

3. **Phase 2 (optional, later):** Introduce **opt-in COM-based reads** (per-tool flag or env) if needed after write routing and `save_workbook` are proven in use. Document agent-facing caveats in README when added.

## Consequences

- Fewer COM code paths in v1; read tools stay simple and fast.
- README must state that **reads reflect on-disk file state**; agents that use COM writes without per-write save should call **`save_workbook` before reads** when they need file reads to match Excel.
- `RoutingBackend` keeps `tool_kind` / `operation_kind` so COM reads can be added later without structural redesign.
- **Inventory:** Register the new tool in `server.py`, `manifest.json`, and `TOOLS.md` when implemented (PRD traceability for tool list).

## Out of scope (unchanged)

- Implicit “save before every read” or “save or discard” automation (no single Excel COM primitive for that; discard does not refresh disk for file reads).
