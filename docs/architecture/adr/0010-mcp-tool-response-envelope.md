# ADR 0010: MCP tool response envelope (routing metadata + warnings)

## Status

Accepted

## Context

Routed workbook tools return a plain **`str`** to MCP clients: either a **tool-specific JSON body** (reads, metadata) or a **human-readable success message**, and **`"Error: …"`** for many caught domain failures ([pre-fork architecture](../pre-fork-architecture.md) § MCP tool surface).

Routing observability fields — **`workbook_transport`**, **`workbook_backend`**, **`routing_reason`**, **`duration_ms`** — are emitted today only as a single structured **log line** on logger `excel-mcp.routing` inside [`execute_routed_workbook_operation`](../../../src/excel_mcp/routing/routed_dispatch.py) (NFR-3, [target architecture](../target-architecture.md) §8). Agents and integrators **cannot** verify which backend executed a call without log access.

[`_workbook_dispatch`](../../../src/excel_mcp/server.py) already invokes routed dispatch but **discards** the executed backend tuple element (`out, _backend = …` at ~L109–121), so handler-level attachment is the natural integration point.

Planned capabilities share one response pattern:

- **Routing `_meta` on tool responses** (e.g. expose executed backend to agents).
- **Non-fatal warnings** (e.g. file-backend formula evaluation limits on `.xlsm`).
- **Optional compact read payloads** (smaller per-cell metadata on large ranges).

[`WorkbookOperationMetadata`](../../../src/excel_mcp/routing/workbook_operation_contract.py) was introduced for routing hints but is **not yet** wired into MCP wire responses. This ADR defines the **MCP-facing envelope** without requiring every follow-up story to invent a new shape.

## Decision

### 1. Backward compatibility — opt-in `include_routing_metadata`

Add an optional tool parameter on **routed** handlers (and any tool that adopts the envelope):

| Parameter | Type | Default | Semantics |
|-----------|------|---------|-----------|
| `include_routing_metadata` | `bool` | **`false`** | When `false`, MCP tool result text is **unchanged** from today. When `true`, eligible successes use the envelope in §2. |

**Compatibility rule for existing agents:** Clients that `json.loads` the entire tool result and expect a **legacy top-level schema** (e.g. `range`, `sheet_name`, `cells` on reads) **must** either leave `include_routing_metadata` at default **`false`**, or update parsers to accept the envelope (§2) and read the nested **`result`** (or documented merge rules). **No routed tool may change its default wire shape** without a new ADR and SemVer note (§5).

Non-routed tools (e.g. pure validation helpers with no workbook routing) **may** omit the parameter until they need `_meta` or `warnings`.

### 2. Stable field names — `_meta` object

When `include_routing_metadata=true` and the call is a **successful envelope response** (§4), attach a top-level **`_meta`** object using **the same names as NFR-3 / ADR 0001 routing logs**:

| Field | Type | Source |
|-------|------|--------|
| `workbook_transport` | `string` | Requested transport (`auto` \| `file` \| `com`) |
| `workbook_backend` | `string` | Executed backend (`file` \| `com`) |
| `routing_reason` | `string` | Resolution reason from `RoutingBackend` |
| `duration_ms` | `number` | Wall time for routed dispatch (milliseconds, rounded) |

Optional keys (same as logs) may appear when applicable: `mcp_tool_name`, `v1_file_forced`. **Do not** rename these for MCP; log and wire stay aligned.

**Success envelope shape (JSON tools):** Parse the backend `result_text` when it is JSON; return a single JSON object:

```json
{
  "result": { },
  "_meta": {
    "workbook_transport": "auto",
    "workbook_backend": "com",
    "routing_reason": "com_workbook_open",
    "duration_ms": 12.345
  },
  "warnings": []
}
```

- **`result`:** Parsed tool payload (object or array). For tools that today return **non-JSON** success strings, `result` is that string (JSON-encoded as a string value).
- **`warnings`:** Array per §3 (may be empty).
- **`_meta`:** Always present when `include_routing_metadata=true` on success.

Follow-up stories (e.g. compact read mode) add **optional request parameters** and may add **optional `_meta` or `result` fields** documented in TOOLS.md; they must not break the default-off wire shape.

### 3. Warnings array — non-fatal signals

Non-fatal operator or agent signals use a top-level **`warnings`** array (only when the envelope is used). Each entry:

```json
{ "code": "file_backend_formula_not_evaluated", "message": "…" }
```

| Field | Required | Notes |
|-------|----------|-------|
| `code` | yes | Stable machine identifier (snake_case); document in TOOLS.md |
| `message` | yes | Human-readable explanation |

Warnings **do not** flip MCP tool success to failure. Example: file backend read of `.xlsm` where cached values are returned but formulas are not evaluated (warning code `file_backend_formula_not_evaluated`).

New warning codes are **additive** (patch SemVer) when `include_routing_metadata` remains opt-in.

### 4. Error vs success policy

| Outcome | `include_routing_metadata=false` (default) | `include_routing_metadata=true` |
|---------|---------------------------------------------|----------------------------------|
| **Success** | Unchanged: legacy JSON string or plain message | JSON envelope per §2 (`result`, `_meta`, `warnings`) |
| **Failure** | Unchanged: plain **`"Error: …"`** string (prefix convention) | **Same:** plain **`"Error: …"`** string — **no** JSON envelope on failure |

Rationale:

- Existing agents detect failure with **`str.startswith("Error:")`** or similar; wrapping errors in JSON would break them even when opt-in.
- Routing metadata for failed operations remains available in **`excel-mcp.routing` logs** (and may be extended in logs later); MCP wire stays error-compatible.

**Special case:** [`execute_routed_workbook_operation`](../../../src/excel_mcp/routing/routed_dispatch.py) may return a user-visible **`"Error: …"`** string inside `result_text` for some routing guardrails (e.g. file backend + HTTPS locator) **without** raising. That is still a **successful MCP tool invocation** with an error-shaped **payload**. Envelope opt-in wraps it as `"result": "Error: …"` plus `_meta`, not as an MCP-level failure string.

**Raises** (`ComExecutionNotImplementedError`, validation exceptions caught in handlers, etc.) continue to map to handler-level **`"Error: …"`** returns per existing patterns.

### 5. SemVer and release notes

Per [release-versioning-policy.md](../release-versioning-policy.md):

| Change | Typical bump (pre-1.0) | CHANGELOG section |
|--------|-------------------------|-------------------|
| **This ADR only** (documentation) | none | Docs |
| **Implement** opt-in `include_routing_metadata` + envelope on routed tools | **minor** (`0.(y+1).0`) — new optional capability | Added |
| **New `warnings` codes** with envelope opt-in | **patch** — additive, opt-in | Added or Changed |
| **Default `include_routing_metadata` → `true`** | **minor** with **Breaking / impact** — legacy JSON parsers break | Changed + Breaking / impact |
| **Rename or remove `_meta` / warning fields** | **minor** with **Breaking / impact** | Changed + Breaking / impact |

Document user-visible envelope behavior in **TOOLS.md** and the release **Breaking / impact** block when defaults or required response shapes change.

## Implementation sketch

Attachment point (no implementation in this ADR):

1. **[`routed_dispatch.py`](../../../src/excel_mcp/routing/routed_dispatch.py)** — `execute_routed_workbook_operation` already returns `(result_text, executed_backend)` and logs `_meta` fields. Extend or add a small helper (e.g. `build_routed_response_envelope`) that accepts `result_text`, resolution metadata, `duration_ms`, and `warnings`, and returns serialized JSON when envelope is requested.

2. **[`server.py`](../../../src/excel_mcp/server.py) `_workbook_dispatch`** — Today:

   ```python
   out, _backend = execute_routed_workbook_operation(...)
   return out
   ```

   Wire `include_routing_metadata` from tool handlers into dispatch; when `true`, replace `return out` with envelope serialization using dispatch metadata (transport, backend, reason, timing). Keep `return out` when `false`.

3. **Handlers** — Add optional `include_routing_metadata: bool = False` to routed `@mcp.tool` schemas; pass through `_workbook_dispatch` (or a shared wrapper). JSON read/write tools are the first consumers; plain-string tools follow the same envelope rules (§2).

4. **`WorkbookOperationMetadata`** — May feed handler/dispatch context (`mcp_tool_name`, `tool_kind`); MCP wire `_meta` uses the stable NFR-3 field set in §2, not necessarily every `TypedDict` key.

Logging **unchanged**: continue emitting one JSON log line per routed operation regardless of envelope opt-in.

## Consequences

- **BEN-120 / BEN-134 / BEN-136** (and similar) implement against this contract instead of one-off response shapes.
- **Default-off** preserves all existing agents that parse plain JSON tool bodies.
- **Opt-in** agents gain routing transparency and warnings without log access.
- **Tests:** Contract tests for envelope on/off, warning codes, and `"Error: …"` failure path unchanged; integration tests via `_workbook_dispatch` mocks.
- **Documentation:** TOOLS.md and README *Routing observability* cross-link this ADR when implementation lands.

## Links

- [ADR 0001 — Workbook transport vs MCP wire transport](0001-workbook-transport-vs-mcp-wire-transport.md)
- [ADR 0008 — COM-first default routing](0008-com-first-default-and-file-lifecycle-tools.md)
- [Release versioning policy](../release-versioning-policy.md)
- [Target architecture](../target-architecture.md) §8 Observability
- [`routed_dispatch.py`](../../../src/excel_mcp/routing/routed_dispatch.py) — `execute_routed_workbook_operation`
- [`server.py`](../../../src/excel_mcp/server.py) — `_workbook_dispatch`
