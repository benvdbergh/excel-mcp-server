# Target architecture (post-routing / PRD-aligned)

This document describes the **to-be** architecture for implementing `docs/specs/PRD-excel-mcp-transport-routing.md`: routing each workbook operation to **file-based I/O** (openpyxl, current behavior) or **COM-based automation** (Excel host on Windows) when the workbook is detected as open in Excel, with explicit modes and observability.

## Goals (traceability)

| PRD theme | Architectural response |
|-----------|-------------------------|
| US-1 / US-2 / FR-3 | `RoutingBackend` selects file vs COM from `auto` \| `file` \| `com` |
| FR-1 / US-4 | `resolve_target(path)` → normalized absolute path; same output used for file, allowlist, and COM comparison |
| FR-2 | `workbook_open_in_excel(resolved_path)` on Windows; injectable for tests |
| FR-4 / FR-5 | `FileWorkbookService` and `ComWorkbookService` implement the **same operation-oriented contract** |
| FR-6 | COM calls serialized on a **single dedicated thread** (queue + worker or equivalent) |
| FR-7 / FR-8 | Optional tool params: workbook transport override, `save_after_write`; COM default no save until requested |
| FR-10 | Do not start Excel for routing by default |
| FR-11 | Allowlist / workspace roots apply to **both** file and COM targets |
| FR-12 | Non-Windows: COM module absent; `com` mode → clear unsupported error; `auto` → file |
| NFR-3 | Structured log fields per routed call |
| NFR-5 | COM dependencies optional (`[com]` extra) |

## Layered view

```text
┌─────────────────────────────────────────────────────────────┐
│  MCP tool handlers (server.py)                               │
│  - Parse args: filepath, optional workbook_transport,       │
│    save_after_write                                         │
└───────────────────────────┬─────────────────────────────────┘
                            │
┌───────────────────────────▼─────────────────────────────────┐
│  Path & policy                                              │
│  - resolve_target(filepath) → normalized absolute path       │
│  - allowlist / workspace roots (unified for stdio + HTTP)   │
└───────────────────────────┬─────────────────────────────────┘
                            │
┌───────────────────────────▼─────────────────────────────────┐
│  RoutingBackend                                             │
│  - resolve_workbook_backend(resolved, mode, tool_kind)      │
│  - workbook_open_in_excel() when mode is auto/com (Win)     │
│  - EXCEL_MCP_COM_STRICT: fail when COM expected but closed  │
│  - Emit structured logs (transport, reason, duration_ms)    │
└───────────┬─────────────────────────────┬───────────────────┘
            │                             │
┌───────────▼──────────────┐   ┌──────────▼───────────────────┐
│  FileWorkbookService      │   │  ComWorkbookService           │
│  Delegates to existing    │   │  xlwings or pywin32 (ADR)    │
│  workbook/sheet/data/…    │   │  All calls via COM executor  │
└───────────────────────────┘   └──────────────────────────────┘
```

## Naming: workbook transport vs wire transport

**Workbook transport** (`auto` | `file` | `com`) selects how the **workbook bytes/state** are accessed. It must not be confused with **`mcp.run(transport="stdio")`**, which selects the **MCP wire** (stdio, sse, streamable-http). Use distinct parameter names (e.g. `workbook_transport`) in code and tool schemas. See ADR 0001.

## Components (target)

### 1. `resolve_target` / path normalization

- Single function (or small module) used by:
  - Allowlist checks
  - File service inputs
  - `workbook_open_in_excel` comparison against Excel `Workbook.FullName`
- Documented order for relative paths: e.g. workspace roots, then cwd (product decision; stdio may gain relative resolution).
- Windows: short/long paths, drive casing, junction/symlink targets; optional file-id equality in a later iteration (PRD P1).

### 2. Allowlist / trust

- Generalize beyond single `EXCEL_FILES_PATH` jail: multiple roots or stricter stdio policy so COM cannot attach to arbitrary open workbooks outside policy (FR-11).
- Same normalized path used for “is this path allowed?” for both backends.

### 3. `workbook_open_in_excel`

- Bind to **running** `Excel.Application` (no auto-start, FR-10).
- Enumerate open workbooks; compare **normalized** full names to `resolve_target` output.
- Policies for duplicates, protected view, read-only: fail with actionable errors (FR-9) — surfaced through shared error types.

### 4. `RoutingBackend`

- Inputs: resolved path, effective mode (`env` default + per-call override), tool metadata (`read` vs `write`), strict flag.
- Outputs: chosen backend + **reason code** for logging (`full_name_match`, `forced_file`, `forced_com`, `excel_not_running`, etc.).
- `auto`: COM if open and allowed, else file.
- `file`: always file.
- `com`: COM if open and allowed; else error (or documented fallback if ADR approves optional fallback flag).

### 5. `FileWorkbookService`

- Thin façade over existing `workbook`, `sheet`, `data`, `formatting`, `calculations`, `chart`, `pivot`, `tables`, `validation`.
- Consolidates **inline** workbook opens in `server.py` (e.g. data validation path) so all file access goes through one place.
- Optionally normalizes **workbook lifecycle** (`close()` on all read paths) as debt paydown.

### 6. `ComWorkbookService`

- Implements the **same method surface** as `FileWorkbookService` for **routed** operations agreed in the tool inventory.
- **Phased parity:** high-value writes first (`write_data`, `apply_formula`, simple formatting); charts last or tool-forced file until COM chart adapter exists (ADR 0004).
- **`save_after_write`:** default false at host; persist only when true (FR-8).

### 7. COM execution model

- Dedicated **single-thread** worker for COM apartment rules (FR-6).
- Sync MCP handlers submit work and block on result; no requirement to convert all tools to `async def` unless FastMCP evolves.

### 8. Observability

- Every routed operation logs a single JSON line (INFO on logger `excel-mcp.routing`, no stdout) with ADR 0001–aligned fields including **`workbook_transport`**, **`workbook_backend`**, **`routing_reason`**, **`duration_ms`**, **`workbook_path`** (basename redaction by default; optional full path via env), **`operation_name`**, and optional **`mcp_tool_name`** (NFR-3). See `README.md` — *Routing observability*.

### 9. Packaging, CI/CD, and releases

- **`pyproject.toml`:** optional dependency group `[project.optional-dependencies] com = [...]` (NFR-5). Declare **dev** dependencies (e.g. `pytest`, linters) so local runs, PR CI, and release gates use the same toolchain.
- **CI:** default job **without** COM stack; router and contract tests with mocks; pin Python minor in workflows for reproducibility. Optional Windows + Excel manual checklist remains PRD acceptance (see plan stories).
- **Governance:** treat **validation**, **distribution build** (`hatch build`), and **PyPI publish** as separate concerns with one **reusable** quality gate; use **trusted publishing** (OIDC) for uploads; document stable check names for branch protection. Full policy: [ci-cd-packaging-governance.md](ci-cd-packaging-governance.md).
- **Versioning and release notes:** SemVer for `0.y.z`, Conventional Commits as bump hints, changelog template, and PyPI channel strategy (PEP 440 pre-releases / TestPyPI). See [release-versioning-policy.md](release-versioning-policy.md).

## Tool routing (product default)

| Class | Default `auto` behavior (target) |
|-------|----------------------------------|
| **Write-class** | Route to COM when workbook open in Excel and allowed; else file |
| **Read-class** | **File-backed only** (openpyxl on disk). No COM read path in v1 (ADR 0003). Optional COM reads may be added later behind an explicit opt-in. |
| **`save_workbook` (new)** | Routed like other **write-class** tools: COM when workbook is open in Excel (persist host state to disk); file stack when not. Lets agents **explicitly save** before file reads when using COM without `save_after_write` on every mutation (ADR 0003). |

`create_workbook` / net-new unsaved books: remain **file**-first; not routable by path until saved and opened in Excel (PRD out-of-scope for unstable paths).

## Migration from pre-fork

1. Introduce path module + allowlist without changing behavior. **(Done — Epic 2.)**
2. Add `FileWorkbookService` façade; move server inline loads. **(Done — Epic 3.)**
3. Add `RoutingBackend` with file-only implementation and mocked `workbook_open_in_excel`. **(Done — Epic 4.)**
4. Extend tool schemas + env vars; document matrix in README. **(Done — Epic 5.)**
5. Add `ComWorkbookService` behind `[com]` and thread executor; expand operation coverage. **(Epic 6 delivered — skeleton + executor + packaging; Epic 7 — broaden COM write parity and release hardening.)**

## Related documents

- `docs/specs/PRD-excel-mcp-transport-routing.md` — requirements
- `docs/excel-mcp-fork-com-vs-file-routing.md` — blueprint
- `docs/architecture/pre-fork-architecture.md` — baseline
- `docs/architecture/ci-cd-packaging-governance.md` — CI/CD, PyPI, reusable workflows, permissions
- `docs/architecture/release-versioning-policy.md` — SemVer, tags, changelog, release automation
- `docs/architecture/adr/` — decisions
