# Game plan: forked Excel MCP with open-workbook detection (file vs COM)

This document describes how to evolve a **fork** of the Cursor **`user-excel`**-style MCP so each mutating (and optionally read) operation **chooses a transport** based on whether the target workbook is **already open in a local Microsoft Excel instance**.

**Goal:** When Excel owns the session, edits go **through the running application** (COM), so the user sees **live UI updates** and cloud sync follows Excel’s normal save/co-authoring path. When the workbook is **not** open in Excel, keep today’s **file-based** path (openpyxl / direct xlsx IO / whatever the upstream MCP uses) for batch, CI-style, and headless use.

**Non-goals (initial fork):** macOS Excel COM parity (different story), driving **LibreOffice** via UNO, or replacing Graph API / SharePoint-only workflows.

**Implementation status (2026-04-25):** Phases through **router + env + MCP tool wiring** are implemented in this fork; see [`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](plan/transport-routing/IMPLEMENTATION-ROADMAP.md) and the PRD *Engineering progress* section. COM-backed execution and packaging remain future work (Epics 6–7).

---

## 1. Problem statement (why fork)

- **File-based MCP** writes the `.xlsx` on disk while **Excel (desktop or browser session)** may hold another view of truth → OneDrive / merge / “unmerged changes” / last-writer-wins surprises.
- **COM-based** edits target `Excel.Application` → `Workbook` → `Worksheet` → `Range`, so the **visible instance** updates and persistence is **Save through Excel**, aligning with user expectation for “live” editing.

The fork adds a **router** in front of existing implementations instead of replacing them.

---

## 2. High-level architecture

```
                    ┌─────────────────────────┐
  MCP tool call ──► │  resolve_target(path) │
                    └───────────┬────────────┘
                                │
                    ┌───────────▼────────────┐
                    │ workbook_open_in_excel? │
                    └───┬──────────────┬─────┘
                        │yes           │no
            ┌───────────▼───┐   ┌──────▼──────────┐
            │ COM transport │   │ File transport  │
            │ (xlwings /    │   │ (existing MCP   │
            │  pywin32)     │   │  implementation)│
            └───────────────┘   └─────────────────┘
```

- **Preserve** upstream file logic as a **backend module** (minimal churn).
- **Add** a `ComWorkbookBackend` (new module) that implements the **same internal interface** as the file backend (see §5).
- **Add** `RoutingBackend` that delegates after detection.

---

## 3. When is the workbook “open in Excel”?

### 3.1 Primary signal (Windows + Excel installed)

Attach to the **running** Excel automation server and enumerate **open workbooks**:

- Obtain `Excel.Application` via COM (`win32com.client.GetActiveObject("Excel.Application")` or `DispatchEx` if you allow starting Excel—policy choice).
- For each `Workbook` in `Application.Workbooks`:
  - Compare **resolved full paths** of `Workbook.FullName` (or `Saved` + `Name` edge cases) to the **resolved** `filepath` the MCP received.

**Normalization required:**

- Resolve **short vs long paths**, drive letter casing, and **symlink/junction** targets (e.g. OneDrive folder may appear under multiple roots).
- Optional: compare **file identity** using `GetFileInformationByHandle` (volume serial + file ID) when paths differ but point to the same file.

### 3.2 Ambiguous cases (explicit policy)

| Situation | Suggested policy |
|-----------|------------------|
| Workbook open **twice** (two Excel instances—rare) | Prefer instance with **foreground** `Hwnd`, or **fail closed** with error asking user to close duplicate. |
| Path passed is **relative** | Resolve against `cwd` **and** against known vault roots; document resolution order. |
| Workbook is **open but unsaved** (`Saved == False`, name `Book1`) | **No stable path match** → treat as **not routable by path**; optional future: match by **process + window title** (fragile). |
| Excel not running | **File backend** only (or optional: start Excel invisible—usually **avoid** for surprise UX). |
| `.xlsm` macros / protected view | COM may be **read-only** or blocked; detect **ReadOnly**, **ProtectStructure**, **AutoSave** state; **surface clear errors**. |

### 3.3 Feature flags (config)

Fork should support explicit override without silent wrong routing:

- `EXCEL_MCP_TRANSPORT=auto|file|com` (default `auto`).
- `EXCEL_MCP_COM_STRICT=1` → if `auto` and workbook **not** found in Excel, **do not** fall back to file when user expected live edit (optional safety for power users).

---

## 4. Transport selection matrix

| `transport` | Workbook open in Excel? | Behavior |
|-------------|-------------------------|----------|
| `auto` | Yes | **COM** |
| `auto` | No | **File** |
| `file` | Any | **File** (current behavior) |
| `com` | Yes | **COM** |
| `com` | No | **Error** (or optional fallback to file behind `allow_com_fallback=1`) |

Document in MCP README: **`auto` is best default**; **`com`** for “never touch disk outside Excel” experiments.

---

## 5. Implementation layering (concrete steps after fork)

### Step A — Inventory upstream MCP

1. Clone/fork the **actual** `user-excel` server repo (the one Cursor runs—not only the JSON tool descriptors in the IDE cache).
2. Map each MCP tool to **one of**:
   - **Read path** (safe to stay file-based for perf, or COM for parity with visible sheet).
   - **Write path** (must route for this feature).
3. Extract file I/O into `FileWorkbookService` with a narrow API, e.g.:
   - `read_range(path, sheet, a1)`
   - `write_range(path, sheet, start_a1, values_2d)`
   - `apply_formula(...)`, `get_metadata(...)`, etc.

### Step B — Add COM service (Windows-only gate)

1. New module `ComWorkbookService` using **xlwings** (recommended for ergonomics) **or** raw **pywin32**.
2. Implement the **same method signatures** as `FileWorkbookService` for routed tools.
3. COM apartment: run COM calls on a **single thread** (many MCP servers are async—use a **dedicated sync executor** or a lock around COM).
4. After writes, policy:
   - **`auto_save=false`** default: user sees edit, presses Save (good for “live” demos).
   - Optional tool flag `save_after_write` for agents that want persistence.

### Step C — Router

1. `resolve_workbook_backend(path, transport_mode) -> Backend`.
2. Log (structured): chosen transport, reason (matched `FullName`, identity match, fallback), duration.
3. Unit tests with **mock** COM; integration tests optional (marked `manual`) on a machine with Excel.

### Step D — Tool schema (backward compatible)

- Add optional argument to mutating tools: `transport: "auto" | "file" | "com"` (default `"auto"`).
- Optionally add `save_after_write: boolean` for COM path only (ignored for file path if file backend always writes atomically).

### Step E — Security & ops

- COM can drive **any open workbook** the user has in Excel—**same trust model** as “agent can edit paths you pass,” but mistakes can flash on screen. Keep **path allowlist** / **workspace root** options if upstream lacks them.
- **Do not** start Excel as admin from MCP without explicit user opt-in.

---

## 6. Tool-by-tool routing notes (typical fork checklist)

| MCP tool | File OK when closed | Prefer COM when open | Notes |
|----------|---------------------|----------------------|--------|
| `read_data_from_excel` | Yes | Optional | COM read matches visible calc state; file may lag unsaved edits. |
| `get_workbook_metadata` | Yes | Optional | Sheet list from COM includes **unsaved** new sheets. |
| `write_data_to_excel` | Yes | **Yes** | Primary win for live edits. |
| `apply_formula` | Yes | **Yes** | Verify host calc vs file lib differences. |
| `format_range`, merge, rows/cols | Yes | **Yes** | COM richer / matches UI. |
| `create_workbook` | Yes | N/A | Creating new file is naturally file-first; optionally then `Open` via COM. |

---

## 7. Testing matrix (manual + automated)

**Automated (no Excel):** router defaults to file; COM module skipped on non-Windows.

**Manual Windows:**

1. Closed workbook → `write` → open Excel → verify content (file path).
2. Open workbook in Excel → `write` with `auto` → cell updates **without** closing Excel.
3. Open in Excel **and** pass `transport=file` → verify file on disk updates but Excel shows stale until reload (documents risk).
4. OneDrive path: open in Excel, `write` via COM, **Save** → confirm no duplicate “unmerged” compared to parallel file MCP (subjective; capture screenshots).

---

## 8. Rollout strategy

1. Ship fork behind **`transport` default `auto`** but log chosen backend.
2. Dogfood with **internal** users; tune path normalization for OneDrive.
3. Upstream PR (optional): contribute router + `ComWorkbookService` behind `extras_require` `[com]` so default install stays lightweight.

---

## 9. Alternatives if maintaining a fork is too heavy

- **Thin local script** (xlwings) invoked by the agent instead of MCP for “live” edits; keep MCP for file batch.
- **Office Add-in** + local bridge (WebSocket) for Excel-in-process edits (larger build).

---

## 10. References (external)

- Microsoft: [How to manage merge conflicts in Excel cloud files](https://support.microsoft.com/en-us/office/how-to-manage-merge-conflicts-in-excel-cloud-files-535fb3f2-e7c9-4701-bdcd-0c447d284a6f) (context for why COM-through-Excel helps).
- xlwings / pywin32 documentation for **COM threading** and **Excel.Application** lifecycle.

---

## Document control

- **Purpose:** implementation blueprint for a **forked Excel MCP** with **open-workbook detection** and **file vs COM routing**.
- **Owner:** adopt in your fork’s `README` and link from `mcp-user-excel.md` when implemented.
