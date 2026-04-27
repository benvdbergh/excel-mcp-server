# COM-first default routing and workbook session design

This document is the **target** software-architecture view for [ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md): **COM-first** when viable, **file/openpyxl** fallback, **explicit** save and host session tools, and **read/write parity** on the default path. It complements the narrower [COM read-class tools design](com-read-class-tools-design.md) (parity and implementation detail for read tools).

**Scope:** Routing matrix, Excel session model, breaking changes, cloud locators, SSE jail, COM threading, security/allowlist, and **tool inventory** evolution. **No code** in this document—implementation follows separately.

---

## 1. Routing matrix (target)

| Condition | `transport=auto` | `transport=com` | `transport=file` |
|-----------|-------------------|-----------------|-------------------|
| **ToolKind.READ** | COM if viable + identity match; else file | COM if viable; else error/fallback per [ADR 0005](adr/0005-com-strict-and-fallback-controls.md) | Always file |
| **ToolKind.WRITE** (non–file-forced) | Same COM-first / file fallback | Same as today (force COM when viable) | Always file |
| **ToolKind.V1_FILE_FORCED** | File ([ADR 0004](adr/0004-chart-pivot-com-parity-scope.md)) | File when forced | Always file |
| **COM not viable** (no Windows / no executor) | File if allowed by strict mode; else error | Per ADR 0005 | File |

**Invert vs current `RoutingBackend` behavior:** today **`ToolKind.READ` → file unconditionally** (`read_class_file_backed`). Target: **READ uses the same decision tree as WRITE** (with the file-forced exception for chart/pivot).

**Reason strings / metrics:** Reuse and extend existing reason codes (e.g. `full_name_match`, `auto_workbook_not_open_file`); add explicit codes for **“COM read selected”** vs **“file read fallback”** for observability.

**Strict / fallback:** [ADR 0005](adr/0005-com-strict-and-fallback-controls.md) remains in force. **Caution:** file **fallback** after a failed COM read can return **stale** on-disk data; product may prefer **fail-closed** for reads when the operator chose COM and identity matched—document in TOOLS and server behavior.

---

## 2. Session model (Excel Application, Workbooks)

- **`Excel.Application`:** Single COM apartment / executor thread ([ADR 0002](adr/0002-com-automation-stack.md)). All **open / save / close / mutate / read** operations that touch the host go through the **same COM queue**.
- **`Workbooks` / `Workbook`:** A **session** is **not** a separate MCP object ID today; identity is **path or normalized cloud URL** matching `Workbook.FullName` ([ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md)).
- **Implicit vs explicit open:**
  - **Implicit:** Operator passes `filepath`; routing checks **open-detector** + **FullName** match. If the book is **not** open, **auto** falls back to **file** (for paths) or errors for **HTTPS** locators (file cannot open them).
  - **Explicit (new):** **“Open in Excel”** tool loads a **disk path** into the host so **subsequent** operations can use **COM-first** without relying on the user to have pre-opened the file. This reduces “path passed but Excel never opened” surprise.
- **`create_workbook` vs file create (ADR 0008):** Today **`create_workbook`** is a **WRITE** in [`tool_inventory.py`](../../src/excel_mcp/routing/tool_inventory.py)—**file-side** creation. Product options: (a) keep one tool, add **optional `open_in_excel: bool`**, (b) split **“create on disk”** and **“create and open in Excel”** for clearer agent scripts. **No extra default save** after create once `save_after_write` is removed—agent calls **`save_workbook`** if they need persistence before close.

---

## 3. Cloud HTTPS locators ([ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md))

- **COM-first default** matches the common case: workbook **already open** in Excel with **https `FullName`**.
- **File backend** cannot read/write these locators; **routing** must select **COM** when the operator targets the cloud identity.
- **URL allowlist** (`EXCEL_MCP_ALLOWED_URL_PREFIXES` per ADR 0006) remains the **authorization** layer for arbitrary HTTPS strings—**not** the routing table alone.

---

## 4. SSE / HTTP jail (`EXCEL_FILES_PATH`)

**Source of truth:** [`path_policy.py`](../../src/excel_mcp/path_policy.py) (FR-11).

- **stdio** (jail unset): absolute paths and **allowed** HTTPS workbook strings (policy-dependent).
- **SSE / streamable HTTP** (jail set): resolved workbook path must sit under **`realpath(EXCEL_FILES_PATH)`**; **cloud HTTPS** locators are **rejected** in current server behavior when the jail is active (see `server.py` path resolution layer).

**Implication for “open in Excel”:** If the operator runs **SSE** with a jail, **only paths inside the jail** are valid for file-based open/create. They must **copy or create** the workbook into the jail first, then **open in Excel** (or rely on a synced path that resolves inside the jail). **Cloud URLs** may require **stdio** or a **future policy change** if remote operators need HTTPS + HTTP transport—explicitly a **product / security** follow-up, not implied by this design.

---

## 5. Threading ([ADR 0002](adr/0002-com-automation-stack.md))

- **Open, save, close, read, mutate** on COM: **same dedicated COM thread** as today’s writes.
- **No** open/save/close from arbitrary worker threads without marshaling to the COM executor.
- **Ordering:** `close_workbook` after in-flight `write` must be **sequenced** on the COM queue to avoid races with Excel’s dialog or save prompts.

---

## 6. Security and trust (FR-11)

- **Open / create** from MCP **must** pass the same **path allowlist** and **URL prefix allowlist** as other tools—**no** “lifecycle tools are trusted” bypass.
- **Explicit open in Excel** increases the **blast radius** of a malicious path (Excel executes macro-enabled content, DDE, etc. per host policy). Document **operator responsibility**: allowlist, AV, **Trust Center**, and **non-macro** templates where required.
- **Relative paths in HTTP modes:** must remain **jail-relative** as today.

---

## 7. Tool inventory (`tool_inventory.py`) — kinds and new tools

**Today:** `ToolKind` = `READ` | `WRITE` | `V1_FILE_FORCED` ([`tool_inventory.py`](../../src/excel_mcp/routing/tool_inventory.py)).

**New / adjusted tools (ADR 0008):**

| Tool (conceptual) | Suggested kind | Notes |
|-------------------|----------------|-------|
| `save_workbook` | **WRITE** (or **SESSION**—see below) | Already present; becomes the **only** persistence knob for routine mutations. |
| **Open in Excel** (name TBD) | **WRITE** or new **SESSION** | Host-side effect; not “read” in the data sense. |
| **Close in Excel** (name TBD) | **WRITE** or **SESSION** | Same. |
| **File create** / `create_workbook` | **WRITE** | Clarify single vs split tool; optional open flag. |

**Optional new enum value `ToolKind.SESSION` (or `LIFECYCLE`):**

- **Purpose:** Distinguish **host session control** (open/close) from **grid mutation** for docs, metrics, and **policy** (e.g. stricter logging for `Workbooks.Open`).
- **Routing:** Session tools are **COM-primary** for open/close; **file** backend may be **N/A** or **no-op** with clear errors—unlike **READ/WRITE** which dual-path through file/COM.
- **Alternative:** Keep **`WRITE`** for all and tag session tools in **manifest metadata** only—simpler enum, less branching in `RoutingBackend`.

**Reads:** No new read **tools**—existing read tools **gain** COM execution via routing change.

---

## 8. Breaking changes (summary)

| Area | Before | After (ADR 0008) |
|------|--------|------------------|
| Default backend for reads | Always file | COM-first when viable + match |
| `save_after_write` | On many mutating tools + env default | **Removed** |
| Persistence contract | Implicit / env / per-tool flag | **Explicit `save_workbook`** |
| Session | Implicit match only | **Optional explicit open/close** |

---

## 9. Cross-links

| Document | Relevance |
|----------|-----------|
| [ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md) | Decision record |
| [ADR 0003](adr/0003-read-path-com-parity.md) | Historical: file reads; explicit save |
| [ADR 0005](adr/0005-com-strict-and-fallback-controls.md) | Strict / fallback |
| [ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md) | HTTPS locators |
| [com-read-class-tools-design.md](com-read-class-tools-design.md) | COM read parity details |
