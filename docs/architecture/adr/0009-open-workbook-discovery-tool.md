# ADR 0009: Open workbook discovery tool (workbook-level enumeration)

## Status

Accepted

## Context

Agents and operators often need the **exact workbook locator** Excel exposes—especially **`Workbook.FullName`**, which for cloud-hosted documents is an **`https://…`** string that must match COM identity ([ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md)). Without it, COM-first routing cannot attach to the intended book when local synced paths differ from Excel’s reported URL.

Today, discovery relies on **external scripting** (e.g. `Excel.Application.ActiveWorkbook.FullName`) outside the MCP contract. That works but is brittle for prompts, CI, and repeatable agent workflows.

A separate question arose: should **`get_workbook_metadata`** accept **no `filepath`** and return metadata for **all** open workbooks, folding discovery into an existing tool?

## Decision

### 1. Separate MCP tool for discovery (not an overloaded `get_workbook_metadata`)

Introduce a **dedicated** tool whose sole job is to **enumerate open workbooks** in the Excel host and return **stable identifiers** agents pass into **`get_workbook_metadata`**, **`read_data_from_excel`**, and other filepath-based tools.

**`get_workbook_metadata` keeps a required `filepath`:** one known locator in, one workbook’s metadata out. Do **not** overload it with optional filepath semantics or union response shapes (single vs list).

**Rationale:**

- **Clear contracts:** discovery vs inspection are different use cases, payloads, and failure modes; MCP consumers (including LLM tool callers) behave better with **one intent per tool**.
- **Stable schemas:** avoiding optional `filepath` preserves predictable validation and documentation for per-book reads.
- **Evolution:** discovery can gain **`detail`** levels (minimal vs sheet summaries) without redefining “metadata” for the existing tool.

### 2. Workbook-level scope only (this iteration)

Discovery operates at **`Workbook`** granularity:

- Enumerate the **`Workbooks`** collection for the bound **`Excel.Application`** used by the COM executor ([ADR 0002](0002-com-automation-stack.md)).
- **Out of scope for now:** treating multiple **`Excel.Application`** instances / PID-level attachment, “which Excel process,” or cross-instance federation.

Product may revisit **multi-instance** or **richer session** models later; they are **not** implied by this ADR.

### 3. Minimum useful payload

Each listed workbook should include at least:

- **`FullName`** — exact string for COM matching (disk path or HTTPS URL per [ADR 0006](0006-cloud-workbook-locator-sharepoint-urls.md)).
- **`Name`** — workbook name (short title).
- **Active indicator** — whether this workbook is the active workbook in Excel (so operators know what the UI is focused on).

Optional fields (formatting, window caption, read-only flags) and optional **`detail`** expansions (e.g. sheet names without full used-range scans) are **implementation choices** documented in TOOLS.md when the tool ships.

### 4. Classification and routing

- **COM-primary:** enumeration requires a running Excel host; **no** meaningful file-backend equivalent for “what is open.”
- **Tool kind:** classify as **SESSION** (if `ToolKind.SESSION` exists) or document as **lifecycle/discovery** adjacent to [ADR 0008](0008-com-first-default-and-file-lifecycle-tools.md) open/close tools—**not** a substitute for **READ** grid operations.

Security and allowlist policy (**FR-11**) applies when paths or URLs are **used as targets** for subsequent tools; discovery itself **reports** host state—operators remain responsible for macro/trust posture ([COM-first workbook session design](../com-first-workbook-session-design.md) §6).

## Consequences

- **New MCP tool** (exact name TBD in implementation; e.g. `excel_list_open_workbooks`) appears in the manifest alongside lifecycle tools.
- **`get_workbook_metadata`** remains **filepath-required**; agents use **two-step** flows when they lack a locator: **list → pick → metadata/read**.
- **Tests:** COM-backed enumeration tests on Windows; headless CI continues to mock or skip as today.
- **Documentation:** [COM-first workbook session design](../com-first-workbook-session-design.md) carries the narrative; this ADR is the decision record.

## Links

- [ADR 0006 — Cloud workbook locators](0006-cloud-workbook-locator-sharepoint-urls.md)
- [ADR 0008 — COM-first default and lifecycle tools](0008-com-first-default-and-file-lifecycle-tools.md)
- [ADR 0002 — COM automation stack](0002-com-automation-stack.md)
- [Design: COM-first workbook session](../com-first-workbook-session-design.md)
