---
kind: epic
id: EPIC-11
title: COM-first default routing, COM read parity, explicit lifecycle, and operator docs
status: done
depends_on:
  - EPIC-6
  - EPIC-9
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/com-first-workbook-session-design.md
  - path: docs/architecture/com-read-class-tools-design.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
slice: vertical
acceptance_criteria:
  - Default routing is COM-first when COM runtime is viable and workbook identity matches an open Excel workbook; otherwise file/openpyxl fallback, with ToolKind.READ participating in the same decision tree as writes except documented V1_FILE_FORCED chart/pivot paths per ADR 0004.
  - All read-class MCP handlers supply com_do_op; ComWorkbookService read methods are real implementations aligned to FileWorkbookService contracts and documented JSON shapes.
  - save_after_write is removed from every mutating tool signature and from env-driven defaults; persistence is explicit via save_workbook only, with server manifest and operator docs updated accordingly.
  - New or clarified lifecycle tools exist per ADR 0008 (file create semantics aligned with create_workbook, explicit open in Excel, close with optional save), with naming documented (e.g. excel_open_workbook as referenced in ADR 0008).
  - README, TOOLS.md, operator environment variables, and automated tests reflect the new defaults, lifecycle tools, and breaking changes; chart/pivot exceptions and strict/fallback behavior are documented.
created: "2026-04-27"
updated: "2026-04-27"
---

# Epic-11: COM-first session and lifecycle (replan)

## Description

Replan and deliver **[ADR 0008](../../../architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md)** and **[COM-first workbook session design](../../../architecture/com-first-workbook-session-design.md)**: **invert** routing so **`ToolKind.READ` uses COM-first / file fallback like writes**, implement **full COM read parity** via **`com_do_op`** and **`ComWorkbookService`**, **remove `save_after_write`** in favor of **explicit `save_workbook`**, add **first-class Excel session lifecycle** (create/open/close semantics per ADR 0008), and complete **tests and operator documentation** (README, TOOLS.md, env). This epic **supersedes the product narrative of [Epic-10](Epic-10-com-read-class-tools-and-routing.md)**, which assumed file-default reads and opt-in COM reads under [ADR 0007](../../../architecture/adr/0007-com-read-class-tools-routing.md).

## Rough effort

**Total (epic):** approximately **6–9 developer-weeks** (one senior or mid implementer on Windows with Excel, plus review for docs and CI), including integration and manual validation. Add buffer if `ToolKind.SESSION` (or equivalent) is introduced for lifecycle tools.

| Story | Rough sizing |
|-------|----------------|
| [11-1](../Stories/Story-11-1-com-first-routing-read-write-and-file-forced.md) | ~4–6 days |
| [11-2](../Stories/Story-11-2-remove-save-after-write-explicit-save-only.md) | ~3–5 days |
| [11-3](../Stories/Story-11-3-com-do-op-wiring-all-read-handlers.md) | ~2–4 days |
| [11-4](../Stories/Story-11-4-com-read-range-with-metadata-and-core-reads.md) | ~5–8 days |
| [11-5](../Stories/Story-11-5-remaining-com-read-parity-and-chart-pivot-exceptions.md) | ~5–8 days |
| [11-6](../Stories/Story-11-6-lifecycle-create-open-close-excel-session.md) | ~5–7 days |
| [11-7](../Stories/Story-11-7-tests-readme-tools-md-operator-env-overhaul.md) | ~4–6 days |

**Parallelization:** After **11-1** and **11-2** (routing and save contract are foundational), **11-3** through **11-5** can overlap (wiring then read implementations; **11-4** and **11-5** can split by contributor once **11-3** lands). **11-6** can start after **11-1** in parallel with read work if lifecycle APIs are agreed. **11-7** runs last and in parallel with late integration, but must finish after feature work stabilizes.

## Risks

| Risk | Mitigation |
|------|------------|
| **Default read path becomes live Excel** (unsaved host state vs file snapshot) | Document in TOOLS.md and ADR 0008 consequences; consider fail-closed vs file fallback for COM-chosen reads per ADR 0005. |
| **Breaking tool schemas and agent prompts** (`save_after_write` removal) | Version or changelog callout; one-shot migration note in README. |
| **Lifecycle tools and macros / Trust Center** | Security section in com-first session design; allowlist unchanged. |
| **HTTPS locators and SSE jail** | Document matrix from session design §4; no silent cloud URL success in jailed HTTP modes. |
| **Chart/pivot remain file-forced** | Keep ADR 0004 visibility in routing and docs; tests assert V1_FILE_FORCED branches. |

## Dependencies (prior epics and ADRs)

- **[Epic-6](Epic-6-com-packaging-executor-and-comworkbookservice-skeleton.md):** COM packaging, executor, `ComWorkbookService` (ADR [0002](../../../architecture/adr/0002-com-automation-stack.md)).
- **[Epic-9](Epic-9-sharepoint-and-cloud-workbook-locators-for-com.md):** Cloud locators and URL allowlist (ADR [0006](../../../architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md)).
- **ADR [0008](../../../architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md):** Controlling decision (COM-first, lifecycle tools, remove `save_after_write`).
- **ADR [0004](../../../architecture/adr/0004-chart-pivot-com-parity-scope.md):** Chart/pivot file-forced exception.
- **ADR [0005](../../../architecture/adr/0005-com-strict-and-fallback-controls.md):** Strict and fallback for COM selection failures.
- **Design notes:** [com-read-class-tools-design.md](../../../architecture/com-read-class-tools-design.md) for read handler parity details.

## User stories (links)

- [Story-11-1](../Stories/Story-11-1-com-first-routing-read-write-and-file-forced.md) — COM-first default for READ and WRITE, file fallback, `V1_FILE_FORCED` tests.
- [Story-11-2](../Stories/Story-11-2-remove-save-after-write-explicit-save-only.md) — Remove `save_after_write`; `save_workbook` only; env and server audit.
- [Story-11-3](../Stories/Story-11-3-com-do-op-wiring-all-read-handlers.md) — Wire `com_do_op` for all read-class handlers.
- [Story-11-4](../Stories/Story-11-4-com-read-range-with-metadata-and-core-reads.md) — COM implementation for `read_range_with_metadata` and other high-traffic reads.
- [Story-11-5](../Stories/Story-11-5-remaining-com-read-parity-and-chart-pivot-exceptions.md) — Remaining read parity; chart/pivot routing and docs.
- [Story-11-6](../Stories/Story-11-6-lifecycle-create-open-close-excel-session.md) — File create vs `create_workbook`, `excel_open_workbook`, close with optional save.
- [Story-11-7](../Stories/Story-11-7-tests-readme-tools-md-operator-env-overhaul.md) — Tests, README, TOOLS.md, operator environment overhaul.

## Recommended sequencing

1. **11-1** then **11-2** (vertical slice: correct routing and persistence contract).
2. **11-3** then **11-4** and **11-5** (dispatch wiring, then COM read depth).
3. **11-6** in parallel with **11-4**–**11-5** once **11-1** is stable (or immediately after **11-2** if API names are fixed early).
4. **11-7** last, continuous documentation updates as stories land.

## Definition of Done (epic)

- Frontmatter acceptance criteria satisfied; ADR 0008 design intent reflected in code and published operator docs.
- Epic-10 narrative explicitly superseded in [IMPLEMENTATION-ROADMAP.md](../IMPLEMENTATION-ROADMAP.md); no requirement to retain file-default read behavior.
- CI green; Windows manual validation path documented for COM-first reads and lifecycle tools.
