---
kind: epic
id: EPIC-12
title: Open workbook discovery MCP tool (ADR 0009)
status: draft
depends_on:
  - EPIC-11
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0009-open-workbook-discovery-tool.md
  - path: docs/architecture/com-first-workbook-session-design.md
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
slice: vertical
acceptance_criteria:
  - A dedicated MCP tool enumerates open workbooks in the Excel host at workbook granularity (same Excel.Application / COM executor as existing COM tools); multi-instance Excel processes remain out of scope per ADR 0009.
  - Each workbook entry exposes at minimum FullName (exact COM locator), Name, and whether it is the active workbook; JSON shape is documented and stable for agents.
  - get_workbook_metadata remains filepath-required; discovery is documented as step one when the operator lacks a locator (HTTPS SharePoint URLs especially).
  - Automated tests cover enumeration logic where feasible (mocked COM); Linux CI remains green; Windows manual confirmation documented.
  - README, TOOLS.md, and changelog reflect the tool name, COM-only semantics, failure modes (Excel not running), and pairing with lifecycle/read tools.
created: "2026-04-27"
updated: "2026-04-27"
---

# Epic-12: Open workbook discovery MCP tool

## Description

Implement **[ADR 0009](../../../architecture/adr/0009-open-workbook-discovery-tool.md)**: a **discovery** MCP tool that lists **`Workbooks`** currently open in Excel and returns **`FullName`** strings agents reuse with **`get_workbook_metadata`**, **`read_data_from_excel`**, writes, and lifecycle tools—**replacing ad-hoc scripting** for `ActiveWorkbook.FullName`.

This epic **does not** overload **`get_filepath`** semantics on **`get_workbook_metadata`** (explicit non-goal per ADR 0009). Scope is **workbook-level enumeration only** on the executor-bound **`Excel.Application`** ([ADR 0002](../../../architecture/adr/0002-com-automation-stack.md)); **PID / multi-instance** attachment is deferred.

## Rough effort

**Total (epic):** approximately **4–10 developer-days** (one mid implementer on Windows with Excel for validation, plus docs/review).

| Story | Rough sizing |
|-------|----------------|
| [12-1](../Stories/Story-12-1-com-enumerate-open-workbooks-json-contract.md) | ~2–4 days |
| [12-2](../Stories/Story-12-2-mcp-tool-wire-discovery-handler-and-schema.md) | ~2–4 days |
| [12-3](../Stories/Story-12-3-tests-tools-md-readme-operator-validation.md) | ~1–3 days |

**Parallelization:** **12-1** must land before **12-2** (vertical slice contracts JSON on COM executor). **12-3** can overlap late **12-2** once tool name and payload are frozen.

## Risks

| Risk | Mitigation |
|------|------------|
| Excel not running / automation errors | Fail with actionable MCP messages; align with FR-10 (no silent Excel start)—document when discovery returns empty vs errors. |
| Sensitive paths in telemetry/logs | Mirror existing basename redaction patterns for routed operations where logs mention workbook identity. |
| **`ToolKind`** / routing confusion | Discovery does **not** resolve `filepath`; **`RoutingBackend`** may treat tool as SESSION-adjacent or bypass routing—finalize in Story 12-2 with **`tool_inventory.py`** consistency. |

## Dependencies (prior epics and ADRs)

- **[Epic-11](Epic-11-com-first-session-and-lifecycle.md)** delivered: COM-first assumptions and lifecycle narrative stable.
- **[ADR 0009](../../../architecture/adr/0009-open-workbook-discovery-tool.md)** decision record (**accepted**).
- **[ADR 0006](../../../architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md)** — **`FullName`** semantics for **`https`** locators.

## User stories (links)

- [Story-12-1](../Stories/Story-12-1-com-enumerate-open-workbooks-json-contract.md) — COM enumeration on executor, JSON contract, unit tests with mocks.
- [Story-12-2](../Stories/Story-12-2-mcp-tool-wire-discovery-handler-and-schema.md) — MCP tool (name finalized), `server.py` / manifest, routing or no-routing decision, error surfaces.
- [Story-12-3](../Stories/Story-12-3-tests-tools-md-readme-operator-validation.md) — TOOLS.md + README, optional manual Windows checklist line, changelog, CI.

## Recommended sequencing

1. **12-1** — Core value: correct `Workbooks` walk and stable JSON from COM thread.
2. **12-2** — Expose to agents via MCP with a clear schema (optional reserved `detail` parameter for future ADR-aligned expansion, can be stub `null`/no-op).
3. **12-3** — Close the loop for operators and release notes.

## Definition of Done (epic)

- ADR 0009 acceptance met; **`get_workbook_metadata`** contract unchanged (required `filepath`).
- PyPI-quality release notes and version bump policy per [release-versioning policy](../../../architecture/release-versioning-policy.md); feature called out in `CHANGELOG.md` when shipped.
- CI green; Windows spot-check documented for cloud + local path workbooks.
