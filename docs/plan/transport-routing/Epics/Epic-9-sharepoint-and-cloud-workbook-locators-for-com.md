---
kind: epic
id: EPIC-9
title: SharePoint and cloud workbook locators for COM
status: done
depends_on: []
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
slice: vertical
acceptance_criteria:
  - Operators can pass SharePoint-style https workbook URLs as tool filepath when targeting COM routing, matching Excel Workbook.FullName after documented normalization.
  - File backend rejects cloud locators; auto mode does not mis-route cloud locators to openpyxl.
  - When EXCEL_MCP_ALLOWED_PATHS is used, cloud locators are governed by an explicit URL prefix allowlist policy (fail-closed).
  - README / manifest / operator docs describe Excel-handled auth (no MCP token) vs path/URL addressing.
created: "2026-04-28"
updated: "2026-04-28"
---

# Epic-9: SharePoint and cloud workbook locators for COM

## Delivery status

**Done** (2026-04-28): Stories [9-1](../Stories/Story-9-1-cloud-locator-parsing-and-get-excel-path.md) and [9-2](../Stories/Story-9-2-com-url-matching-allowlist-and-operator-docs.md). Operator docs: README, `TOOLS.md`, `manifest.json`, MCP server `instructions` in `server.py`; ADR [0006](../../architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md) accepted.

## Architecture summary (lean)

**Problem:** COM identity for SharePoint-hosted workbooks is an **https `FullName`**, while the stack today models **`filepath` as a disk path** only. Matching and policy (`get_excel_path`, `_norm_workbook_path`) assume filesystem semantics.

**Approach:** Treat **cloud locators** as a parallel category of **workbook target** (not MCP wire transport—see ADR 0001). Excel remains the trust boundary for Microsoft 365 **auth and sync**; the MCP only **compares** operator-supplied strings to **COM `FullName`** after stable normalization.

**Non-goals (v1):** Calling SharePoint/Graph APIs, headless cloud open without Excel, or embedding OAuth in the MCP.

**Size estimate:** **Small–medium** (roughly **2–4** focused dev days): one path/locator module touch, COM normalizer + tests, allowlist extension, docs/ADR. Fits **two vertical stories** below.

## User stories (links)

- [Story-9-1](../Stories/Story-9-1-cloud-locator-parsing-and-get-excel-path.md) — Parse cloud locators, extend `get_excel_path`, file-backend rejection, baseline tests.
- [Story-9-2](../Stories/Story-9-2-com-url-matching-allowlist-and-operator-docs.md) — COM URL normalization + matching, URL allowlist with `EXCEL_MCP_ALLOWED_PATHS`, ADR 0006 acceptance, README/manifest.

## Dependencies (narrative)

Builds on delivered **path normalization (Epic-2)** and **COM services (Epics 6–7)**. No PRD change required if framed as FR-1 extension + operator ergonomics for Microsoft 365; optional PRD footnote if product wants explicit AC.

## Related sources

- `docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md`
- `src/excel_mcp/server.py` — `get_excel_path`
- `src/excel_mcp/routing/com_workbook_service.py` — `_norm_workbook_path`, `_get_open_workbook_com`
