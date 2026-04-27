---
kind: story
id: STORY-9-2
title: COM URL matching, allowlist, and operator docs
status: done
parent: EPIC-9
depends_on:
  - STORY-9-1
traces_to:
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
  - path: src/excel_mcp/routing/com_workbook_service.py
  - path: src/excel_mcp/path_policy.py
slice: vertical
acceptance_criteria:
  - COM workbook resolution uses URL normalization aligned with Excel FullName (document rules; unit tests for equivalent URL forms).
  - When EXCEL_MCP_ALLOWED_PATHS is enforced, https locators require EXCEL_MCP_ALLOWED_URL_PREFIXES (or documented fail-closed behavior) so enterprises can scope tenants.
  - ADR 0006 status moved to Accepted (or superseded with rationale).
  - README, manifest tool strings, and optional TOOLS.md cover cloud_filepath, SharePoint open workflow, and “Excel handles Microsoft 365 auth.”
created: "2026-04-28"
updated: "2026-04-28"
---

# Story-9-2: COM URL matching, allowlist, and operator docs

## Description

Complete the **vertical slice**: **`ComWorkbookService`** must match operator-supplied **https** locators to **`Workbook.FullName`** without collapsing URLs through `os.path.realpath`. Align **`path_policy`** with **URL prefix** allowlisting when filesystem allowlist is enabled. Ship **operator documentation** and accept **ADR 0006**.

## Implementation notes

- Reuse existing **Protected View** path logic where URLs appear in `SourcePath`/`SourceName` if needed (see `_protected_view_candidate_paths`).
- Add tests that mock or inject workbook `FullName` as https to prove match after normalization.
- Keep **secrets out of MCP**: documentation states auth is via **Excel / Office** session, not MCP env.

## Definition of Done

- Manual smoke on Windows with SharePoint-opened workbook (optional checklist in Story-7-5 doc if extended).
- Lint and pytest green.
