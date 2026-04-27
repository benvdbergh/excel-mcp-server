---
kind: story
id: STORY-9-1
title: Cloud locator parsing and get_excel_path
status: done
parent: EPIC-9
depends_on: []
traces_to:
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
  - path: src/excel_mcp/path_resolution.py
  - path: src/excel_mcp/server.py
slice: vertical
acceptance_criteria:
  - https workbook targets (SharePoint-style) are accepted as valid workbook locator strings for stdio when they match documented rules (scheme allowlist, no NUL, basic validation).
  - resolve_target / get_excel_path does not apply os.path.realpath to https locators or treat them as relative.
  - Routing to file backend for a cloud locator returns a clear, stable error (e.g. auto or file).
  - Unit tests cover acceptance vs rejection (malformed URL, wrong scheme) and file-backend branch.
created: "2026-04-28"
updated: "2026-04-28"
---

# Story-9-1: Cloud locator parsing and get_excel_path

## Description

Introduce **cloud workbook locator** detection and wire it through **`get_excel_path`** (or a small helper used there) so **`https://…`** strings are no longer rejected solely for failing `os.path.isabs`. File (`openpyxl`) operations must **not** receive raw URLs.

## Implementation notes

- Prefer a dedicated helper (e.g. `parse_workbook_locator` / `is_cloud_workbook_locator`) in `path_resolution.py` or adjacent module to keep **`resolve_target`** filesystem-pure.
- Consider interaction with **`EXCEL_FILES_PATH`** (SSE/HTTP): explicitly document or guard cloud locators in that mode in this story or defer with a clear `ValueError` until Story-9-2.

## Definition of Done

- Tests pass; CI green.
- No regression for existing disk path behavior (Epic-2 equivalence tests still pass).
