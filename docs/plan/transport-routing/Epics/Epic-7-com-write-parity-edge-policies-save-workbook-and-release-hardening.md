---
kind: epic
id: EPIC-7
title: COM write parity, edge policies, save_workbook, and release hardening
status: draft
depends_on:
  - EPIC-6
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
slice: vertical
acceptance_criteria:
  - Write-class tools in inventory use COM when auto selects COM and workbook is open; file otherwise (PRD US-1, US-2, release AC).
  - FR-9 cases return actionable errors (protected view, read-only, duplicate instances, unsaved path policy).
  - Chart/pivot behavior matches chosen ADR 0004 policy and is documented on tools and README.
  - save_workbook exists and is documented with read/stale caveats (ADR 0003).
  - README and manual checklist satisfy PRD release-level acceptance.
created: "2026-04-24"
updated: "2026-04-27"
---

# Epic-7: COM write parity, edge policies, save_workbook, and release hardening

## Description

Complete **COM implementations** for routed **write-class** tools per the inventory, apply **edge-case policies**, add **`save_workbook`** per **ADR 0003**, implement **ADR 0004** chart/pivot policy, and close **Gate 3 — Release ready** (README, CI, observability, manual Windows checklist).

## Objectives

- Deliver **US-1** / **US-2** on supported tools with real Excel where applicable.
- Document **read vs disk staleness** and agent use of `save_workbook` (`ADR 0003`).
- Capture **NFR-2** evidence (note or benchmark) as appropriate.

## User stories (links)

- [Story-7-1](../Stories/Story-7-1-implement-com-write-class-operations-per-inventory-matrix.md)
- [Story-7-2](../Stories/Story-7-2-fr-9-edge-policies-and-actionable-errors-protected-view-read-only.md)
- [Story-7-3](../Stories/Story-7-3-chart-and-pivot-v1-policy-per-adr-0004.md)
- [Story-7-4](../Stories/Story-7-4-save-workbook-tool-and-read-vs-disk-documentation-adr-0003.md)
- [Story-7-5](../Stories/Story-7-5-readme-ci-nfr-2-note-and-manual-windows-checklist.md)

## Dependencies (narrative)

Depends on **EPIC-6** (done): COM runtime, executor, and `ComWorkbookService` skeleton are in-tree; this epic broadens write parity and release hardening.

## Related sources

- `docs/architecture/adr/0003-read-path-com-parity.md` — reads file-backed; explicit save tool.
- `docs/architecture/adr/0004-chart-pivot-com-parity-scope.md` — v1 chart/pivot scope.
- `docs/excel-mcp-fork-com-vs-file-routing.md` — blueprint §7 manual checklist (link or inline in docs).
