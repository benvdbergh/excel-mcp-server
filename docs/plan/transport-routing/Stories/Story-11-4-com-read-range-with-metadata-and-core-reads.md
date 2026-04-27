---
kind: story
id: STORY-11-4
title: COM read — read_range_with_metadata and core read paths
status: draft
parent: EPIC-11
depends_on:
  - STORY-11-3
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/com-read-class-tools-design.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - ComWorkbookService implements read_range_with_metadata (or equivalent contract method) with behavior aligned to FileWorkbookService for value, formula, and metadata fields; COM-only deltas documented in TOOLS.md.
  - Performance-sensitive paths use Value2 or batching where appropriate per com-read-class-tools-design; single-thread COM executor discipline preserved.
  - Automated tests cover COM read for the primary range read where the test host supports it; gaps noted for Story-11-7 if environment-limited.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-4: COM read — `read_range_with_metadata` and core read paths

## Description

Deliver **non-stub COM implementations** for the **highest-traffic read path** (`read_range_with_metadata` and closely related “core” reads) so COM-first routing returns **useful host-backed data**. Align JSON shapes to **FileWorkbookService** and **[com-read-class-tools-design.md](../../../architecture/com-read-class-tools-design.md)**.

## User story

As an **operator**, I want **live grid reads** from Excel to match **agent expectations** for range inspection and validation workflows.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Corresponds to **Epic-10 Story-10-3** scope but under **ADR 0008** defaults.
- Strict/fallback interactions for stale file vs COM failure belong in tests or docs cross-linked with **Story-11-1**.

## Dependencies (narrative)

Requires **Story-11-3**. Can parallelize with **Story-11-5** by splitting method ownership across contributors.
