---
kind: story
id: STORY-10-5
title: Tests, TOOLS.md, README, and operator UX for COM reads
status: draft
parent: EPIC-10
depends_on:
  - STORY-10-1
  - STORY-10-2
  - STORY-10-3
  - STORY-10-4
traces_to:
  - path: docs/architecture/adr/0007-com-read-class-tools-routing.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
  - path: TOOLS.md
  - path: README.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - TOOLS.md documents each read-class tool’s behavior under file vs COM (including opt-in knobs), stale vs live grid, cloud URLs, and interaction with save_workbook (ADR 0003).
  - README sections mention COM read opt-in, defaults, and compatibility caveats (disk path + Excel open).
  - MCP manifest / server instructions updated if tools gain parameters or env vars (match existing patterns from Epic-9).
  - Test plan executed or documented—routing matrix tests, handler wiring tests, COM implementations where CI allows; MANUAL-WINDOWS checklist updated if applicable.
  - ADR 0007 linkage: planning artifact references acceptance once maintainers accept the ADR (no secrets in docs).
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-10-5: Tests, TOOLS.md, README, and operator UX for COM reads

## Description

Close the **documentation and quality gate** for Epic-10: consolidate **operator-facing** guidance (**`TOOLS.md`**, **`README.md`**, optional **`manifest.json`** / **`server.py`** instructions), extend automated tests for routing and parity, and align **manual Windows** validation notes with COM read scenarios.

## User story

As an **operator**, I can **configure and predict COM-backed reads** without reading source code, and **trust CI** for regressions on routing and contracts.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Avoid duplicating LintPlan or repo scripts not present; validation is **markdown and pytest coverage** as exists today.
- Cross-link **ADR 0005** for strict/fallback semantics on failed COM reads.
- If Epic-8 release processes apply, note packaging/changelog expectations briefly without inventing automation.

## Dependencies (narrative)

Depends on **STORY-10-1** (routing semantics stable), **STORY-10-3**, and **STORY-10-4** (implementations to describe accurately). **STORY-10-2** is implicitly covered once behavior is final.
