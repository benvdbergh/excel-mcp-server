---
kind: epic
id: EPIC-10
title: COM read-class tools, routing, and host parity
status: superseded
superseded_by: EPIC-11
depends_on:
  - EPIC-6
  - EPIC-9
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/com-read-class-tools-design.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/adr/0007-com-read-class-tools-routing.md
  - path: docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md
slice: vertical
acceptance_criteria:
  - Read-class tools can route to COM when opt-in is active and ADR 0007 resolution rules are met (cloud https locators, transport/“open in Excel” cases per decided product rules); default remains file-backed reads for backward compatibility.
  - All read-class MCP handlers pass a real com_do_op so routed_dispatch can execute ComWorkbookService read methods without ComExecutionNotImplementedError.
  - ComWorkbookService read methods are non-stub implementations that match the FileWorkbookService contract and documented JSON shapes, with intentional deltas called out in TOOLS.md or README.
  - Operators have clear documentation for opt-in knobs, stale-read vs live-host behavior, strict/fallback interaction for reads, and cloud URL identity (Epic 9, ADR 0006).
  - Automated tests cover routing branches, handler wiring, and COM read behavior where the test environment supports it; remaining gaps documented in STORY-10-5.
created: "2026-04-27"
updated: "2026-04-27"
---

> **Superseded:** The narrative, sequencing, and default read behavior described here assumed **file-default reads** and **opt-in COM reads** under [ADR 0007](../../architecture/adr/0007-com-read-class-tools-routing.md). **[ADR 0008](../../architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md)** and **[Epic-11](Epic-11-com-first-session-and-lifecycle.md)** replace this epic. Story files **Story-10-*** remain as **historical** references; do not execute against this plan for current delivery.

# Epic-10: COM read-class tools, routing, and host parity

## Description

Deliver **COM-backed execution** for **read-class** MCP tools (`ToolKind.READ`) behind an **explicit opt-in**, extend **`RoutingBackend.resolve_workbook_backend`** so reads are no longer unconditionally `read_class_file_backed`, wire **`com_do_op`** for every read handler, and replace **`ComWorkbookService`** read stubs with implementations aligned to **`FileWorkbookService`** and the design note [COM read-class tools](../../architecture/com-read-class-tools-design.md). This epic implements the direction captured in [ADR 0007](../../architecture/adr/0007-com-read-class-tools-routing.md) once product agrees the draft decision.

## Rough effort

**Total (epic):** approximately **3–5 developer-weeks** (one senior/mid implementer on Windows with Excel), assuming ADR 0007 choices are settled early.

| Story | Rough sizing |
|-------|----------------|
| [10-1](../Stories/Story-10-1-routing-and-com-read-opt-in.md) | ~3–5 days |
| [10-2](../Stories/Story-10-2-wire-com-do-op-on-read-handlers.md) | ~2–4 days |
| [10-3](../Stories/Story-10-3-read-range-with-metadata-com-implementation.md) | ~5–8 days |
| [10-4](../Stories/Story-10-4-remaining-read-operations-com-parity.md) | ~5–8 days |
| [10-5](../Stories/Story-10-5-tests-docs-tools-and-operator-ux.md) | ~3–5 days |

Parallelization: after **10-2**, **10-3** and **10-4** can proceed in parallel by different contributors; **10-5** follows integration of both.

## Risks

| Risk | Mitigation |
|------|------------|
| **Behavior change** for the same disk path when Excel has the workbook open (live host vs file snapshot) | Default file reads; document opt-in; consider requiring `workbook_transport="com"` for host reads per design note §7. |
| **COM threading / latency** (large ranges) | Reuse single-thread executor (ADR 0002); batch `Value2` where possible; document performance expectations. |
| **Value vs formula vs display** semantics | Lock parity with `read_excel_range_with_metadata` / file contract; document any COM-only differences. |
| **Strict mode and fallback** (ADR 0005) | Fail-closed or no silent fallback to stale file for reads when policy demands; align with operator docs. |
| **Cloud `FullName` matching** | Rely on Epic 9 normalization and allowlist; extend tests for read routing with https locators. |
| **ADR 0007 still “Proposed”** | Time-box product decision; engineering can prototype behind feature flag or env until ADR is accepted. |

## Dependencies (prior epics and ADRs)

- **[Epic-6](Epic-6-com-packaging-executor-and-comworkbookservice-skeleton.md):** COM packaging, executor, `ComWorkbookService` skeleton (ADR [0002](../../architecture/adr/0002-com-automation-stack.md)).
- **[Epic-9](Epic-9-sharepoint-and-cloud-workbook-locators-for-com.md):** Cloud workbook locators and URL allowlist for COM identity (ADR [0006](../../architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md)).
- **ADR [0003](../../architecture/adr/0003-read-path-com-parity.md):** Baseline file-backed reads and `save_workbook` story; Epic-10 extends Phase 2 opt-in COM reads without breaking v1 defaults.
- **ADR [0007](../../architecture/adr/0007-com-read-class-tools-routing.md):** Target routing and opt-in model for read-class tools (draft; epic implements once accepted).

## User stories (links)

- [Story-10-1](../Stories/Story-10-1-routing-and-com-read-opt-in.md) — Routing rules, opt-in surface (`WorkbookOperationMetadata.com_read_opt_in`, env/tool params).
- [Story-10-2](../Stories/Story-10-2-wire-com-do-op-on-read-handlers.md) — Pass `com_do_op` from all read handlers through `_workbook_dispatch`.
- [Story-10-3](../Stories/Story-10-3-read-range-with-metadata-com-implementation.md) — `read_range_with_metadata` over COM (highest-value / heaviest path).
- [Story-10-4](../Stories/Story-10-4-remaining-read-operations-com-parity.md) — Other read contract methods (metadata, merged cells, validation, range validate, formula syntax).
- [Story-10-5](../Stories/Story-10-5-tests-docs-tools-and-operator-ux.md) — Tests, README, `TOOLS.md`, manifest / tool instructions as needed.

## Recommended sequencing

1. **10-1** then **10-2** (vertical slice: route + dispatchable COM path, even if stubs return structured “not ready” only in early increments if needed; prefer completing 10-1/10-2 before wide COM impl).
2. **10-3** and **10-4** in parallel after **10-2**.
3. **10-5** last to consolidate tests and operator-facing documentation.

## Definition of Done (epic)

- Acceptance criteria in frontmatter satisfied and ADR 0007 status updated by maintainers when the decision is accepted.
- No regression to default file-only read behavior without opt-in.
- CI green; Windows/manual validation note for COM reads where automated coverage is incomplete.
