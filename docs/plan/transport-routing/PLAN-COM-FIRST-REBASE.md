# Plan: COM-first rebase (ADR 0008 → Epic-11)

This one-pager orients implementers and reviewers on **why** transport planning shifted from **Epic-10** to **Epic-11** after **[ADR 0008](../../architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md)** was accepted.

## Decision anchor

- **[ADR 0008](../../architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md)** defines **COM-first default routing** for reads and writes when viable, **file/openpyxl fallback** otherwise, **removal of `save_after_write`**, **explicit `save_workbook`**, and **lifecycle tools** (file create clarity, open in Excel, close with optional save). It **supersedes** the default-read and opt-in framing of **[ADR 0007](../../architecture/adr/0007-com-read-class-tools-routing.md)** for product direction.

## Design depth

- **[COM-first workbook session design](../../architecture/com-first-workbook-session-design.md)** — routing matrix, Excel session model, threading, security, SSE jail implications, optional **SESSION** kind vs manifest tagging.
- **[COM read-class tools design](../../architecture/com-read-class-tools-design.md)** — handler parity, `com_do_op`, and COM implementation notes aligned with ADR 0008.

## Backlog home

- **[Epic-11 — COM-first session and lifecycle](Epics/Epic-11-com-first-session-and-lifecycle.md)** and **Story-11-*** under `docs/plan/transport-routing/Stories/`.
- **[IMPLEMENTATION-ROADMAP.md](IMPLEMENTATION-ROADMAP.md)** — phased table, Epic-11 dependency graph, effort range, **Epic-10 superseded** narrative.

## Historical artifacts

- **Epic-10** and **Story-10-*** remain in the repo as **historical** plans under the pre-0008 model; do not use them for current sequencing without reconciling to ADR 0008.
