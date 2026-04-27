# Architecture documentation

Index of architecture material for this repository.

**Package:** PyPI distribution **`excel-com-mcp`** (fork); import package remains **`excel_mcp`**.

| Document | Purpose |
|----------|---------|
| [target-architecture.md](target-architecture.md) | To-be workbook transport routing architecture (PRD-aligned). |
| [pre-fork-architecture.md](pre-fork-architecture.md) | As-is baseline before routing/COM fork work. |
| [ci-cd-packaging-governance.md](ci-cd-packaging-governance.md) | CI/CD layout, reusable workflows, PyPI trusted publishing, permissions, branch checks. |
| [release-versioning-policy.md](release-versioning-policy.md) | SemVer, tags, changelog, Conventional Commits, release automation options. |
| [adr/README.md](adr/README.md) | Architecture Decision Records (ADRs). |
| [com-first-workbook-session-design.md](com-first-workbook-session-design.md) | **Target:** COM-first vs file routing matrix, Excel session, lifecycle tools, jail/SSE, threading, security, tool inventory. |
| [com-read-class-tools-design.md](com-read-class-tools-design.md) | COM read-class tools: `ComWorkbookService` parity, handler wiring, risks; updated for [ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md). |
