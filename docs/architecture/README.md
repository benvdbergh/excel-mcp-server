# Architecture documentation

Index of architecture material for this repository.

**Package:** PyPI distribution **`excel-com-mcp`** (fork); import package remains **`excel_mcp`**.

| Document | Purpose |
|----------|---------|
| [target-architecture.md](target-architecture.md) | Current workbook transport architecture and package layout. |
| [pre-fork-architecture.md](pre-fork-architecture.md) | Historical baseline before routing/COM fork work. |
| [ci-cd-packaging-governance.md](ci-cd-packaging-governance.md) | CI/CD layout, reusable workflows, PyPI trusted publishing, permissions, branch checks. |
| [release-versioning-policy.md](release-versioning-policy.md) | SemVer, tags, changelog, Conventional Commits, release automation options. |
| [adr/README.md](adr/README.md) | Architecture Decision Records (ADRs). |
| [com-first-workbook-session-design.md](com-first-workbook-session-design.md) | COM-first vs file routing, Excel session, lifecycle tools, open-workbook discovery ([ADR 0009](adr/0009-open-workbook-discovery-tool.md)), jail/SSE, threading, security. |
| [com-read-class-tools-design.md](com-read-class-tools-design.md) | COM read-class design note, including the pre-Epic-11 file-only state kept for traceability ([ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md)). |
