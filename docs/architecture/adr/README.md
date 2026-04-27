# Architecture Decision Records (ADRs)

This folder records decisions that shape the Excel MCP workbook transport fork and related behavior.

| ADR | Title |
|-----|--------|
| [0001](0001-workbook-transport-vs-mcp-wire-transport.md) | Workbook transport vs MCP wire transport naming |
| [0002](0002-com-automation-stack.md) | COM automation stack (Windows Excel) |
| [0003](0003-read-path-com-parity.md) | Read-path: file-only + explicit `save_workbook` tool; optional COM reads later |
| [0004](0004-chart-pivot-com-parity-scope.md) | Chart and pivot COM parity scope (v1) |
| [0005](0005-com-strict-and-fallback-controls.md) | COM strict mode and optional file fallback |
| [0006](0006-cloud-workbook-locator-sharepoint-urls.md) | Cloud workbook locators (SharePoint `https` URLs) for COM routing *(accepted)* |
| [0007](0007-com-read-class-tools-routing.md) | COM routing for read-class tools *(superseded by 0008)* |
| [0008](0008-com-first-default-and-file-lifecycle-tools.md) | COM-first default routing, explicit file lifecycle tools, `save_after_write` removal *(accepted)* |
| [0009](0009-open-workbook-discovery-tool.md) | Open workbook discovery tool (workbook-level enumeration); `get_workbook_metadata` stays single-book *(accepted)* |

**Convention:** `Status` is one of *Proposed*, *Accepted*, *Superseded* (link to replacing ADR), *Deprecated*.

**Related:** [Pre-fork architecture](../pre-fork-architecture.md) · [Target architecture](../target-architecture.md) · [CI/CD and PyPI governance](../ci-cd-packaging-governance.md) · [Release versioning policy](../release-versioning-policy.md) · [PRD](../../specs/PRD-excel-mcp-transport-routing.md)
