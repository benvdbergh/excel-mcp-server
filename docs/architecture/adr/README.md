# Architecture Decision Records (ADRs)

This folder records decisions that shape the Excel MCP workbook transport fork and related behavior.

| ADR | Title |
|-----|--------|
| [0001](0001-workbook-transport-vs-mcp-wire-transport.md) | Workbook transport vs MCP wire transport naming |
| [0002](0002-com-automation-stack.md) | COM automation stack (Windows Excel) |
| [0003](0003-read-path-com-parity.md) | Read-path: file-only + explicit `save_workbook` tool; optional COM reads later |
| [0004](0004-chart-pivot-com-parity-scope.md) | Chart and pivot COM parity scope (v1) |
| [0005](0005-com-strict-and-fallback-controls.md) | COM strict mode and optional file fallback |

**Convention:** `Status` is one of *Proposed*, *Accepted*, *Superseded* (link to replacing ADR), *Deprecated*.

**Related:** [Pre-fork architecture](../pre-fork-architecture.md) · [Target architecture](../target-architecture.md) · [PRD](../../specs/PRD-excel-mcp-transport-routing.md)
