# Changelog

All notable changes to this project are documented in this file.

## [1.0.0] - 2026-04-01
### The Enterprise Release

- Added full JSON-RPC / Model Context Protocol compliance, including `initialize`, `tools/list`, and `tools/call` support.
- Added STA COM host support with a dedicated Access worker thread to safely execute Microsoft Access automation.
- Implemented async request queueing, timeout handling, response dispatch, and zero stdio pollution for MCP clients.
- Added parameterized SQL execution with placeholder validation and safe command parameter binding to reduce injection risk.
- Added `ReadTableData(limit, offset)` pagination support for table row browsing.
- Added macro and VBA tooling, including `run_macro_or_vba`, `get_vba_code`, `set_vba_code`, `add_vba_procedure`, and `compile_vba`.
- Added export/report workflow support, including `export_report_to_pdf` and import/export persistence helpers.
- Added `generate_ef_core_models` to scaffold EF Core model classes from Access schemas.
- Added metadata and discovery tools for system tables, object metadata, forms, reports, macros, modules, and control properties.
- Added IDE integration documentation for Windsurf, Cursor, Roo Code/Cline, and VS Code task automation.
- Added Windows-based CI workflow recommendations for `dotnet restore`, `dotnet build`, and `dotnet test`.
