# Microsoft Access MCP Server

![.NET 8](https://img.shields.io/badge/.NET-8.0-blue)
![Windows](https://img.shields.io/badge/platform-Windows-blue)
![Status](https://img.shields.io/badge/status-alpha-orange)
![CI](https://img.shields.io/badge/ci-windows--actions-green)

A Windows-native MCP server for Microsoft Access that exposes Access automation over JSON-RPC stdio for AI-enabled IDEs and MCP clients.

## Overview

This repository delivers a complete Microsoft Access Model Context Protocol implementation with secure database access, COM automation, VBA extensibility, and export capabilities.

The project is structured into two main components:

- `MS.Access.MCP.Interop` — .NET 8 library for Access data access, COM automation, and metadata tools
- `MS.Access.MCP.Official` — MCP server implementation that exposes Access functions through JSON-RPC and MCP-compatible tool calls

## Key Features

- Full JSON-RPC / MCP protocol compliance
- STA COM execution via a dedicated Access worker thread
- Async request queue with timeout enforcement and safe response delivery
- Zero stdio pollution for JSON-RPC clients
- Parameterized SQL execution with placeholder validation to reduce injection risk
- Table pagination with `ReadTableData(limit, offset)`
- Microsoft Access form, report, macro, and module discovery
- VBA and macro execution, code retrieval, code injection, and compile support
- System table metadata access and schema introspection
- Export report to PDF from Access
- Generate EF Core model scaffolding from Access schema
- IDE integration guidance for Windsurf, Cursor, Roo Code/Cline, and VS Code

## Prerequisites

- Windows operating system
- Microsoft Access installed
- .NET 8.0 SDK / runtime
- Access database file (`.accdb` or `.mdb`)

## Installation & Setup

1. Clone the repository.
2. Restore and build the solution:

```powershell
dotnet restore My-Access-AI-Server.sln
dotnet build My-Access-AI-Server.sln -c Release
```

3. Publish the official MCP server:

```powershell
dotnet publish MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release -r win-x64 --self-contained false -o .\publish
```

4. Configure your MCP client or IDE to launch the published executable from `publish\MS.Access.MCP.Official.exe`.

## Running the Server

Use the published executable for reliable MCP stdio hosting:

```powershell
.\publish\MS.Access.MCP.Official.exe
```

For development, run the server from source with:

```powershell
dotnet run --project MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release
```

## IDE Integration

See `MCP_IDEs_SETUP.md` for step-by-step examples connecting the server to:

- Windsurf
- Cursor
- Roo Code
- Cline

For Claude Desktop configuration, refer to `CLAUDE_DESKTOP_CONFIG.md`.

## Development & Local Tasks

Local VS Code tasks are configured in `.vscode/tasks.json`:

- `Build MCP Official`
- `Publish MCP Official`

Use these tasks to build, publish, and validate the project locally.

## Testing

Run the complete solution tests with:

```powershell
dotnet test My-Access-AI-Server.sln -c Release
```

The test suite verifies JSON-RPC parsing, tool registration, Access connectivity validation, macro/VBA tooling, PDF export readiness, EF Core model generation, and error handling.

## Release Notes

See `CHANGELOG.md` for the initial `v1.0.0  The Enterprise Release` summary.

## References

- `MCP_IDEs_SETUP.md`  AI IDE integration examples
- `CLAUDE_DESKTOP_CONFIG.md`  Claude Desktop MCP connection settings
- `.vscode/tasks.json`  local build and publish automation
- `MS.Access.MCP.Official`  MCP server implementation
- `MS.Access.MCP.Interop`  Access interop and automation library

## CI Workflow

A Windows-based CI workflow is recommended to validate every commit with:

- `dotnet restore`
- `dotnet build My-Access-AI-Server.sln -c Release`
- `dotnet test My-Access-AI-Server.sln -c Release`

This workflow supports the projects Windows Access COM and .NET 8 requirements.
