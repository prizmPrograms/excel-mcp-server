# Excel MCP Server - Copilot Instructions

MCP (Model Context Protocol) server that enables VBA automation in running Excel instances on Windows via COM interop.

## Build and Test Commands

```bash
# Build (compiles TypeScript + copies PowerShell script)
npm run build

# Watch mode for development
npm run watch

# Run the MCP server
node dist/index.js
```

No automated tests exist. Manual testing requires:
1. Excel running with workbook open
2. MCP server running
3. MCP client (e.g., Claude Desktop) connected

## Architecture

### Three-Layer Design

1. **excel-com.ps1** (PowerShell)
   - Direct COM automation of Excel Application objects
   - Functions named like `Get-VBAModules`, `Set-Code`, `Run-Macro`
   - Returns JSON via `ConvertTo-Json`
   - Exits with code 1 on errors (stderr captured)

2. **excel-wrapper.ts** (TypeScript)
   - Spawns PowerShell with `-ExecutionPolicy Bypass` to run excel-com.ps1
   - Maps JavaScript methods to PowerShell commands
   - Parses JSON responses and handles errors
   - Uses ES modules (`.js` extensions in imports required)

3. **index.ts** (MCP Server)
   - Implements MCP SDK server using StdioServerTransport
   - Defines tool schemas (input/output descriptions)
   - Routes tool calls to ExcelWrapper methods
   - Returns results as MCP tool responses

### Key Design Decisions

**Why PowerShell instead of native Node.js COM?**
- Avoids dependency on `winax` or similar native modules that require Visual Studio Build Tools
- PowerShell has built-in COM support on Windows
- Reduces installation friction for users

**Why `-ExecutionPolicy Bypass`?**
- Script execution policy varies by system configuration
- Bypass ensures scripts run without requiring system-wide policy changes
- Safe since script path is controlled by this package

**Module Type Constants (VBA)**
- `1`: Standard Module (for Sub/Function procedures)
- `2`: Class Module
- `3`: UserForm

## Important Conventions

### PowerShell Script Editing

When modifying `src/excel-com.ps1`:

- **Encoding**: Save as UTF-8 without BOM
- **Comments**: Write in English only (Japanese comments cause encoding issues when executed via Node.js spawn)
- **Error handling**: Use `Write-Error` + `exit 1` pattern
- **Return values**: Let `ConvertTo-Json` output naturally without explicit `return`
- **Multi-instance limitation**: COM typically accesses only the first Excel process; multiple Excel.exe instances are not fully supported

### TypeScript/MCP Integration

- **ES Modules**: Use `.js` extensions in import paths (TypeScript compiles to `.js` but paths must match runtime)
- **Tool names**: Use snake_case (MCP convention), e.g., `get_workbooks`, `run_macro_safe`
- **Tool descriptions**: Written in Japanese to match user-facing documentation
- **Input schemas**: Required fields must be specified in `required` array

### Build Process

The custom build script does two things:
1. Compiles TypeScript with `tsc`
2. Copies `src/excel-com.ps1` to `dist/excel-com.ps1` (required at runtime)

Without the copy, ExcelWrapper cannot find the PowerShell script.

## Special Tools

### run_macro_safe

Automatically installs a `_MCPHelper` module in the target workbook to capture VBA runtime errors. Returns structured error info:
- `VBAErrorNumber`: VBA error code (e.g., 13 = type mismatch)
- `ErrorDescription`: Human-readable error message
- `ErrorSource`: Source of the error

This enables AI-driven autonomous debugging: AI can read error details, fetch the module code, fix it, and re-run.

### Immediate Window Tools

`read_immediate_window` and `write_immediate_window` activate Excel windows and send keyboard input (`Ctrl+G`, `Ctrl+A`, `Ctrl+C`, etc.). These tools:
- Interrupt user workflow (Excel window becomes foreground)
- Use clipboard (may overwrite existing clipboard content)
- Require no other applications to intercept keyboard events

Always warn users before using these tools (per tool descriptions in README).

## Commit Message Conventions

From CONTRIBUTING.md:
- `Add:` New features
- `Fix:` Bug fixes
- `Update:` Existing feature changes
- `Docs:` Documentation only
- `Refactor:` Code restructuring
- `Test:` Test-related changes

## Prerequisites

- **OS**: Windows only (COM dependency)
- **Node.js**: 18+
- **Excel**: Must be running with a workbook open
- **Excel Security**: "Trust access to the VBA project object model" must be enabled in Excel Trust Center settings

Without VBA trust setting enabled, all VBA operations fail with "Programmatic access to Visual Basic Project is not trusted".
