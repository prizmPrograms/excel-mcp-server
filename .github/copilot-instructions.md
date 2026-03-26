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

### VBA Module Naming

Module names MUST follow VBA rules:
- **Start with a letter** (A-Z, a-z) - NOT underscore or number
- **No spaces or special characters** (underscore is allowed)
- **Maximum 31 characters**

Example:
```typescript
// ✅ CORRECT
const moduleName = `ErrorCapture_${Date.now()}`;
const moduleName = `TempModule123`;

// ❌ WRONG - will cause "マクロが見つかりません" error
const moduleName = `_ErrorCapture_${Date.now()}`;
const moduleName = `${Date.now()}_Module`;
```

### Timing: add_module + edit_vba

After `add_module`, Excel needs a moment to register the module internally. If `edit_vba` is called too quickly, you may get "インデックスが有効範囲にありません" (Index out of range). Always use proper error handling:

```typescript
try {
    await add_module(workbook, moduleName);
    await edit_vba(workbook, moduleName, code);
} catch (error) {
    // If edit_vba fails, an empty module may remain
    // Retry or inform user
}
```

### PowerShell Script Editing

When modifying `src/excel-com.ps1`:

- **Encoding**: Save as UTF-8 without BOM
- **Comments**: Write in English only (Japanese comments cause encoding issues when executed via Node.js spawn)
- **Error handling**: Use `Write-Error` + `exit 1` pattern
- **Return values**: Let `ConvertTo-Json` output naturally without explicit `return`
- **Multi-instance limitation**: COM typically accesses only the first Excel process; multiple Excel.exe instances are not fully supported

### TypeScript/MCP Integration

- **ES Modules**: Use `.js` extensions in import paths (TypeScript compiles to `.js` but paths must match runtime)
- **Tool names**: Use snake_case (MCP convention), e.g., `get_workbooks`, `run_macro`
- **Tool descriptions**: Keep objective and factual - AI usage policies are in Prompts, not tool descriptions
- **Input schemas**: Required fields must be specified in `required` array
- **Prompts**: Guidelines for AI assistants are in `src/prompts/` directory

### Build Process

The custom build script does three things:
1. Compiles TypeScript with `tsc`
2. Copies `src/excel-com.ps1` to `dist/excel-com.ps1` (required at runtime)
3. Compiles `src/prompts/*.ts` to `dist/prompts/*.js` for MCP Prompts feature

Without the PowerShell script copy, ExcelWrapper cannot find the script.
Without the prompts, AI clients cannot receive usage guidelines.

## Special Features (v2.0.0)

### MCP Prompts

The server provides AI usage guidelines via MCP Prompts feature:
- **excel-vba-guidelines**: Comprehensive guidelines including error capture pattern, immediate window protocol, and VBA best practices
- Automatically distributed to AI clients without user configuration
- Updates when the package is updated

### Error Capture Pattern

Replaces the removed `run_macro_safe` tool. AI assistants create temporary modules with error handlers:
- Module name MUST start with a letter (e.g., `ErrorCapture_123`, not `_ErrorCapture_123`)
- Captures VBA runtime errors as JSON: `{"status":"error","number":13,"description":"Type mismatch"}`
- More transparent than run_macro_safe - users can see the temporary module
- Enables autonomous debugging: AI reads error, fixes code, retries

### Immediate Window Tools

`read_immediate_window` and `write_immediate_window` activate Excel windows and send keyboard input (`Ctrl+G`, `Ctrl+A`, `Ctrl+C`, etc.). These tools:
- Interrupt user workflow (Excel window becomes foreground)
- Use clipboard (may overwrite existing clipboard content)
- Require no other applications to intercept keyboard events
- Should only be used with explicit user permission (documented in Prompts)

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
