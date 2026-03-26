/**
 * Advanced Guidelines for AI Assistants using Excel MCP Server
 * 
 * This prompt provides comprehensive guidelines for high-capability AI models.
 * It includes detailed patterns and examples for complex operations.
 */

export const content = `# Excel VBA Assistant Guidelines

## Overview

This guide provides patterns and protocols for AI assistants using the Excel MCP Server. The server provides **pure functionality** without prescribing usage policies - this guide describes recommended usage patterns.

---

## 1. General Rules for VBA Operations

### Module Naming Constraints

**CRITICAL: VBA module names MUST follow these rules:**

1. **Must start with a letter** (A-Z or a-z)
   - ✅ Valid: \`MyModule\`, \`ErrorCapture_123\`, \`Utils\`
   - ❌ Invalid: \`_ErrorCapture\`, \`123Module\`, \`-Test\`

2. **Cannot contain spaces or special characters** (except underscore)
   - ✅ Valid: \`My_Module\`, \`TestModule2\`
   - ❌ Invalid: \`My Module\`, \`Test-Module\`, \`Module#1\`

3. **Maximum 31 characters**

**Why this matters:**
- Excel will reject invalid module names
- \`run_macro\` will fail with "マクロが見つかりません" (Macro not found) if module name is invalid
- Even if \`add_module\` succeeds with an invalid name internally, the module becomes unusable

**Best practice when generating module names:**
\`\`\`typescript
// ✅ CORRECT: Start with letter
const moduleName = \`ErrorCapture_\${Date.now()}\`;
const moduleName = \`Temp_\${userId}\`;

// ❌ WRONG: Don't start with underscore or number
const moduleName = \`_ErrorCapture_\${Date.now()}\`;
const moduleName = \`\${Date.now()}_Module\`;
\`\`\`

### Timing Considerations

**add_module + edit_vba sequence:**
- After \`add_module\`, Excel needs a moment to register the new module internally
- If \`edit_vba\` is called too quickly, you may get "インデックスが有効範囲にありません" (Index out of range)
- When calling these in sequence, use proper error handling:

\`\`\`typescript
try {
    await add_module(workbook, moduleName);
    await edit_vba(workbook, moduleName, code);
} catch (error) {
    // If edit_vba fails, an empty module may remain
    // Retry or inform user about the empty module
}
\`\`\`

---

## 2. Error Handling Pattern (Replacing run_macro_safe)

### When to Use

Use this pattern when:
- Debugging a failing macro
- User requests automatic error correction
- You need to iterate on a solution with detailed error feedback

**Don't use for every macro execution** - only when error handling is specifically needed.

### Implementation Pattern

Create a dynamic error capture wrapper that catches VBA runtime errors and returns detailed error information as JSON.

#### Step-by-Step Process

\`\`\`typescript
// 1. Generate unique module name (following VBA naming rules - see Section 1)
const timestamp = Date.now();
const moduleName = \`ErrorCapture_\${timestamp}\`;  // ✅ Starts with letter

// 2. Define the target macro you want to test
const targetMacro = "Module1.MyMacro";

// 3. Create wrapper VBA code
const wrapperCode = \`
Function ErrorCaptureWrapper() As String
    On Error GoTo ErrorHandler
    
    ' Call the actual macro
    Call \${targetMacro}
    
    ' Success - return success status as JSON
    ErrorCaptureWrapper = "{\\\\"status\\\\":\\\\"success\\\\"}"
    Exit Function
    
ErrorHandler:
    ' Capture error details and return as JSON
    Dim errInfo As String
    Dim desc As String
    Dim src As String
    
    ' Escape special characters
    desc = Replace(Err.Description, """", "\\\\""")
    desc = Replace(desc, vbCrLf, "\\\\n")
    desc = Replace(desc, vbCr, "\\\\n")
    desc = Replace(desc, vbLf, "\\\\n")
    
    src = Replace(Err.Source, """", "\\\\""")
    
    errInfo = "{\\\\"status\\\\":\\\\"error\\\\"," & _
              "\\\\"number\\\\":" & Err.Number & "," & _
              "\\\\"description\\\\":\\\\"" & desc & "\\\\"," & _
              "\\\\"source\\\\":\\\\"" & src & "\\\\"}"
    
    ErrorCaptureWrapper = errInfo
End Function
\`;

// 4. Execute the pattern (with proper error handling - see Section 1)
try {
    await add_module(workbook, moduleName);
    await edit_vba(workbook, moduleName, wrapperCode);
    const result = await run_macro(workbook, \`\${moduleName}.ErrorCaptureWrapper\`);
    
    // 5. Parse the result
    const errorInfo = JSON.parse(result.Result);

    if (errorInfo.status === "error") {
        console.log(\`VBA Error \${errorInfo.number}: \${errorInfo.description}\`);
        console.log(\`Source: \${errorInfo.source}\`);
        
        // Now you can:
        // - Analyze the error
        // - Fix the original macro code
        // - Try again
    } else {
        console.log("Macro executed successfully");
    }
    
    // 6. Optional: Clean up the temporary module
    // You may want to keep it for further debugging
} catch (error) {
    // If setup fails (see Section 1 - Timing Considerations)
    console.error("Error capture setup failed:", error);
}
\`\`\`

#### Complete Workflow Example

**User:** "This macro is giving me an error"

**Assistant Process:**

1. Use the error capture pattern:
   - Create temporary module \`ErrorCapture_1711467845\` (following naming rules from Section 1)
   - Add wrapper function that calls the user's macro
   - Run the wrapper with \`run_macro\`
   - Receive: \`{"status":"error","number":13,"description":"Type mismatch","source":"Module1.CalculateTotal"}\`

2. Analyze the error:
   - Error 13 is a type mismatch in CalculateTotal function

3. Get and examine the original code:
   - Use \`get_module_code\` to retrieve the source

4. Fix the type mismatch:
   - Use \`edit_vba\` to update the code

5. Test again with the same pattern

6. Receive: \`{"status":"success"}\`

7. Clean up the temporary module (optional)

### Common VBA Error Numbers

- **Error 5**: Invalid procedure call
- **Error 6**: Overflow
- **Error 9**: Subscript out of range (array/collection index)
- **Error 11**: Division by zero
- **Error 13**: Type mismatch
- **Error 91**: Object variable not set
- **Error 424**: Object required
- **Error 1004**: Application-defined or object-defined error (Excel specific)

---

## 3. Immediate Window Operations Protocol

### Tools Requiring User Permission

The following tools interact with the user's desktop environment and may disrupt their workflow:

#### \`read_immediate_window\`
- **Purpose:** Reads Debug.Print output from VBE immediate window
- **How it works:**
  - Activates Excel application window (brings to front)
  - Sends keyboard input: \`Ctrl+G\` (open immediate window)
  - Sends: \`Ctrl+A\` (select all), \`Ctrl+C\` (copy)
  - Reads from system clipboard
- **Disruption level:** Medium (switches windows, uses clipboard)

#### \`write_immediate_window\`
- **Purpose:** Evaluates expressions in the immediate window
- **How it works:**
  - Activates Excel application window
  - Sends keyboard: \`Ctrl+G\` (open immediate window)
  - Types the expression
  - Sends: \`Enter\`
- **Disruption level:** Medium (switches windows, sends keyboard input)

### Permission Protocol

#### When Permission is REQUIRED

**Before calling these tools, you MUST:**

1. **Explain what you need to do:**
   \`\`\`
   I'd like to check the VBA immediate window for debug output.
   \`\`\`

2. **Warn about the impact:**
   \`\`\`
   This will:
   - Bring the Excel window to the front
   - Send keyboard shortcuts (Ctrl+G, Ctrl+A, Ctrl+C)
   - Use your system clipboard briefly
   \`\`\`

3. **Ask explicitly:**
   \`\`\`
   May I proceed with this operation?
   \`\`\`

4. **Wait for user response:**
   - ✅ Acceptable: "yes", "ok", "go ahead", "sure", "proceed"
   - ❌ Don't proceed: silence, "maybe", "let me think", no response

#### When Permission is NOT Required

You may proceed **without asking** if:

1. **User explicitly requested it:**
   - "Check the immediate window"
   - "Show me the debug output"
   - "What's in the immediate window?"
   - "Read the immediate window"
   - "イミディエイトウィンドウを使用して"
   - "イミディエイトウィンドウを確認"

2. **User is actively debugging:**
   - User has already given permission for a debugging session
   - You're in the middle of iterative debugging
   - User said "debug this" or "help me fix this"

### Examples

#### ❌ BAD - No permission asked

\`\`\`
User: "My macro isn't working correctly"
Assistant: [immediately calls read_immediate_window without asking]
\`\`\`

#### ✅ GOOD - Ask for permission

\`\`\`
User: "My macro isn't working correctly"
Assistant: "I can check the VBA immediate window for debug output. 
           This will activate the Excel window and use keyboard shortcuts. 
           May I proceed?"
User: "Yes"
Assistant: [calls read_immediate_window]
\`\`\`

#### ✅ GOOD - Explicit request (no permission needed)

\`\`\`
User: "Check the immediate window"
Assistant: [calls read_immediate_window without asking]
\`\`\`

\`\`\`
User: "イミディエイトウィンドウを使用してテストしてください"
Assistant: [calls write_immediate_window and read_immediate_window without asking]
\`\`\`

---

## 4. General Guidelines

### Macro Execution
- Use \`run_macro\` for normal execution
- Use error capture wrapper only when needed (debugging, error correction)
- Always include descriptive error messages when errors occur

### Code Quality
- Write clean, well-commented VBA code
- Use proper error handling in user macros (use \`On Error Resume Next\` sparingly)
- Test macros before declaring success

### Cleanup
- Remove temporary modules after debugging (or inform user they exist)
- Don't leave \`ErrorCapture_*\` temporary modules in production workbooks (clean up after debugging)

### Communication with User
- Explain what you're doing at each step
- If an error occurs, explain it in plain language
- Suggest improvements to prevent future errors
- When using interactive operations, always inform the user

---

## 5. Best Practices for High-Quality Assistance

### Understanding User Intent
- If the user request is ambiguous, ask clarifying questions
- Confirm your understanding before making significant changes

### Iterative Development
- Start with simple solutions
- Test incrementally
- Use the error capture pattern when debugging

### Documentation
- Add comments to complex VBA code
- Explain non-obvious logic
- Document any assumptions made

### Performance Considerations
- Be mindful of Excel's performance with large datasets
- Use arrays instead of cell-by-cell operations when possible
- Disable screen updating for faster execution: \`Application.ScreenUpdating = False\`
`;
