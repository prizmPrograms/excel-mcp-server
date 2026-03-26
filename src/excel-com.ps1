# Excel COM automation PowerShell script

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Get Excel instance (currently supports single instance or first accessible instance)
function Get-AllExcelInstances {
    $instances = @()
    
    # Check if Excel is running
    $processes = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    
    if ($processes.Count -eq 0) {
        Write-Error "Excel is not running"
        exit 1
    }
    
    # Get the accessible Excel instance
    # Note: Due to COM limitations, only one Excel process can typically be accessed
    # If multiple Excel.exe processes are running, this will connect to the first accessible one
    try {
        # Try GetActiveObject first (PowerShell 5)
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $instances += $excel
    }
    catch {
        Write-Error "Cannot access Excel COM object. Make sure Excel is running and accessible."
        exit 1
    }
    
    if ($instances.Count -eq 0) {
        Write-Error "Cannot access Excel COM object"
        exit 1
    }
    
    # Within a single Excel instance, all workbooks are accessible
    return ,$instances
}

# Find workbook across all instances
function Find-WorkbookInInstances {
    param([string]$WorkbookName)
    
    $instances = Get-AllExcelInstances
    
    foreach ($excel in $instances) {
        foreach ($wb in $excel.Workbooks) {
            if ($wb.Name -eq $WorkbookName) {
                return @{
                    Excel = $excel
                    Workbook = $wb
                }
            }
        }
    }
    
    Write-Error "Workbook not found: $WorkbookName"
    exit 1
}

# Backward compatibility wrapper
function Get-ExcelInstance {
    $instances = Get-AllExcelInstances
    return $instances[0]
}

function Get-Workbooks {
    $instances = Get-AllExcelInstances
    $workbooks = @()
    
    foreach ($excel in $instances) {
        foreach ($wb in $excel.Workbooks) {
            $workbooks += @{
                Name = $wb.Name
                FullName = $wb.FullName
                Path = $wb.Path
                Sheets = $wb.Worksheets.Count
                HasVBProject = $wb.HasVBProject
            }
        }
    }
    
    $workbooks | ConvertTo-Json -Depth 10
}

function Get-VBAModules {
    param(
        [string]$WorkbookName
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    $modules = @()
    
    try {
        $vbProject = $workbook.VBProject
        
        foreach ($component in $vbProject.VBComponents) {
            $modules += @{
                Name = $component.Name
                Type = $component.Type
                CodeModule = @{
                    CountOfLines = $component.CodeModule.CountOfLines
                }
            }
        }
    }
    catch {
        Write-Error "Cannot access VBAProject. Enable Trust access to the VBA project object model in Excel security settings."
        exit 1
    }
    
    $modules | ConvertTo-Json -Depth 10
}

function Get-VBACode {
    param(
        [string]$WorkbookName,
        [string]$ModuleName
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    try {
        $vbProject = $workbook.VBProject
        $component = $vbProject.VBComponents.Item($ModuleName)
        
        if (-not $component) {
            Write-Error "Module not found: $ModuleName"
            exit 1
        }
        
        $lineCount = $component.CodeModule.CountOfLines
        if ($lineCount -gt 0) {
            $code = $component.CodeModule.Lines(1, $lineCount)
        } else {
            $code = ""
        }
        
        @{
            ModuleName = $ModuleName
            Code = $code
            LineCount = $lineCount
        } | ConvertTo-Json -Depth 10
    }
    catch {
        Write-Error "Failed to get VBA code: $_"
        exit 1
    }
}

function Add-VBAModule {
    param(
        [string]$WorkbookName,
        [string]$ModuleName,
        [int]$ModuleType = 1
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    try {
        $vbProject = $workbook.VBProject
        $component = $vbProject.VBComponents.Add($ModuleType)
        $component.Name = $ModuleName
        
        @{
            Success = $true
            ModuleName = $ModuleName
            Type = $ModuleType
        } | ConvertTo-Json
    }
    catch {
        Write-Error "Failed to add module: $_"
        exit 1
    }
}

function Set-VBACode {
    param(
        [string]$WorkbookName,
        [string]$ModuleName,
        [string]$Code
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    try {
        $vbProject = $workbook.VBProject
        $component = $vbProject.VBComponents.Item($ModuleName)
        
        if (-not $component) {
            Write-Error "Module not found: $ModuleName"
            exit 1
        }
        
        $codeModule = $component.CodeModule
        $lineCount = $codeModule.CountOfLines
        if ($lineCount -gt 0) {
            $codeModule.DeleteLines(1, $lineCount)
        }
        
        $codeModule.AddFromString($Code)
        
        @{
            Success = $true
            ModuleName = $ModuleName
            LinesAdded = $codeModule.CountOfLines
        } | ConvertTo-Json
    }
    catch {
        Write-Error "Failed to set VBA code: $_"
        exit 1
    }
}

function Invoke-VBAMacro {
    param(
        [string]$WorkbookName,
        [string]$MacroName,
        [object[]]$Parameters = @()
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    $excel = $result.Excel
    
    try {
        $fullMacroName = "$WorkbookName!$MacroName"
        
        # Build argument array with macro name as first element
        if ($Parameters.Count -gt 0) {
            $runArgs = @($fullMacroName) + $Parameters
            $macroResult = $excel.Application.Run.Invoke($runArgs)
        } else {
            # No parameters - call with just macro name
            $macroResult = $excel.Application.Run($fullMacroName)
        }
        
        @{
            Success = $true
            MacroName = $MacroName
            Result = $macroResult
        } | ConvertTo-Json
    }
    catch {
        Write-Error "Failed to run macro: $_"
        exit 1
    }
}

function Get-SheetNames {
    param([string]$WorkbookName)
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    $sheets = @()
    $index = 1
    foreach ($sheet in $workbook.Worksheets) {
        $sheets += @{
            Name = $sheet.Name
            Index = $index
            Visible = $sheet.Visible
        }
        $index++
    }
    
    $sheets | ConvertTo-Json -Depth 10
}

function Get-RangeValues {
    param(
        [string]$WorkbookName,
        [string]$SheetName,
        [string]$RangeAddress
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    try {
        $sheet = $workbook.Worksheets.Item($SheetName)
        $range = $sheet.Range($RangeAddress)
        
        $rowCount = $range.Rows.Count
        $columnCount = $range.Columns.Count
        
        $values = @()
        for ($i = 1; $i -le $rowCount; $i++) {
            $row = @()
            for ($j = 1; $j -le $columnCount; $j++) {
                $cellValue = $range.Cells.Item($i, $j).Value2
                $row += $cellValue
            }
            $values += ,@($row)
        }
        
        @{
            WorkbookName = $WorkbookName
            SheetName = $SheetName
            RangeAddress = $RangeAddress
            RowCount = $rowCount
            ColumnCount = $columnCount
            Values = $values
        } | ConvertTo-Json -Depth 10
    }
    catch {
        Write-Error "Failed to get range values: $_"
        exit 1
    }
}

function Get-CellValue {
    param(
        [string]$WorkbookName,
        [string]$SheetName,
        [string]$CellAddress
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    
    try {
        $sheet = $workbook.Worksheets.Item($SheetName)
        $cell = $sheet.Range($CellAddress)
        
        $formula = $cell.Formula
        $value = $cell.Value2
        $text = $cell.Text
        
        @{
            WorkbookName = $WorkbookName
            SheetName = $SheetName
            Address = [string]$cell.Address()
            Value = $value
            Text = $text
            Formula = $formula
            HasFormula = ($formula -match "^=")
        } | ConvertTo-Json -Depth 10
    }
    catch {
        Write-Error "Failed to get cell value: $_"
        exit 1
    }
}

function Write-ImmediateWindow {
    param(
        [string]$WorkbookName,
        [string]$Expression
    )
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    $excel = $result.Excel
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        
        # Add WinAPI type if not already defined
        if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    public const int SW_RESTORE = 9;
}
"@
        }
        
        $vbe = $excel.VBE
        $immediateWindow = $null
        
        foreach ($win in $vbe.Windows) {
            if ($win.Type -eq 5) {
                $immediateWindow = $win
                break
            }
        }
        
        if (-not $immediateWindow) {
            Write-Error "Immediate Window not found"
            exit 1
        }
        
        $vbe.MainWindow.Visible = $true
        $immediateWindow.Visible = $true
        
        $vbeHwnd = [IntPtr]$vbe.MainWindow.HWnd
        [void][WinAPI]::ShowWindow($vbeHwnd, [WinAPI]::SW_RESTORE)
        [void][WinAPI]::SetForegroundWindow($vbeHwnd)
        Start-Sleep -Milliseconds 500
        
        [System.Windows.Forms.SendKeys]::SendWait("^g")
        Start-Sleep -Milliseconds 500
        
        $escapedExpression = $Expression.Replace('{', '{{').Replace('}', '}}').Replace('+', '{+}').Replace('^', '{^}').Replace('%', '{%}').Replace('~', '{~}').Replace('(', '{(}').Replace(')', '{)}').Replace('[', '{[}').Replace(']', '{]}')
        [System.Windows.Forms.SendKeys]::SendWait("?$escapedExpression{ENTER}")
        
        Start-Sleep -Milliseconds 500
        
        @{
            Success = $true
            Expression = $Expression
        } | ConvertTo-Json -Depth 10
    }
    catch {
        Write-Error "Failed to evaluate expression: $_"
        exit 1
    }
}

function Get-ImmediateWindow {
    param([string]$WorkbookName)
    
    $result = Find-WorkbookInInstances -WorkbookName $WorkbookName
    $workbook = $result.Workbook
    $excel = $result.Excel
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        
        # Add WinAPI type if not already defined
        if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    public const int SW_RESTORE = 9;
}
"@
        }
        
        $vbe = $excel.VBE
        $immediateWindow = $null
        
        foreach ($win in $vbe.Windows) {
            if ($win.Type -eq 5) {
                $immediateWindow = $win
                break
            }
        }
        
        if (-not $immediateWindow) {
            @{
                Success = $true
                Content = ""
                LineCount = 0
            } | ConvertTo-Json -Depth 10
            return
        }
        
        $vbe.MainWindow.Visible = $true
        $immediateWindow.Visible = $true
        
        $vbeHwnd = [IntPtr]$vbe.MainWindow.HWnd
        [void][WinAPI]::ShowWindow($vbeHwnd, [WinAPI]::SW_RESTORE)
        [void][WinAPI]::SetForegroundWindow($vbeHwnd)
        Start-Sleep -Milliseconds 300
        
        [System.Windows.Forms.SendKeys]::SendWait("^g")
        Start-Sleep -Milliseconds 300
        
        [System.Windows.Forms.Clipboard]::Clear()
        
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 300
        
        $content = [System.Windows.Forms.Clipboard]::GetText()
        
        [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}")
        
        [System.Windows.Forms.Clipboard]::Clear()
        
        $lineCount = 0
        if ($content) {
            $lineCount = ($content -split "`r`n|`r|`n").Count
        }
        
        @{
            Success = $true
            Content = $content
            LineCount = $lineCount
        } | ConvertTo-Json -Depth 10
    }
    catch {
        Write-Error "Failed to read Immediate Window: $_"
        exit 1
    }
}

# Command line argument parsing
if ($args.Count -gt 0) {
    $command = $args[0]
    
    switch ($command) {
        "GetWorkbooks" {
            Get-Workbooks
        }
        "GetModules" {
            if ($args.Count -lt 2) {
                Write-Error "Please specify workbook name"
                exit 1
            }
            Get-VBAModules -WorkbookName $args[1]
        }
        "GetCode" {
            if ($args.Count -lt 3) {
                Write-Error "Please specify workbook name and module name"
                exit 1
            }
            Get-VBACode -WorkbookName $args[1] -ModuleName $args[2]
        }
        "AddModule" {
            if ($args.Count -lt 3) {
                Write-Error "Please specify workbook name and module name"
                exit 1
            }
            $moduleType = if ($args.Count -ge 4) { [int]$args[3] } else { 1 }
            Add-VBAModule -WorkbookName $args[1] -ModuleName $args[2] -ModuleType $moduleType
        }
        "SetCode" {
            if ($args.Count -lt 4) {
                Write-Error "Please specify workbook name, module name, and code"
                exit 1
            }
            Set-VBACode -WorkbookName $args[1] -ModuleName $args[2] -Code $args[3]
        }
        "RunMacro" {
            if ($args.Count -lt 3) {
                Write-Error "Please specify workbook name and macro name"
                exit 1
            }
            $params = if ($args.Count -ge 4) { $args[3..($args.Count-1)] } else { @() }
            Invoke-VBAMacro -WorkbookName $args[1] -MacroName $args[2] -Parameters $params
        }
        "GetSheetNames" {
            if ($args.Count -lt 2) {
                Write-Error "Please specify workbook name"
                exit 1
            }
            Get-SheetNames -WorkbookName $args[1]
        }
        "GetRangeValues" {
            if ($args.Count -lt 4) {
                Write-Error "Please specify workbook name, sheet name, and range address"
                exit 1
            }
            Get-RangeValues -WorkbookName $args[1] -SheetName $args[2] -RangeAddress $args[3]
        }
        "GetCellValue" {
            if ($args.Count -lt 4) {
                Write-Error "Please specify workbook name, sheet name, and cell address"
                exit 1
            }
            Get-CellValue -WorkbookName $args[1] -SheetName $args[2] -CellAddress $args[3]
        }
        "GetImmediateWindow" {
            if ($args.Count -lt 2) {
                Write-Error "Please specify workbook name"
                exit 1
            }
            Get-ImmediateWindow -WorkbookName $args[1]
        }
        "WriteImmediateWindow" {
            if ($args.Count -lt 3) {
                Write-Error "Please specify workbook name and expression"
                exit 1
            }
            Write-ImmediateWindow -WorkbookName $args[1] -Expression $args[2]
        }
        default {
            Write-Error "Unknown command: $command"
            exit 1
        }
    }
}
