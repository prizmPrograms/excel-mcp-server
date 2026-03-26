import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export interface Workbook {
  Name: string;
  FullName: string;
  Path: string;
  Sheets: number;
  HasVBProject: boolean;
}

export interface VBAModule {
  Name: string;
  Type: number;
  CodeModule: {
    CountOfLines: number;
  };
}

export interface VBACode {
  ModuleName: string;
  Code: string;
  LineCount: number;
}

export interface SheetInfo {
  Name: string;
  Index: number;
  Visible: number;
}

export interface RangeValues {
  WorkbookName: string;
  SheetName: string;
  RangeAddress: string;
  RowCount: number;
  ColumnCount: number;
  Values: (string | number | boolean | null)[][];
}

export interface CellValue {
  WorkbookName: string;
  SheetName: string;
  Address: string;
  Value: string | number | boolean | null;
  Text: string;
  Formula: string;
  HasFormula: boolean;
}

export interface ImmediateWindowContent {
  Success: boolean;
  Content?: string;
  LineCount?: number;
}

export interface ImmediateWindowResult {
  Success: boolean;
  Expression?: string;
}

export class ExcelWrapper {
  private scriptPath: string;

  constructor() {
    this.scriptPath = join(__dirname, 'excel-com.ps1');
  }

  private async executePowerShell(command: string, args: string[] = []): Promise<string> {
    return new Promise((resolve, reject) => {
      const allArgs = [
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-File', this.scriptPath,
        command,
        ...args
      ];

      const ps = spawn('powershell.exe', allArgs, {
        windowsHide: true
      });

      let stdout = '';
      let stderr = '';

      ps.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      ps.stderr.on('data', (data) => {
        stderr += data.toString();
      });

      ps.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`PowerShell error: ${stderr || stdout}`));
        } else {
          resolve(stdout.trim());
        }
      });

      ps.on('error', (err) => {
        reject(new Error(`Failed to start PowerShell: ${err.message}`));
      });
    });
  }

  async getWorkbooks(): Promise<Workbook[]> {
    const output = await this.executePowerShell('GetWorkbooks');
    if (!output) {
      return [];
    }
    return JSON.parse(output);
  }

  async listModules(workbookName: string): Promise<VBAModule[]> {
    const output = await this.executePowerShell('GetModules', [workbookName]);
    if (!output) {
      return [];
    }
    return JSON.parse(output);
  }

  async getModuleCode(workbookName: string, moduleName: string): Promise<VBACode> {
    const output = await this.executePowerShell('GetCode', [workbookName, moduleName]);
    return JSON.parse(output);
  }

  async addModule(workbookName: string, moduleName: string, moduleType: number = 1): Promise<{ Success: boolean; ModuleName: string; Type: number }> {
    const output = await this.executePowerShell('AddModule', [
      workbookName,
      moduleName,
      moduleType.toString()
    ]);
    return JSON.parse(output);
  }

  async setModuleCode(workbookName: string, moduleName: string, code: string): Promise<{ Success: boolean; ModuleName: string; LinesAdded: number }> {
    const output = await this.executePowerShell('SetCode', [
      workbookName,
      moduleName,
      code
    ]);
    return JSON.parse(output);
  }

    async runMacro(workbookName: string, macroName: string, parameters: any[] = []): Promise<{ Success: boolean; MacroName: string; Result: any }> {
    const args = [workbookName, macroName];
    if (parameters.length > 0) {
      // Pass each parameter as a separate argument so PowerShell receives them
      // as individual array elements (not a single JSON-serialized string).
      args.push(...parameters.map(String));
    }
    const output = await this.executePowerShell('RunMacro', args);
    return JSON.parse(output);
  }

  async getSheetNames(workbookName: string): Promise<SheetInfo[]> {
    const output = await this.executePowerShell('GetSheetNames', [workbookName]);
    if (!output) return [];
    return JSON.parse(output);
  }

  async getRangeValues(workbookName: string, sheetName: string, rangeAddress: string): Promise<RangeValues> {
    const output = await this.executePowerShell('GetRangeValues', [workbookName, sheetName, rangeAddress]);
    return JSON.parse(output);
  }

  async getCellValue(workbookName: string, sheetName: string, cellAddress: string): Promise<CellValue> {
    const output = await this.executePowerShell('GetCellValue', [workbookName, sheetName, cellAddress]);
    return JSON.parse(output);
  }

  async getImmediateWindow(workbookName: string): Promise<ImmediateWindowContent> {
    const output = await this.executePowerShell('GetImmediateWindow', [workbookName]);
    return JSON.parse(output);
  }

  async evaluateInImmediate(workbookName: string, expression: string): Promise<ImmediateWindowResult> {
    const output = await this.executePowerShell('WriteImmediateWindow', [workbookName, expression]);
    return JSON.parse(output);
  }
}
