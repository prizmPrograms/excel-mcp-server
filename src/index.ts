#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { ExcelWrapper } from './excel-wrapper.js';

const excel = new ExcelWrapper();

const server = new Server(
  {
    name: 'excel-mcp-server',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'get_workbooks',
        description: '現在起動しているExcelの全ワークブック一覧を取得します',
        inputSchema: {
          type: 'object',
          properties: {},
          required: [],
        },
      },
      {
        name: 'list_modules',
        description: '指定したワークブック内の全VBAモジュール一覧を取得します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名（例: "Book1.xlsx"）',
            },
          },
          required: ['workbook'],
        },
      },
      {
        name: 'get_module_code',
        description: '指定したモジュールのVBAコードを取得します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            module_name: {
              type: 'string',
              description: 'モジュール名',
            },
          },
          required: ['workbook', 'module_name'],
        },
      },
      {
        name: 'add_module',
        description: '新しいVBAモジュールを追加します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            module_name: {
              type: 'string',
              description: '新しいモジュール名',
            },
            module_type: {
              type: 'number',
              description: 'モジュールタイプ（1=標準モジュール、2=クラスモジュール、3=フォーム）',
              default: 1,
            },
          },
          required: ['workbook', 'module_name'],
        },
      },
      {
        name: 'edit_vba',
        description: '既存のVBAモジュールのコードを編集（全体を置換）します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            module_name: {
              type: 'string',
              description: 'モジュール名',
            },
            code: {
              type: 'string',
              description: '新しいVBAコード',
            },
          },
          required: ['workbook', 'module_name', 'code'],
        },
      },
      {
        name: 'run_macro',
        description: '指定したマクロを実行します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            macro_name: {
              type: 'string',
              description: 'マクロ名（例: "Module1.MyMacro"）',
            },
            parameters: {
              type: 'array',
              description: 'マクロへの引数（オプション）',
              items: {
                type: 'string',
              },
            },
          },
          required: ['workbook', 'macro_name'],
        },
      },
      {
        name: 'run_macro_safe',
        description: 'エラーキャプチャ付きでマクロを実行します。コンパイルエラー・実行時エラーの種別と発生行をJSONで返すため、Copilotによる自律修正ループに使用してください。**このツールを呼び出す前に、必ずユーザーへ次のメッセージを表示すること: 「自律修正モードで実行します。エラーダイアログが表示された場合は必ず［終了］ボタンを押してください（［デバッグ］は押さないでください）。」**',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名（例: "Book1.xlsx"）',
            },
            macro_name: {
              type: 'string',
              description: 'マクロ名（例: "Module1.MyMacro"）',
            },
          },
          required: ['workbook', 'macro_name'],
        },
      },
      {
        name: 'get_sheet_names',
        description: 'ワークブック内のシート名一覧を取得します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
          },
          required: ['workbook'],
        },
      },
      {
        name: 'get_range_values',
        description: '指定したシートの範囲のセル値を2次元配列で取得します。VBAが書き込んだデータの読み出しに使えます。',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            sheet: {
              type: 'string',
              description: 'シート名',
            },
            range: {
              type: 'string',
              description: '範囲アドレス（例: "A1:C10"）',
            },
          },
          required: ['workbook', 'sheet', 'range'],
        },
      },
      {
        name: 'get_cell_value',
        description: '指定したセルの値・数式を取得します',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名',
            },
            sheet: {
              type: 'string',
              description: 'シート名',
            },
            cell: {
              type: 'string',
              description: 'セルアドレス（例: "A1"）',
            },
          },
          required: ['workbook', 'sheet', 'cell'],
        },
      },
      {
        name: 'read_immediate_window',
        description: 'VBAイミディエイトウィンドウの内容を読み取ります。Debug.Print出力やテスト結果の確認に使用します。**このツールを呼び出す前に、必ずユーザーへ次のメッセージを表示すること: 「イミディエイトウィンドウ操作を実行します。Excelウィンドウがアクティブになり、キーボード操作（Ctrl+G、Ctrl+A、Ctrl+C）が送信され、クリップボードが使用されます。」**',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名（例: "Book1.xlsx"）',
            },
          },
          required: ['workbook'],
        },
      },
      {
        name: 'write_immediate_window',
        description: 'VBAイミディエイトウィンドウで式を評価します。変数の値確認やテストコードの実行に使用します。一時マクロを生成して式を評価し、結果をDebug.Printで出力します。**このツールを呼び出す前に、必ずユーザーへ次のメッセージを表示すること: 「イミディエイトウィンドウ操作を実行します。Excelウィンドウがアクティブになり、キーボード操作（Ctrl+G、式の入力、Enter）が送信されます。」**',
        inputSchema: {
          type: 'object',
          properties: {
            workbook: {
              type: 'string',
              description: 'ワークブック名（例: "Book1.xlsx"）',
            },
            expression: {
              type: 'string',
              description: '評価する式（例: "myVariable", "1+1", "MyFunction(5)"）',
            },
          },
          required: ['workbook', 'expression'],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'get_workbooks': {
        const workbooks = await excel.getWorkbooks();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(workbooks, null, 2),
            },
          ],
        };
      }

      case 'list_modules': {
        const { workbook } = args as { workbook: string };
        const modules = await excel.listModules(workbook);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(modules, null, 2),
            },
          ],
        };
      }

      case 'get_module_code': {
        const { workbook, module_name } = args as {
          workbook: string;
          module_name: string;
        };
        const code = await excel.getModuleCode(workbook, module_name);
        return {
          content: [
            {
              type: 'text',
              text: `Module: ${code.ModuleName}\nLines: ${code.LineCount}\n\n${code.Code}`,
            },
          ],
        };
      }

      case 'add_module': {
        const { workbook, module_name, module_type = 1 } = args as {
          workbook: string;
          module_name: string;
          module_type?: number;
        };
        const result = await excel.addModule(workbook, module_name, module_type);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'edit_vba': {
        const { workbook, module_name, code } = args as {
          workbook: string;
          module_name: string;
          code: string;
        };
        const result = await excel.setModuleCode(workbook, module_name, code);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'run_macro': {
        const { workbook, macro_name, parameters = [] } = args as {
          workbook: string;
          macro_name: string;
          parameters?: any[];
        };
        const result = await excel.runMacro(workbook, macro_name, parameters);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'run_macro_safe': {
        const { workbook, macro_name } = args as {
          workbook: string;
          macro_name: string;
        };
        const result = await excel.runMacroSafe(workbook, macro_name);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'get_sheet_names': {
        const { workbook } = args as { workbook: string };
        const sheets = await excel.getSheetNames(workbook);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(sheets, null, 2),
            },
          ],
        };
      }

      case 'get_range_values': {
        const { workbook, sheet, range } = args as {
          workbook: string;
          sheet: string;
          range: string;
        };
        const result = await excel.getRangeValues(workbook, sheet, range);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'get_cell_value': {
        const { workbook, sheet, cell } = args as {
          workbook: string;
          sheet: string;
          cell: string;
        };
        const result = await excel.getCellValue(workbook, sheet, cell);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'read_immediate_window': {
        const { workbook } = args as { workbook: string };
        const result = await excel.getImmediateWindow(workbook);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      case 'write_immediate_window': {
        const { workbook, expression } = args as {
          workbook: string;
          expression: string;
        };
        const result = await excel.evaluateInImmediate(workbook, expression);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${errorMessage}`,
        },
      ],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Excel MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
