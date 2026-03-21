# Excel MCP Server

Windows環境で起動中のExcelインスタンスに対して、VBAモジュールの追加・編集・実行を可能にするMCPサーバーです。

## 機能

- **get_workbooks**: 現在起動しているExcelの全ワークブック一覧を取得
- **list_modules**: ワークブック内の全VBAモジュール一覧を取得
- **get_module_code**: 指定したモジュールのVBAコードを取得
- **add_module**: 新しいVBAモジュールを追加
- **edit_vba**: 既存のVBAモジュールのコードを編集
- **run_macro**: 指定したマクロを実行
- **get_sheet_names**: ワークブック内のシート名一覧を取得
- **get_range_values**: 指定した範囲のセル値を2次元配列で取得
- **get_cell_value**: 指定したセルの値・数式を取得
- **run_macro_safe**: エラーキャプチャ付きでマクロを実行（VBAエラー時もダイアログなしでエラー詳細を返す）
- **read_immediate_window**: VBAイミディエイトウィンドウの内容を読み取り（Debug.Print出力の確認）⚠️
- **write_immediate_window**: VBAイミディエイトウィンドウで式を評価（変数値の確認やテストコード実行）⚠️

### ⚠️ 注意が必要なツール

以下のツールはExcelウィンドウをアクティブにし、キーボード操作を送信するため、実行中は他の作業が中断される可能性があります。使用前にAIアシスタント（Copilot）が確認メッセージを表示します：

- **read_immediate_window**: Excelウィンドウをアクティブにし、Ctrl+G、Ctrl+A、Ctrl+Cのキー操作を送信し、クリップボードを使用します
- **write_immediate_window**: Excelウィンドウをアクティブにし、Ctrl+Gのキー操作、式の入力、Enterキーを送信します
- **run_macro_safe**: マクロ実行中にエラーダイアログが表示された場合は［終了］ボタンを押してください

## 前提条件

### システム要件
- Windows OS
- Node.js 18以上
- Microsoft Excel（起動している必要があります）

### Excelのセキュリティ設定

VBAプロジェクトにアクセスするには、Excelで以下の設定を有効にする必要があります：

1. Excelを開く
2. ファイル → オプション → トラストセンター → トラストセンターの設定
3. 「マクロの設定」を選択
4. **「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」** にチェックを入れる
5. OKをクリック

## インストール

### 1. リポジトリをクローン

```bash
git clone https://github.com/prizmPrograms/excel-mcp-server.git
cd excel-mcp-server
```

または、[Releases](https://github.com/prizmPrograms/excel-mcp-server/releases)からZIPファイルをダウンロードして展開することもできます。

### 2. 依存関係をインストール

```bash
npm install
```

### 3. ビルド

```bash
npm run build
```

## 使い方

### MCPサーバーとして実行

```bash
node dist/index.js
```

### MCP クライアント設定

Claude Desktop や他のMCPクライアントの設定ファイル（通常は `%APPDATA%\Claude\claude_desktop_config.json`）に以下を追加します：

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": [
        "C:\\path\\to\\excel-mcp-server\\dist\\index.js"
      ]
    }
  }
}
```

**注意**: パスは絶対パスで指定してください。

## ツールの使用例

### ワークブック一覧の取得
```
get_workbooks
```

### VBAモジュール一覧の取得
```
list_modules
- workbook: "Book1.xlsx"
```

### VBAコードの取得
```
get_module_code
- workbook: "Book1.xlsx"
- module_name: "Module1"
```

### 新しいモジュールの追加
```
add_module
- workbook: "Book1.xlsx"
- module_name: "MyModule"
- module_type: 1  # 1=標準モジュール、2=クラスモジュール
```

### VBAコードの編集
```
edit_vba
- workbook: "Book1.xlsx"
- module_name: "Module1"
- code: |
    Sub HelloWorld()
        MsgBox "Hello, World!"
    End Sub
```

### マクロの実行（エラーキャプチャ付き）
```
run_macro_safe
- workbook: "Book1.xlsx"
- macro_name: "Module1.MyMacro"
```

**成功時レスポンス**:
```json
{
  "Success": true,
  "Status": "success",
  "MacroName": "Module1.MyMacro",
  "Result": null
}
```

**エラー時レスポンス**（ダイアログなし）:
```json
{
  "Success": false,
  "Status": "error",
  "MacroName": "Module1.MyMacro",
  "ErrorMessage": "Type mismatch",
  "HResult": "0x80020005"
}
```

### シート名一覧の取得
```
get_sheet_names
- workbook: "Book1.xlsx"
```

**レスポンス例**:
```json
[
  {"Name": "Sheet1", "Index": 1, "Visible": -1},
  {"Name": "Sheet2", "Index": 2, "Visible": -1},
  {"Name": "Hidden", "Index": 3, "Visible": 0}
]
```

### セル範囲の値取得
```
get_range_values
- workbook: "Book1.xlsx"
- sheet: "Sheet1"
- range: "A1:C3"
```

**レスポンス例**:
```json
{
  "WorkbookName": "Book1.xlsx",
  "SheetName": "Sheet1",
  "RangeAddress": "A1:C3",
  "RowCount": 3,
  "ColumnCount": 3,
  "Values": [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ]
}
```

### 単一セルの値取得
```
get_cell_value
- workbook: "Book1.xlsx"
- sheet: "Sheet1"
- cell: "A1"
```

**レスポンス例**:
```json
{
  "WorkbookName": "Book1.xlsx",
  "SheetName": "Sheet1",
  "Address": "$A$1",
  "Value": 100,
  "Text": "100",
  "Formula": "=SUM(B1:B10)",
  "HasFormula": true
}
```

### イミディエイトウィンドウでの式評価
```
write_immediate_window
- workbook: "Book1.xlsx"
- expression: "arrFrom(1, 2, 3)"
```

この機能は、VBA関数の戻り値を確認したり、変数の内容を確認したりするのに便利です。
式は `?` プレフィックス付きで評価され、結果がイミディエイトウィンドウに表示されます。

### イミディエイトウィンドウの内容読み取り
```
read_immediate_window
- workbook: "Book1.xlsx"
```

**レスポンス例**:
```json
{
  "Success": true,
  "Content": "Array 1: dim=1 size=(3) lbs=(0) data=[1,2,3]\nArray 2: dim=1 size=(3) lbs=(0) data=[4,5,6]",
  "LineCount": 2
}
```

## AIによるVBA自動修正ループ

`run_macro_safe` を使うことで、AIがエラーを検知して自動的にVBAを修正できます：

```
1. run_macro_safe → エラー詳細取得（ErrorNumber + ErrorDescription）
2. get_module_code → 問題のVBAコードを取得
3. AI がエラーを解析して修正コードを生成
4. edit_vba → 修正コードをモジュールに書き込む
5. run_macro_safe → 再実行（成功するまでループ）
```

> **注意**: `_MCPHelper` という名前のVBAモジュールが対象ワークブックに自動インストールされます（MCPサーバー内部利用）。このモジュールを削除しても次回 `run_macro_safe` 実行時に自動再インストールされます。

## トラブルシューティング

### Excelが認識されない

**症状**: `get_workbooks` が空の配列を返す、または "Excel is not running" エラー

**解決方法**:
1. Excelが実際に起動しているか確認
2. 管理者権限でExcelを起動している場合、MCPサーバーも管理者権限で実行する必要があります
3. PowerShellの実行ポリシーを確認: `Get-ExecutionPolicy`（RemoteSignedまたはUnrestrictedが必要）

### VBAプロジェクトにアクセスできない

**症状**: "Programmatic access to Visual Basic Project is not trusted"

**解決方法**:
1. Excelを開く
2. ファイル → オプション → トラストセンター → トラストセンターの設定
3. 「マクロの設定」を選択
4. 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック

### イミディエイトウィンドウ操作が失敗する

**症状**: `read_immediate_window` または `write_immediate_window` がタイムアウトまたはエラー

**解決方法**:
1. Excelウィンドウが他のウィンドウに隠れていないか確認
2. VBEエディタが既に開いている場合は閉じる
3. 実行中に他のキーボード操作を行わない
4. セキュリティソフトがキーボード操作をブロックしていないか確認

### ビルドエラー

**症状**: `npm run build` がエラーを出す

**解決方法**:
1. Node.jsのバージョン確認: `node --version`（18以上が必要）
2. `node_modules` を削除して再インストール: `rm -r node_modules; npm install`
3. TypeScriptのグローバルインストール: `npm install -g typescript`

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) ファイルを参照してください。

## 作者

**prizmPrograms**

## 貢献

バグ報告や機能リクエストは [Issues](https://github.com/prizmPrograms/excel-mcp-server/issues) でお願いします。

プルリクエストも歓迎します！


### シート名一覧の取得
```
get_sheet_names
- workbook: "Book1.xlsx"
```

### セル範囲の値を取得（VBA実行結果の読み出しに活用）
```
get_range_values
- workbook: "Book1.xlsx"
- sheet: "Sheet1"
- range: "A1:C10"
```

### 単一セルの値・数式を取得
```
get_cell_value
- workbook: "Book1.xlsx"
- sheet: "Sheet1"
- cell: "A1"
```

## VBAからAIへのデータ受け渡しパターン

### パターン1: VBA Functionの戻り値を使う（単純な値）

```vba
Function GetSummary() As String
    GetSummary = "処理完了: " & Cells(1,1).Value
End Function
```

```
run_macro
- workbook: "Book1.xlsx"
- macro_name: "Module1.GetSummary"
```
→ `Result` フィールドに戻り値が入る

### パターン2: VBAがセルに書き込み → AIが読み取る（複雑なデータ）

```vba
Sub CalcAndWrite()
    Dim result As Double
    result = 複雑な計算...
    Worksheets("Output").Cells(1, 1).Value = result
End Sub
```

```
1. run_macro → Module1.CalcAndWrite
2. get_cell_value → sheet: "Output", cell: "A1"
```

## デバッグ機能

AIが自律的にVBAコードをテスト・デバッグできる機能を提供します。

### イミディエイトウィンドウの読み取り

VBAの`Debug.Print`出力やテスト結果を確認できます。

```
read_immediate_window
- workbook: "Book1"
```

**使用例:**
```vba
Sub TestFunction()
    Debug.Print "Test 1: "; AddNumbers(2, 3)
    Debug.Print "Test 2: "; AddNumbers(-1, 1)
End Sub
```

マクロ実行後に`read_immediate_window`を呼び出すと、Debug.Print出力を取得できます。

### イミディエイトウィンドウで式を評価

変数の値を確認したり、テスト式を実行したりできます。

```
write_immediate_window
- workbook: "Book1"
- expression: "AddNumbers(10, 20)"
```

**式の例:**
- `myVariable` - 変数の値を確認
- `1+1` - 簡単な計算
- `MyFunction(5)` - 関数を呼び出して結果を確認
- `Cells(1,1).Value` - セルの値を確認

**仕組み:** 一時マクロを生成して式を評価し、結果を`Debug.Print`で出力します。

## AIによる自律的なテスト/デバッグワークフロー

これらの機能を組み合わせて、AIが自動的にVBAコードをテスト・検証できます：

### パターン1: 自動テスト実行

```
1. edit_vba → テスト対象の関数を作成
2. edit_vba → テストマクロを作成（Debug.Printで結果を出力）
3. run_macro_safe → テストマクロを実行
4. read_immediate_window → テスト結果を読み取り
5. AIが結果を解析して成功/失敗を判定
6. 失敗している場合は修正して再テスト（ステップ1に戻る）
```

**テストマクロの例:**
```vba
Sub TestAddNumbers()
    Debug.Print "Test 1: "; (AddNumbers(2, 3) = 5)
    Debug.Print "Test 2: "; (AddNumbers(-1, 1) = 0)
    Debug.Print "Test 3: "; (AddNumbers(0, 0) = 0)
End Sub
```

### パターン2: 式の評価によるテスト

```
1. edit_vba → 関数を作成
2. write_immediate_window → 関数を呼び出して結果を確認
3. read_immediate_window → 結果を読み取り
4. AIが期待値と比較して判定
5. 必要に応じて修正
```

**例:**
```
write_immediate_window -expression: "AddNumbers(5, 10)"
read_immediate_window → "15" が出力されていることを確認
```

## 技術詳細

### アーキテクチャ

このMCPサーバーは以下のコンポーネントで構成されています：

1. **excel-com.ps1**: PowerShellスクリプトでExcel COMオブジェクトを直接操作
2. **excel-wrapper.ts**: PowerShellスクリプトを呼び出すTypeScriptラッパー
3. **index.ts**: MCPサーバーのメインエントリーポイント

### なぜPowerShellを使用するのか？

Node.jsから直接Windows COM オブジェクトを操作するには、ネイティブモジュール（`winax`など）が必要ですが、これらはVisual Studio Build Toolsを必要とします。PowerShellを使用することで、追加のビルドツールなしでCOM操作が可能になります。

### モジュールタイプ

- `1`: 標準モジュール（VBAコード用）
- `2`: クラスモジュール
- `3`: ユーザーフォーム

## トラブルシューティング

### Excelが認識されない

**症状**: `get_workbooks` が空の配列を返す、または "Excel is not running" エラー

**解決方法**:
1. Excelが実際に起動しているか確認
2. 管理者権限でExcelを起動している場合、MCPサーバーも管理者権限で実行する必要があります
3. PowerShellの実行ポリシーを確認: `Get-ExecutionPolicy`（RemoteSignedまたはUnrestrictedが必要）

### VBAプロジェクトにアクセスできない

**症状**: "Programmatic access to Visual Basic Project is not trusted"

**解決方法**:
1. Excelを開く
2. ファイル → オプション → トラストセンター → トラストセンターの設定
3. 「マクロの設定」を選択
4. 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック

### イミディエイトウィンドウ操作が失敗する

**症状**: `read_immediate_window` または `write_immediate_window` がタイムアウトまたはエラー

**解決方法**:
1. Excelウィンドウが他のウィンドウに隠れていないか確認
2. VBEエディタが既に開いている場合は閉じる
3. 実行中に他のキーボード操作を行わない
4. セキュリティソフトがキーボード操作をブロックしていないか確認

### ビルドエラー

**症状**: `npm run build` がエラーを出す

**解決方法**:
1. Node.jsのバージョン確認: `node --version`（18以上が必要）
2. `node_modules` を削除して再インストール: `rm -r node_modules; npm install`
3. TypeScriptのグローバルインストール: `npm install -g typescript`

### PowerShell実行ポリシーエラー

スクリプトは `-ExecutionPolicy Bypass` で実行されるため、通常は問題ありませんが、エラーが発生する場合は管理者権限でPowerShellを開き：
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 制限事項

- Windows専用（COM経由のため）
- Excelが既に起動している必要がある
- VBAプロジェクトへのアクセスにはセキュリティ設定の変更が必要

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) ファイルを参照してください。

## 作者

**prizmPrograms**

## 貢献

バグ報告や機能リクエストは [Issues](https://github.com/prizmPrograms/excel-mcp-server/issues) でお願いします。

プルリクエストも歓迎します！

