# Excel MCP Server

Windows環境で起動中のExcelインスタンスに対して、VBAモジュールの追加・編集・実行を可能にするMCPサーバーです。

## 機能

### コアツール

- **get_workbooks**: 現在起動しているExcelの全ワークブック一覧を取得
- **list_modules**: ワークブック内の全VBAモジュール一覧を取得
- **get_module_code**: 指定したモジュールのVBAコードを取得
- **add_module**: 新しいVBAモジュールを追加
- **edit_vba**: 既存のVBAモジュールのコードを編集
- **run_macro**: 指定したマクロを実行
- **get_sheet_names**: ワークブック内のシート名一覧を取得
- **get_range_values**: 指定した範囲のセル値を2次元配列で取得
- **get_cell_value**: 指定したセルの値・数式を取得
- **read_immediate_window**: VBAイミディエイトウィンドウの内容を読み取り（Debug.Print出力の確認）⚠️
- **write_immediate_window**: VBAイミディエイトウィンドウで式を評価（変数値の確認やテストコード実行）⚠️

### AIアシスタント向けガイドライン（Prompts機能）

Excel MCP Serverは、AIアシスタントに対して使用方法のガイドラインを提供します：

- **excel-vba-guidelines**: エラーキャプチャパターン、イミディエイト操作の許可プロトコルなどの詳細なガイドライン

このガイドラインはMCPサーバー自体が提供するため、ユーザーによる追加設定は不要です。

### ⚠️ 注意が必要なツール

以下のツールはVBEウィンドウをアクティブにし、キーボード操作を送信するため、実行中は他の作業が中断される可能性があります。AIアシスタントは使用前にユーザーに確認します：

- **read_immediate_window**: VBEウィンドウをアクティブにし、Ctrl+G、Ctrl+A、Ctrl+Cのキー操作を送信し、クリップボードを使用します
- **write_immediate_window**: VBEウィンドウをアクティブにし、Ctrl+Gのキー操作、式の入力、Enterキーを送信します

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

### 方法1: npx で直接使用（推奨・最も簡単）

クローン不要で、すぐに使えます：

**Claude Desktop の設定ファイル**（`%APPDATA%\Claude\claude_desktop_config.json`）に追加：

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["github:prizmPrograms/excel-mcp-server"]
    }
  }
}
```

**メリット**:
- ✅ インストール不要（ローカルにファイルを残さない）
- ✅ 常に最新版を使用
- ✅ 最も簡単

**注意**: 初回起動時は npx がパッケージをダウンロードするため少し時間がかかりますが、2回目以降はキャッシュが使われて高速になります。

---

### 方法2: ローカルにクローンして使用

より高速な起動やオフライン使用が必要な場合：

#### 1. リポジトリをクローン

```bash
git clone https://github.com/prizmPrograms/excel-mcp-server.git
cd excel-mcp-server
```

または、[Releases](https://github.com/prizmPrograms/excel-mcp-server/releases)からZIPファイルをダウンロードして展開することもできます。

#### 2. 依存関係をインストール

```bash
npm install
```

#### 3. ビルド

```bash
npm run build
```

## 使い方

### 方法1を使用した場合（npx）

設定ファイルに追加後、Claude Desktop を再起動するだけです。自動的に起動します。

### 方法2を使用した場合（ローカルクローン）

#### MCPサーバーとして実行

```bash
node dist/index.js
```

#### MCP クライアント設定

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

### マクロの実行
```
run_macro
- workbook: "Book1.xlsx"
- macro_name: "Module1.MyMacro"
```

### エラーキャプチャパターン（デバッグ時）

AIアシスタントは、マクロのデバッグ時に自動的にエラーキャプチャパターンを使用します。このパターンは、一時的なVBAラッパー関数を生成してエラー情報をキャプチャします。

**パターンの仕組み:**
1. 一時モジュール（例: `_ErrorCapture_1234567890`）を作成
2. エラーキャプチャ用のラッパー関数を追加
3. `run_macro`でラッパー関数を実行
4. JSON形式でエラー情報を取得: `{"status":"error","number":13,"description":"Type mismatch"...}`
5. エラー情報を基にコードを修正
6. 一時モジュールをクリーンアップ

**主なVBAエラー番号**:
- `5`: 無効なプロシージャ呼び出し
- `6`: オーバーフロー
- `9`: インデックスが有効範囲にありません（配列）
- `11`: 0で除算
- `13`: 型が一致しません
- `91`: オブジェクト変数が設定されていません
- `424`: オブジェクトが必要です
- `1004`: アプリケーション定義またはオブジェクト定義のエラー

詳細な使用方法は、AIアシスタントが提供する`excel-vba-guidelines`プロンプトに記載されています。

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

## AIによるVBA自動デバッグ

AIアシスタントは、マクロのエラー時に自動的にデバッグと修正を行います：

**エラーキャプチャパターンによる自動修正フロー:**

```
1. エラーキャプチャパターン実行 → 詳細なエラー情報取得（ErrorNumber + ErrorDescription）
2. get_module_code → 問題のVBAコードを取得
3. AI がエラーを解析して修正コードを生成
4. edit_vba → 修正コードをモジュールに書き込む
5. エラーキャプチャパターン再実行 → 成功するまでループ
```

詳細な実装方法は、AIアシスタントが利用する`excel-vba-guidelines`プロンプトに記載されています。

## 移行ガイド（v2.0.0）

### 🔄 run_macro_safe の削除について

**v2.0.0で変更:** `run_macro_safe`ツールは削除されました。

**代替方法:** AIアシスタントは、MCPの**Prompts機能**で提供される`excel-vba-guidelines`プロンプトに基づき、エラーキャプチャパターンを実装します。

このパターンは以下のツールを組み合わせて実現されます：
- `add_module`: 一時モジュール `_ErrorCapture_<timestamp>` の作成
- `edit_vba`: エラーキャプチャラッパー関数の追加
- `run_macro`: ラッパー関数の実行でJSON形式のエラー情報を取得

**変更の利点:**
- ✅ **透明性**: 何が起きているか明確（一時モジュールが作成される）
- ✅ **柔軟性**: 必要に応じてカスタマイズ可能
- ✅ **保守性**: ガイドラインはMCPサーバーに含まれ、自動更新される
- ✅ **拡張性**: 将来的に他のガイドラインも追加可能

**ユーザーへの影響:**
- 既存のツール（add_module, edit_vba, run_macro）はすべて引き続き利用可能
- AIアシスタントの動作は変わらず、より明確になります
- 追加の設定は不要（`npx excel-mcp-server`で自動的に利用可能）
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

## 制限事項

### 複数のExcelプロセスについて

このMCPサーバーは、**単一のExcelプロセス内のすべてのワークブック**にアクセスできます。

ただし、複数のExcel.exeプロセスが同時に起動している場合（Excelを複数の独立したウィンドウで起動した場合）、Windows COMの制限により、通常は**最初に起動したExcelプロセスのみ**がアクセス可能です。

#### 推奨される使用方法

- すべてのExcelファイルを同じExcelインスタンス（同じプロセス）で開いてください
- 通常、Excelファイルをダブルクリックして開くと、既存のExcelプロセスで開かれます
- 新しいExcelプロセスは、明示的に「新しいウィンドウ」を開いた場合のみ起動します

#### 複数プロセスが起動している場合

もし複数のExcelプロセスが起動していて、特定のブックにアクセスできない場合：

1. 対象のブックを閉じる
2. 既存のExcelウィンドウから「ファイル → 開く」で再度開く
3. これにより、同じプロセス内で開かれ、アクセス可能になります

## AIによる自律的なテスト/デバッグワークフロー

これらの機能を組み合わせて、AIが自動的にVBAコードをテスト・検証できます：

### パターン1: 自動テスト実行

```
1. edit_vba → テスト対象の関数を作成
2. edit_vba → テストマクロを作成（Debug.Printで結果を出力）
3. run_macro → テストマクロを実行
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

