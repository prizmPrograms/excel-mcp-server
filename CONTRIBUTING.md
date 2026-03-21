# 貢献ガイド

Excel MCP Server への貢献に興味を持っていただきありがとうございます！

## 貢献方法

### バグ報告

バグを見つけた場合は、[Issues](https://github.com/prizmPrograms/excel-mcp-server/issues) で報告してください。

報告する際は以下の情報を含めてください：
- Windows のバージョン
- Node.js のバージョン
- Excel のバージョン
- 再現手順
- 期待される動作
- 実際の動作
- エラーメッセージ（あれば）

### 機能リクエスト

新機能のアイデアがある場合も、[Issues](https://github.com/prizmPrograms/excel-mcp-server/issues) で提案してください。

以下を含めると議論がスムーズになります：
- 機能の説明
- ユースケース
- 既存の機能との関連性

### プルリクエスト

プルリクエストは歓迎します！

#### 手順

1. **フォーク**
   ```bash
   # GitHubでリポジトリをフォーク
   ```

2. **クローン**
   ```bash
   git clone https://github.com/あなたのユーザー名/excel-mcp-server.git
   cd excel-mcp-server
   ```

3. **ブランチ作成**
   ```bash
   git checkout -b feature/your-feature-name
   ```

4. **依存関係のインストール**
   ```bash
   npm install
   ```

5. **変更を加える**
   - コードの変更
   - テストの追加（該当する場合）
   - ドキュメントの更新

6. **ビルド確認**
   ```bash
   npm run build
   ```

7. **コミット**
   ```bash
   git add .
   git commit -m "Add: your feature description"
   ```

8. **プッシュ**
   ```bash
   git push origin feature/your-feature-name
   ```

9. **プルリクエスト作成**
   - GitHubでプルリクエストを作成
   - 変更内容を説明

## コーディング規約

### TypeScript

- 型定義を明示的に記述
- async/await を使用（Promiseのチェーンは避ける）
- エラーハンドリングを適切に行う

### PowerShell

- 関数名は動詞-名詞形式（Get-VBACode など）
- コメントは英語で記述（文字化け防止）
- エラー時は Write-Error を使用
- JSON 出力は ConvertTo-Json で統一

### コミットメッセージ

以下のプレフィックスを使用：
- `Add:` 新機能追加
- `Fix:` バグ修正
- `Update:` 既存機能の更新
- `Docs:` ドキュメントのみの変更
- `Refactor:` リファクタリング
- `Test:` テスト追加・修正

例：
```
Add: get_workbook_properties function
Fix: parameter handling in Invoke-VBAMacro
Update: README with new examples
Docs: improve installation instructions
```

## 開発環境

### 必要なもの

- Windows 10/11
- Node.js 18+
- Microsoft Excel
- Git
- PowerShell 5.1+

### セットアップ

```bash
# クローン
git clone https://github.com/prizmPrograms/excel-mcp-server.git
cd excel-mcp-server

# 依存関係インストール
npm install

# ビルド
npm run build

# ウォッチモード（開発時）
npm run watch
```

### ディレクトリ構成

```
excel-mcp-server/
├── src/
│   ├── index.ts           # MCPサーバーのエントリーポイント
│   ├── excel-wrapper.ts   # PowerShell呼び出しラッパー
│   └── excel-com.ps1      # Excel COM操作スクリプト
├── dist/                  # ビルド出力（git管理外）
├── package.json
├── tsconfig.json
├── README.md
└── LICENSE
```

## PowerShell スクリプトの編集

`src/excel-com.ps1` を編集する際の注意点：

1. **文字エンコーディング**: UTF-8（BOMなし）で保存
2. **コメント**: 必ず英語で記述（日本語は文字化けの原因）
3. **出力**: `ConvertTo-Json` の前に `return` を付けない
4. **エラー**: `Write-Error` でエラーメッセージを出力し `exit 1`

### 関数追加の例

```powershell
function Get-NewFeature {
    param(
        [string]$WorkbookName,
        [string]$Parameter
    )
    
    $excel = Get-ExcelInstance
    $workbook = $excel.Workbooks | Where-Object { $_.Name -eq $WorkbookName } | Select-Object -First 1
    
    if (-not $workbook) {
        Write-Error "Workbook not found: $WorkbookName"
        exit 1
    }
    
    try {
        # Your logic here
        $result = "Your result"
        
        @{
            Success = $true
            Result = $result
        } | ConvertTo-Json
    }
    catch {
        Write-Error "Error: $_"
        exit 1
    }
}
```

## TypeScript コードの編集

`src/index.ts` にツール定義を追加：

```typescript
{
  name: "new_feature",
  description: "機能の説明",
  inputSchema: {
    type: "object",
    properties: {
      workbook: {
        type: "string",
        description: "ワークブック名"
      },
      parameter: {
        type: "string",
        description: "パラメータの説明"
      }
    },
    required: ["workbook", "parameter"]
  }
}
```

## テスト

現在、自動テストは実装されていません。
手動テストを行う場合：

1. Excelを起動
2. ワークブックを開く
3. MCPサーバーを起動
4. Claude Desktop等のMCPクライアントで機能を実行

## ライセンス

貢献したコードは MIT License でライセンスされます。

## 質問

わからないことがあれば、[Issues](https://github.com/prizmPrograms/excel-mcp-server/issues) で質問してください！
