# Prompt Structure Improvement - v2.0.0

## 改善の背景

テスト中に発見されたモジュール名の制約（アルファベットで始める必要がある）は、エラーキャプチャパターン固有の問題ではなく、**VBAの一般的なルール**です。

ユーザーからのフィードバック：
> モジュール名は必ずアルファベットで始めるのはエラーハンドリングのルールではなく、モジュール追加時の共通したルールとしたほうがよいと思います。

## 実施した改善

### 構造の再編成

**変更前:**
```
## 1. Error Handling Pattern (エラーキャプチャ)
   - モジュール名はアルファベットで始める ← ここに記載
   
## 2. Immediate Window Operations (イミディエイト操作)
```

**変更後:**
```
## 1. General Rules for VBA Operations (全般的なルール) ← 新設
   ├── Module Naming Constraints (モジュール名の制約)
   └── Timing Considerations (タイミングの考慮事項)
   
## 2. Error Handling Pattern (エラーキャプチャ)
   ├── Section 1を参照（モジュール名）
   └── Section 1を参照（タイミング）
   
## 3. Immediate Window Operations Protocol (イミディエイト操作)

## 4. General Guidelines (一般的なガイドライン)

## 5. Best Practices for High-Quality Assistance (ベストプラクティス)
```

---

## Section 1: General Rules for VBA Operations（新設）

### 1.1 Module Naming Constraints

**記載内容:**

1. **アルファベットで始める必要がある**
   - ✅ Valid: `MyModule`, `ErrorCapture_123`, `Utils`
   - ❌ Invalid: `_ErrorCapture`, `123Module`, `-Test`

2. **スペースや特殊文字は使用不可**（アンダースコアは可）
   - ✅ Valid: `My_Module`, `TestModule2`
   - ❌ Invalid: `My Module`, `Test-Module`, `Module#1`

3. **最大31文字**

**重要性の説明:**
- Excelが無効なモジュール名を拒否
- `run_macro` が "マクロが見つかりません" で失敗
- `add_module` が成功してもモジュールが使用不可になる

**ベストプラクティス例:**
```typescript
// ✅ CORRECT: Start with letter
const moduleName = `ErrorCapture_${Date.now()}`;
const moduleName = `Temp_${userId}`;

// ❌ WRONG: Don't start with underscore or number
const moduleName = `_ErrorCapture_${Date.now()}`;
const moduleName = `${Date.now()}_Module`;
```

### 1.2 Timing Considerations

**記載内容:**

`add_module` + `edit_vba` の順次実行について：
- Excel内部でのモジュール登録に時間が必要
- `edit_vba` を早く呼びすぎると "インデックスが有効範囲にありません" エラー
- 適切なエラーハンドリング推奨：

```typescript
try {
    await add_module(workbook, moduleName);
    await edit_vba(workbook, moduleName, code);
} catch (error) {
    // 失敗時は空モジュールが残る可能性
    // リトライまたはユーザーに通知
}
```

---

## Section 2: Error Handling Pattern（改善）

### 変更内容

1. **Section 1への参照追加:**
   ```typescript
   // 1. Generate unique module name (following VBA naming rules - see Section 1)
   const moduleName = `ErrorCapture_${timestamp}`;  // ✅ Starts with letter
   ```

2. **タイミング考慮の参照:**
   ```typescript
   // 4. Execute the pattern (with proper error handling - see Section 1)
   try {
       await add_module(workbook, moduleName);
       await edit_vba(workbook, moduleName, wrapperCode);
   ```

3. **Common Issuesセクション削除:**
   - Issue 1 (モジュール名) → Section 1に移動
   - Issue 2 (タイミング) → Section 1に移動
   - Issue 3 (空モジュール) → Section 1に統合

4. **例の修正:**
   ```
   変更前: `ErrorCapture_1711467845` (starts with letter!)
   変更後: `ErrorCapture_1711467845` (following naming rules from Section 1)
   ```

---

## 変更されたファイル

### `src/prompts/guidelines-advanced.ts`

**追加:**
- Line 16-65: Section 1 "General Rules for VBA Operations" 全体（新設）
  - Module Naming Constraints (19-51行)
  - Timing Considerations (53-65行)

**変更:**
- Line 67: Section番号を1→2に変更
- Line 75: コメント修正（Section 1参照）
- Line 84: コメント修正（Section 1参照）
- Line 90-110: try-catchのコメント簡素化（Section 1参照）
- Line 113-125: Common Issues セクション削除
- Line 137: 例のコメント修正（Section 1参照）
- Line 196: Section番号を2→3に変更
- Line 299: Section番号を3→4に変更
- Line 323: Section番号を4→5に変更

---

## 検証結果

### 構造の正しさ
```
✅ Section 1: General Rules for VBA Operations
  ✅ Module Naming Constraints subsection
  ✅ Naming rule documented  
  ✅ Timing Considerations subsection

✅ Section 2: Error Handling Pattern
  ✅ References Section 1 for naming rules
  ✅ References Section 1 for timing

✅ Section 3: Immediate Window Operations Protocol

✅ Section 4: General Guidelines

✅ Section 5: Best Practices for High-Quality Assistance

✅ No duplicate section numbers
```

### サイズ
- **変更前:** 9,451 文字
- **変更後:** 10,032 文字
- **増加量:** +581 文字 (+6.1%)

Section 1の新設により増加しましたが、情報が整理され重複が削減されました。

---

## 改善効果

### 1. 情報の論理的な整理

**Before:**
- エラーキャプチャのセクションにVBA全般のルールが混在
- 読者がエラーキャプチャ固有の制約と誤解する可能性

**After:**
- VBA全般のルール → Section 1
- エラーキャプチャ固有のパターン → Section 2
- 明確な責任分離

### 2. 再利用性の向上

**Before:**
- モジュール名の制約がエラーキャプチャセクションのみに記載
- 他のユースケースで参照できない

**After:**
- Section 1に一般ルールとして独立
- どこからでも参照可能
- 将来的な拡張が容易

### 3. AIの理解向上

**Before:**
```typescript
// ⚠️ IMPORTANT: Module name MUST start with a letter
const moduleName = `ErrorCapture_${timestamp}`;
```
→ なぜこのルールが必要か不明確

**After:**
```typescript
// Generate unique module name (following VBA naming rules - see Section 1)
const moduleName = `ErrorCapture_${timestamp}`;
```
→ Section 1を読めば理由が明確

### 4. メンテナンス性の向上

- VBA全般のルール変更 → Section 1のみ修正
- エラーキャプチャパターン改善 → Section 2のみ修正
- 変更の影響範囲が明確

---

## ユーザーフィードバックの反映

### 元の要望
> モジュール名は必ずアルファベットで始めるのはエラーハンドリングのルールではなく、モジュール追加時の共通したルールとしたほうがよい

### 実装結果
✅ **完全に反映**
- Section 1 "General Rules for VBA Operations" として独立
- エラーハンドリングセクションからは参照のみ
- 全てのモジュール追加操作に適用される一般ルールとして明確化

---

## 結論

**Promptsの構造が大幅に改善されました！**

- ✅ VBA全般のルールを適切に分離
- ✅ 論理的な階層構造
- ✅ セクション間の参照が明確
- ✅ 情報の重複を削減
- ✅ 将来の拡張が容易
- ✅ AIの理解と適用が向上

この構造により、新しいパターンや制約を追加する際も、適切なセクションに配置できます。
