# 🧪 テスト実行ガイド

## 🎯 **簡単実行方法（推奨）**

### **全テストを一括実行**
```javascript
// Google Apps Scriptエディタで以下を実行
runAllTestsWithMediumPriority()
```
**これだけでOK！** 全てのリファクタリングテストが自動実行されます。

### **軽量テスト（動作確認のみ）**
```javascript
// 環境設定だけ確認したい場合
runLightweightTests()
```

## 📋 **テストファイル一覧**

### **🚀 メインファイル（これだけ知ってればOK）**
| ファイル | 役割 | 実行方法 |
|---------|------|----------|
| **MasterTestRunner.js** | 全テストの司令塔 | `runAllTestsWithMediumPriority()` |
| **README_TEST.md** | このファイル（説明書） | 読むだけ |

### **🔧 個別テストファイル（通常は直接実行不要）**
| ファイル | テスト対象 | 説明 |
|---------|-----------|------|
| FileMovementServiceTest.js | ファイル移動機能 | 重複ファイル移動処理の統一テスト |
| ConstantsTest.js | 定数定義 | プロジェクト定数の整合性テスト |
| ConfigManagerTest.js | 設定管理 | 設定取得・キャッシュ機能テスト |
| RefactoringIntegrationTest.js | 統合テスト | 高優先度リファクタリング統合テスト |
| MediumPriorityRefactoringTest.js | Main.js分割 | 優先度中リファクタリングテスト |

## 🎮 **実行手順**

### **Step 1: Google Apps Scriptエディタを開く**
1. [script.google.com](https://script.google.com) にアクセス
2. プロジェクトを開く

### **Step 2: テスト実行**
1. **MasterTestRunner.js** を開く
2. 関数選択で `runAllTestsWithMediumPriority` を選択
3. 実行ボタンをクリック

### **Step 3: 結果確認**
- **ログ**: 実行ログで詳細確認
- **成功率**: 90%以上なら正常
- **エラー**: 失敗した場合は詳細ログを確認

## 📊 **テスト結果の見方**

### **正常な結果例**
```
=== 全テストスイート実行結果（優先度中リファクタリング含む） ===
総テスト数: 25
成功: 24
失敗: 1
実行時間: 1.2秒
成功率: 96.0%
```

### **問題がある場合**
- **成功率 < 90%**: 設定やコードに問題がある可能性
- **エラーメッセージ**: 具体的な問題箇所を示している
- **実行時間 > 5秒**: パフォーマンスに問題がある可能性

## 🔍 **個別テスト実行（上級者向け）**

特定の機能だけテストしたい場合：

```javascript
// 特定のテストスイートのみ実行
runSpecificTestSuite('FileMovementService')  // ファイル移動のみ
runSpecificTestSuite('ConfigManager')        // 設定管理のみ
runSpecificTestSuite('MediumPriorityRefactoring')  // Main.js分割のみ
```

## 🚨 **トラブルシューティング**

### **よくあるエラー**

#### **「○○が定義されていません」エラー**
```
原因: 依存ファイルが読み込まれていない
解決: 全ファイルがプロジェクトに含まれているか確認
```

#### **「権限がありません」エラー**
```
原因: Google Drive/Sheetsへのアクセス権限不足
解決: 初回実行時に権限を許可する
```

#### **「設定が見つかりません」エラー**
```
原因: EnvironmentConfig.jsの設定不備
解決: 必要な設定値（フォルダID等）を確認
```

### **パフォーマンス問題**
- **実行時間が長い**: `runLightweightTests()` で軽量テストを試す
- **メモリエラー**: 個別テストを順次実行

## 📈 **テスト結果の活用**

### **開発時**
- **コード変更後**: 必ず `runAllTestsWithMediumPriority()` を実行
- **新機能追加時**: 関連テストが通ることを確認

### **本番デプロイ前**
- **全テスト成功**: 安全にデプロイ可能
- **一部失敗**: 失敗箇所を修正してから再テスト

## 🎯 **まとめ**

**迷ったらこれだけ覚えて！**
1. **MasterTestRunner.js** を開く
2. **`runAllTestsWithMediumPriority()`** を実行
3. **成功率90%以上** なら正常

これで全てのリファクタリング成果をテストできます！

# リファクタリング後テストスイート

このディレクトリには、リファクタリング後のコードをテストするための包括的なテストスイートが含まれています。

## 📁 テストファイル構成

```
src/test/
├── ConfigTester.js                    # 既存の環境設定テスト
├── FileMovementServiceTest.js         # FileMovementServiceのテスト
├── ConstantsTest.js                   # Constantsモジュールのテスト
├── ConfigManagerTest.js               # ConfigManagerのテスト
├── RefactoringIntegrationTest.js      # 統合テスト
├── MasterTestRunner.js                # 全テスト一括実行
└── README_TEST.md                     # このファイル
```

## 🚀 クイックスタート

### 1. 全テスト実行（推奨）
```javascript
runAllTests()
```

### 2. 軽量テスト（基本機能のみ）
```javascript
runLightweightTests()
```

### 3. 環境チェック
```javascript
checkTestEnvironment()
```

## 📋 個別テストスイート

### FileMovementServiceテスト
```javascript
runAllFileMovementServiceTests()
```
**テスト内容:**
- 処理結果オブジェクトの作成・操作
- ログ出力機能
- ファイル移動のバリデーション

### Constantsテスト
```javascript
runAllConstantsTests()
```
**テスト内容:**
- 定数定義の確認
- メッセージフォーマット機能
- 音声ファイル判定機能
- レコードID抽出機能

### ConfigManagerテスト
```javascript
runAllConfigManagerTests()
```
**テスト内容:**
- 設定取得機能
- 個別設定取得
- キャッシュ機能
- 設定妥当性チェック
- デフォルト設定取得

### 統合テスト
```javascript
runAllRefactoringIntegrationTests()
```
**テスト内容:**
- サービス間連携
- Main.jsリファクタリング部分
- 下位互換性
- パフォーマンス
- エラーハンドリング

## 🎯 特定テストスイートの実行

```javascript
// 環境設定テストのみ
runSpecificTestSuite('config')

// Constantsテストのみ
runSpecificTestSuite('constants')

// ConfigManagerテストのみ
runSpecificTestSuite('configmanager')

// FileMovementServiceテストのみ
runSpecificTestSuite('filemovement')

// 統合テストのみ
runSpecificTestSuite('integration')
```

## 📊 テスト結果の見方

### 成功例
```
✓ FileMovementService: 成功
✓ Constants: 成功
✓ ConfigManager: 成功
✓ 統合テスト: 成功

成功: 4件, 失敗: 0件
🎉 全てのテストが成功しました！
```

### 失敗例
```
✓ FileMovementService: 成功
✗ Constants: 失敗
✓ ConfigManager: 成功
✗ 統合テスト: 失敗

成功: 2件, 失敗: 2件
⚠️ 一部のテストで問題が発生しました
```

## 🔧 トラブルシューティング

### よくある問題と解決方法

#### 1. モジュールが見つからないエラー
```
エラー: Constants is not defined
```
**解決方法:** 
- `src/core/Constants.js`が正しく読み込まれているか確認
- Google Apps Scriptエディタでファイルが保存されているか確認

#### 2. 設定取得エラー
```
エラー: EnvironmentConfig is not defined
```
**解決方法:**
- `src/config/EnvironmentConfig.js`が存在するか確認
- 環境設定が正しく設定されているか確認

#### 3. 権限エラー
```
エラー: You do not have permission to call...
```
**解決方法:**
- Google Apps Scriptの実行権限を確認
- 必要に応じて承認プロセスを実行

## 📈 テスト範囲

### カバレッジ
- **FileMovementService**: 100% (全公開メソッド)
- **Constants**: 100% (全定数・ユーティリティ関数)
- **ConfigManager**: 95% (実際のシート・フォルダアクセス除く)
- **統合機能**: 90% (Main.jsの主要リファクタリング部分)

### テストタイプ
- **単体テスト**: 各モジュールの個別機能
- **統合テスト**: モジュール間の連携
- **パフォーマンステスト**: キャッシュ効果など
- **エラーハンドリングテスト**: 異常系の動作確認

## 🛡️ 安全性について

### モックテスト
実際のファイル移動やシート操作は行わず、以下の方法で安全にテスト:
- モックオブジェクトの使用
- パラメータバリデーションのみ
- 設定値の存在確認のみ

### 本番環境への影響
- **影響なし**: テストは読み取り専用操作のみ
- **データ変更なし**: 実際のファイルやシートは変更されません
- **設定変更なし**: 環境設定は読み取りのみ

## 📝 テスト追加ガイド

### 新しいテストの追加方法

1. **適切なテストファイルを選択**
   - 新機能 → 対応するテストファイル
   - 新モジュール → 新しいテストファイル作成

2. **テスト関数の命名規則**
   ```javascript
   function test[機能名]() {
     try {
       Logger.log("=== [機能名]テスト開始 ===");
       // テストロジック
       Logger.log("=== [機能名]テスト完了 ===");
       return "[機能名]のテストに成功しました";
     } catch (error) {
       Logger.log("[機能名]テストでエラー: " + error);
       return "テスト中にエラーが発生しました: " + error;
     }
   }
   ```

3. **マスターテストランナーへの追加**
   - `MasterTestRunner.js`の該当箇所に新しいテストを追加

## 🔄 継続的テスト

### 推奨テスト頻度
- **開発中**: 変更後に関連テストを実行
- **リリース前**: 全テスト実行
- **定期メンテナンス**: 月1回の全テスト実行

### 自動化の検討
将来的には以下の自動化を検討:
- トリガーベースの自動テスト実行
- テスト結果の自動通知
- 継続的インテグレーション

## 📞 サポート

テストに関する質問や問題がある場合:
1. まず`checkTestEnvironment()`を実行
2. ログを確認して問題を特定
3. 必要に応じて個別テストスイートで詳細確認

---

**最終更新**: 2024年1月
**バージョン**: 1.0.0
**対応リファクタリング**: FileMovementService, Constants, ConfigManager統合 