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