# GitHub Actions CI クイックスタートガイド

## 🚀 5分で始めるCI導入

### ステップ1: 自動セットアップの実行
```bash
chmod +x scripts/setup-ci.sh
./scripts/setup-ci.sh
```

### ステップ2: GitHubシークレットの設定
1. GitHubリポジトリの `Settings` > `Secrets and variables` > `Actions` へ移動
2. 以下のシークレットを追加:

#### CLASP_CREDENTIALS の取得方法
```bash
# ローカルでclasp認証
clasp login

# 認証情報をコピー
cat ~/.clasprc.json
```
この内容をそのまま `CLASP_CREDENTIALS` シークレットに設定

#### GAS_SCRIPT_ID の取得方法
```bash
# .clasp.jsonからスクリプトIDを確認
cat .clasp.json
```
`scriptId` の値を `GAS_SCRIPT_ID` シークレットに設定

### ステップ3: ワークフローのテスト
```bash
# 変更をコミット
git add .
git commit -m "Add GitHub Actions CI configuration"

# プッシュしてCIを実行
git push origin feature/github-actions-ci
```

### ステップ4: 動作確認
1. GitHubリポジトリの `Actions` タブを確認
2. ワークフローが正常に実行されることを確認
3. プルリクエストを作成してPRチェックを確認

## 📋 利用可能なワークフロー

### 1. CI Pipeline (`.github/workflows/ci.yml`)
- **トリガー**: push, pull_request
- **実行内容**: テスト実行、Linting、構文チェック
- **対象ブランチ**: main, develop, feature/*

### 2. Deploy (`.github/workflows/deploy.yml`)
- **トリガー**: mainブランチへのpush
- **実行内容**: Google Apps Scriptへの自動デプロイ
- **前提条件**: CLASP_CREDENTIALS の設定必須

### 3. PR Checks (`.github/workflows/pr-checks.yml`)
- **トリガー**: プルリクエストの作成・更新
- **実行内容**: コード品質チェック、構造検証
- **結果**: PRにコメントで結果を通知

## 🛠️ 利用可能なnpmスクリプト

```bash
# テスト実行
npm test                    # 全テスト実行
npm run test:unit          # 単体テスト実行
npm run test:integration   # 統合テスト実行

# コード品質
npm run lint               # Linting実行
npm run lint:fix          # Linting自動修正
npm run validate          # 構文検証

# Google Apps Script
npm run push              # clasp push
npm run watch             # clasp push --watch
npm run deploy            # clasp deploy
```

## 🔧 カスタマイズ

### Node.jsバージョンの変更
`.github/workflows/ci.yml` の `matrix.node-version` を編集:
```yaml
strategy:
  matrix:
    node-version: [16.x, 18.x, 20.x]  # 必要なバージョンを追加
```

### Lintingルールの調整
`.eslintrc.js` の `rules` セクションを編集:
```javascript
rules: {
  'max-len': ['error', { code: 100 }],  # 行の長さを100文字に変更
  'require-jsdoc': 'off',               # JSDoc必須を無効化
}
```

### テスト設定の変更
`package.json` の `scripts` セクションを編集:
```json
"scripts": {
  "test": "node src/test/CustomTestRunner.js",  # カスタムテストランナー
}
```

## 🚨 トラブルシューティング

### よくあるエラーと解決方法

#### 1. `npm ci` が失敗する
```bash
# package-lock.jsonを削除して再生成
rm package-lock.json
npm install
```

#### 2. clasp認証エラー
```bash
# 認証情報を再取得
clasp logout
clasp login
cat ~/.clasprc.json  # 内容をCLASP_CREDENTIALSに設定
```

#### 3. テスト実行エラー
```bash
# テストファイルの存在確認
ls -la src/test/
# MasterTestRunner.jsが存在することを確認
```

#### 4. Linting エラー
```bash
# 自動修正を実行
npm run lint:fix
# 手動で修正が必要な場合は、エラーメッセージに従って修正
```

## 📊 ワークフロー実行状況の確認

### GitHub Actions ダッシュボード
1. リポジトリの `Actions` タブにアクセス
2. 各ワークフローの実行履歴を確認
3. 失敗した場合は、ログを確認してエラー原因を特定

### 通知設定
GitHub設定で以下を有効化:
- Actions の実行結果をメール通知
- 失敗時のSlack通知（Webhook設定が必要）

---

**ヒント**: 初回設定後は、定期的にワークフローを見直して最適化することをお勧めします。 