# GitHub Actions CI導入手順書

## 概要
このドキュメントでは、voice-transcription-appプロジェクトにGitHub Actionsを使用したCI（継続的インテグレーション）プロセスを導入する手順を説明します。

## 前提条件
- GitHubリポジトリへのアクセス権限
- Google Apps Scriptプロジェクトの設定済み
- Node.js環境（ローカル開発用）

## 1. 基本的なCI設定

### 1.1 ワークフローディレクトリの確認
既に `.github/workflows/` ディレクトリが存在することを確認済みです。

### 1.2 メインCIワークフローの作成
以下のワークフローファイルを作成します：

**ファイル名**: `.github/workflows/ci.yml`

```yaml
name: CI Pipeline

on:
  push:
    branches: [ main, develop, feature/* ]
  pull_request:
    branches: [ main, develop ]

jobs:
  test:
    runs-on: ubuntu-latest
    
    strategy:
      matrix:
        node-version: [16.x, 18.x]
    
    steps:
    - name: Checkout repository
      uses: actions/checkout@v4
      
    - name: Setup Node.js ${{ matrix.node-version }}
      uses: actions/setup-node@v4
      with:
        node-version: ${{ matrix.node-version }}
        cache: 'npm'
        
    - name: Install dependencies
      run: npm ci
      
    - name: Run linting
      run: npm run lint
      continue-on-error: true
      
    - name: Run tests
      run: npm test
      
    - name: Upload test results
      uses: actions/upload-artifact@v4
      if: always()
      with:
        name: test-results-${{ matrix.node-version }}
        path: test-results/

  clasp-validation:
    runs-on: ubuntu-latest
    needs: test
    
    steps:
    - name: Checkout repository
      uses: actions/checkout@v4
      
    - name: Setup Node.js
      uses: actions/setup-node@v4
      with:
        node-version: '18.x'
        cache: 'npm'
        
    - name: Install dependencies
      run: npm ci
      
    - name: Install clasp
      run: npm install -g @google/clasp
      
    - name: Validate clasp configuration
      run: |
        echo "Validating .clasp.json configuration"
        cat .clasp.json
        
    - name: Check for syntax errors
      run: |
        echo "Checking JavaScript syntax in src directory"
        find src -name "*.js" -exec node -c {} \;
```

## 2. テスト環境の設定

### 2.1 package.jsonの更新
現在のpackage.jsonにテストスクリプトとlintingを追加します：

```json
{
  "scripts": {
    "test": "node src/test/MasterTestRunner.js",
    "test:unit": "node src/test/unit-tests.js",
    "test:integration": "node src/test/integration-tests.js",
    "lint": "eslint src/**/*.js",
    "lint:fix": "eslint src/**/*.js --fix",
    "validate": "node scripts/validate-gas-syntax.js"
  },
  "devDependencies": {
    "@types/google-apps-script": "^1.0.97",
    "eslint": "^8.0.0",
    "eslint-config-google": "^0.14.0"
  }
}
```

### 2.2 ESLint設定ファイルの作成
**ファイル名**: `.eslintrc.js`

```javascript
module.exports = {
  env: {
    browser: true,
    es2021: true,
    googleappsscript: true
  },
  extends: [
    'google'
  ],
  parserOptions: {
    ecmaVersion: 12,
    sourceType: 'script'
  },
  globals: {
    // Google Apps Script globals
    Logger: 'readonly',
    DriveApp: 'readonly',
    GmailApp: 'readonly',
    SpreadsheetApp: 'readonly',
    UrlFetchApp: 'readonly',
    Utilities: 'readonly',
    PropertiesService: 'readonly',
    ScriptApp: 'readonly'
  },
  rules: {
    'max-len': ['error', { code: 120 }],
    'require-jsdoc': 'warn',
    'valid-jsdoc': 'warn'
  }
};
```

## 3. 高度なCI設定

### 3.1 デプロイメントワークフロー
**ファイル名**: `.github/workflows/deploy.yml`

```yaml
name: Deploy to Google Apps Script

on:
  push:
    branches: [ main ]
  workflow_dispatch:

jobs:
  deploy:
    runs-on: ubuntu-latest
    environment: production
    
    steps:
    - name: Checkout repository
      uses: actions/checkout@v4
      
    - name: Setup Node.js
      uses: actions/setup-node@v4
      with:
        node-version: '18.x'
        cache: 'npm'
        
    - name: Install dependencies
      run: npm ci
      
    - name: Install clasp
      run: npm install -g @google/clasp
      
    - name: Setup clasp credentials
      run: |
        echo '${{ secrets.CLASP_CREDENTIALS }}' > ~/.clasprc.json
        
    - name: Deploy to Google Apps Script
      run: |
        clasp push
        clasp deploy --description "Automated deployment from GitHub Actions"
```

### 3.2 プルリクエスト用ワークフロー
**ファイル名**: `.github/workflows/pr-checks.yml`

```yaml
name: Pull Request Checks

on:
  pull_request:
    types: [opened, synchronize, reopened]

jobs:
  code-quality:
    runs-on: ubuntu-latest
    
    steps:
    - name: Checkout repository
      uses: actions/checkout@v4
      with:
        fetch-depth: 0
        
    - name: Setup Node.js
      uses: actions/setup-node@v4
      with:
        node-version: '18.x'
        cache: 'npm'
        
    - name: Install dependencies
      run: npm ci
      
    - name: Run comprehensive tests
      run: |
        npm run test
        npm run lint
        
    - name: Check for TODO/FIXME comments
      run: |
        echo "Checking for TODO/FIXME comments..."
        grep -r "TODO\|FIXME" src/ || echo "No TODO/FIXME found"
        
    - name: Validate file structure
      run: |
        echo "Validating project structure..."
        test -f .clasp.json || (echo ".clasp.json not found" && exit 1)
        test -d src/ || (echo "src directory not found" && exit 1)
        echo "Project structure validation passed"
        
    - name: Comment PR with results
      uses: actions/github-script@v7
      if: always()
      with:
        script: |
          const { owner, repo, number } = context.issue;
          await github.rest.issues.createComment({
            owner,
            repo,
            issue_number: number,
            body: '🚀 CI checks completed! Please review the results above.'
          });
```

## 4. シークレット設定

### 4.1 必要なシークレット
GitHubリポジトリの Settings > Secrets and variables > Actions で以下を設定：

1. **CLASP_CREDENTIALS**: clasp認証情報
   ```bash
   # ローカルで実行して認証情報を取得
   clasp login
   cat ~/.clasprc.json
   ```

2. **GAS_SCRIPT_ID**: Google Apps ScriptのスクリプトID
   ```
   .clasp.jsonから取得
   ```

### 4.2 環境変数の設定
**ファイル名**: `.github/workflows/env-template.yml`

```yaml
env:
  NODE_ENV: 'test'
  GAS_SCRIPT_ID: ${{ secrets.GAS_SCRIPT_ID }}
  NOTIFICATION_EMAIL: ${{ secrets.NOTIFICATION_EMAIL }}
```

## 5. 導入手順

### 5.1 ステップ1: 基本設定
```bash
# 1. 依存関係のインストール
npm install --save-dev eslint eslint-config-google

# 2. package.jsonの更新（上記の内容を反映）

# 3. ESLint設定ファイルの作成
```

### 5.2 ステップ2: ワークフローファイルの作成
```bash
# 1. CIワークフローファイルの作成
# .github/workflows/ci.yml

# 2. デプロイワークフローファイルの作成
# .github/workflows/deploy.yml

# 3. PRチェックワークフローファイルの作成
# .github/workflows/pr-checks.yml
```

### 5.3 ステップ3: シークレット設定
1. GitHubリポジトリの設定画面でシークレットを追加
2. clasp認証情報の設定
3. 必要な環境変数の設定

### 5.4 ステップ4: テスト実行
```bash
# 1. ローカルでテスト実行
npm test
npm run lint

# 2. GitHubにプッシュしてCI動作確認
git add .
git commit -m "Add GitHub Actions CI configuration"
git push origin feature/github-actions-ci
```

## 6. 監視とメンテナンス

### 6.1 ワークフロー実行状況の確認
- GitHub Actions タブで実行状況を確認
- 失敗時の通知設定
- ログの確認とデバッグ

### 6.2 定期的なメンテナンス
- 依存関係の更新
- ワークフローの最適化
- テストカバレッジの向上

## 7. トラブルシューティング

### 7.1 よくある問題
1. **clasp認証エラー**
   - シークレットの設定確認
   - 認証情報の更新

2. **テスト実行エラー**
   - Node.jsバージョンの確認
   - 依存関係の確認

3. **デプロイエラー**
   - Google Apps Scriptの権限確認
   - スクリプトIDの確認

### 7.2 デバッグ方法
```yaml
# ワークフローでのデバッグ出力
- name: Debug information
  run: |
    echo "Node version: $(node --version)"
    echo "NPM version: $(npm --version)"
    echo "Current directory: $(pwd)"
    ls -la
```

## 8. 参考資料
- [GitHub Actions Documentation](https://docs.github.com/en/actions)
- [Google Apps Script clasp](https://github.com/google/clasp)
- [ESLint Configuration](https://eslint.org/docs/user-guide/configuring/)

---

**注意**: この設定は基本的な構成です。プロジェクトの要件に応じて調整してください。 