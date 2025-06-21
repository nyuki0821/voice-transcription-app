#!/bin/bash

# GitHub Actions CI セットアップスクリプト
echo "🚀 GitHub Actions CI セットアップを開始します..."

# 1. 依存関係のインストール
echo "📦 依存関係をインストールしています..."
npm install --save-dev eslint eslint-config-google

# 2. ディレクトリ構造の確認
echo "📁 プロジェクト構造を確認しています..."
if [ ! -d ".github/workflows" ]; then
    echo "❌ .github/workflows ディレクトリが見つかりません"
    exit 1
fi

if [ ! -f ".clasp.json" ]; then
    echo "❌ .clasp.json ファイルが見つかりません"
    exit 1
fi

# 3. 設定ファイルの確認
echo "⚙️  設定ファイルを確認しています..."
if [ ! -f ".eslintrc.js" ]; then
    echo "❌ .eslintrc.js ファイルが見つかりません"
    exit 1
fi

# 4. テストの実行確認
echo "🧪 テスト環境を確認しています..."
if [ ! -f "src/test/MasterTestRunner.js" ]; then
    echo "❌ MasterTestRunner.js が見つかりません"
    exit 1
fi

# 5. Lintingの実行
echo "🔍 コードの構文チェックを実行しています..."
npm run lint || echo "⚠️  Lintingで警告がありますが、継続します"

echo "✅ GitHub Actions CI セットアップが完了しました！"
echo ""
echo "次のステップ:"
echo "1. GitHubリポジトリの Settings > Secrets で以下を設定してください:"
echo "   - CLASP_CREDENTIALS: clasp認証情報"
echo "   - GAS_SCRIPT_ID: Google Apps ScriptのスクリプトID"
echo ""
echo "2. 変更をコミットしてプッシュしてください:"
echo "   git add ."
echo "   git commit -m \"Add GitHub Actions CI configuration\""
echo "   git push origin feature/github-actions-ci"
echo ""
echo "3. GitHub Actions タブでワークフローの実行状況を確認してください" 