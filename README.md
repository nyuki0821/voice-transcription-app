# Voice Transcription App

音声ファイルを文字起こしし、情報を抽出・分析するGoogleAppsScriptアプリケーション。

## 機能概要

- Zoom通話録音ファイルの自動処理
- AssemblyAIを使用した高精度な文字起こし
- OpenAIを使用した会話内容の分析・要約
- 定期的な通知・レポート機能
- スプレッドシートへの結果保存

## プロジェクト構成

```
voice-transcription-app/
├── FileProcessor.js         # ファイル処理モジュール
├── InformationExtractor.js  # 情報抽出モジュール
├── Logger.js                # ロギングモジュール
├── Main.js                  # メインコントローラー
├── NotificationService.js   # 通知サービスモジュール
├── SpreadsheetManager.js    # スプレッドシート操作モジュール
├── TranscriptionService.js  # 文字起こしサービスモジュール
├── Utilities.js             # ユーティリティ関数モジュール
├── ZoomAPIManager.js        # Zoom API管理モジュール
├── ZoomPhoneTriggersSetup.js # Zoomフォントリガー設定
├── ZoomphoneProcessor.js    # Zoomフォン処理モジュール
├── ZoomphoneService.js      # Zoomフォンサービスモジュール
├── appsscript.json          # AppsScript設定ファイル
└── test.js                  # テストスクリプト
```

## セットアップ方法

### 前提条件
- Google アカウント
- AssemblyAI APIキー
- OpenAI APIキー（オプション）
- Clasp CLI（ローカル開発用）

### インストール手順

1. リポジトリをクローン
   ```
   git clone https://github.com/nyuki0821/voice-transcription-app.git
   cd voice-transcription-app
   ```

2. 依存関係のインストール
   ```
   npm install
   ```

3. Claspでログイン（初回のみ）
   ```
   npx clasp login
   ```

4. スクリプトをデプロイ
   ```
   npx clasp push
   ```

5. スクリプトのプロパティを設定
   - スクリプトエディタから「プロジェクトの設定」→「スクリプトのプロパティ」
   - 以下のプロパティを設定:
     - `SPREADSHEET_ID`: データ管理用スプレッドシートID
     - その他必要な設定値はスプレッドシートの「システム設定」シートから管理

## 使用方法

1. スプレッドシートの「システム設定」シートでAPI設定と処理設定を行う
2. トリガーを設定して定期実行（`setupTriggers()`関数を実行）
3. 指定したGoogleドライブフォルダに音声ファイルを配置
4. 自動的に文字起こし・分析が実行され結果がスプレッドシートに保存
5. 指定時間に処理結果のサマリーがメール通知

## ライセンス

ISC

## 注意事項

- API使用量と料金に注意してください
- 個人情報を含む会話の処理には適切なセキュリティ対策を講じてください

## ドキュメント

主要なセットアップ・運用ガイドは docs/zoom_phone_integration_hybrid.md を参照してください。 