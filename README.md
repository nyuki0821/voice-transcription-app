# Voice Transcription App

音声ファイルを文字起こしし、情報を抽出・分析するGoogleAppsScriptアプリケーション。

## 機能概要

- Zoom通話録音ファイルの自動処理
- AssemblyAIを使用した高精度な文字起こし
- OpenAIを使用した会話内容の分析・要約
- 定期的な通知・レポート機能
- スプレッドシートへの結果保存
- クラウドファンクションによるZoom WebhookおよびAPI連携

## スケジュール

- **録音ファイル取得（Fetch）:** 平日・土日祝日問わず6:00~24:00の間、30分ごとに実行
- **文字起こし処理:** 平日・土日祝日問わず6:00~24:00の間、10分ごとに実行
- **夜間バッチ処理:** 毎朝6:15に実行

## 処理フロー

1. Zoom通話の録音ファイルが生成されると、クラウドファンクションによってRecordingsシートに記録されます
2. 30分ごとにRecordingsシートに記録された音声ファイルの中から、timestampが直近2時間かつまだ取得していないrecording_idを対象に未処理フォルダへfetchします
3. 取得（fetch）が完了したものはRecordingsシートのステータス（Status列）を更新します
4. 10分ごとに最大バッチサイズ10で未処理フォルダにあるファイルを文字起こし処理します
5. 文字起こし処理に成功したファイルは完了フォルダに移動し、エラーが発生した場合はエラーフォルダに移動します
6. 文字起こし中に処理が停止した場合、処理中フォルダに残ったファイルは未処理フォルダへ自動的に戻されます

## プロジェクト構成

```
voice-transcription-app/
├── src/                        # ソースコードディレクトリ
│   ├── main/                   # メインコントローラー
│   │   ├── Main.js             # アプリケーションのエントリーポイント
│   │   └── ZoomPhoneTriggersSetup.js # トリガー設定モジュール
│   ├── core/                   # コア機能
│   │   ├── FileProcessor.js    # ファイル処理モジュール
│   │   ├── Logger.js           # ロギングモジュール
│   │   ├── NotificationService.js # 通知サービスモジュール
│   │   ├── SpreadsheetManager.js # スプレッドシート操作モジュール
│   │   └── Utilities.js        # ユーティリティ関数モジュール
│   ├── zoom/                   # Zoom関連モジュール
│   │   ├── ZoomAPIManager.js   # Zoom API管理モジュール
│   │   ├── ZoomphoneProcessor.js # Zoomフォン処理モジュール
│   │   └── ZoomphoneService.js # Zoomフォンサービスモジュール
│   ├── transcription/          # 文字起こし関連モジュール
│   │   ├── TranscriptionService.js # 文字起こしサービスモジュール
│   │   └── InformationExtractor.js # 情報抽出モジュール
│   ├── config/                 # 設定関連
│   │   └── EnvironmentConfig.js # 環境設定モジュール
│   ├── RetentionCleaner.js     # データ保持期間管理モジュール
│   └── appsscript.json         # AppsScript設定ファイル
├── cloud_function/             # GCPクラウドファンクション
│   ├── index.js                # Zoom Webhook処理関数
│   └── package.json            # 依存関係設定
├── docs/                       # ドキュメント
│   └── zoom_phone_integration_hybrid.md # セットアップガイド
└── zoom-webhook-signature-verification.js # Webhookシグネチャ検証ユーティリティ
```

## セットアップ方法

### 前提条件
- Google アカウント
- Google Cloud Platform プロジェクト
- AssemblyAI APIキー
- OpenAI APIキー（オプション）
- Zoom開発者アカウント
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

5. クラウドファンクション設定（Zoom Webhook用）
   ```
   cd cloud_function
   gcloud functions deploy zoomWebhook --runtime nodejs18 --trigger-http
   ```

6. スクリプトのプロパティを設定
   - スクリプトエディタから「プロジェクトの設定」→「スクリプトのプロパティ」
   - 以下のプロパティを設定:
     - `SPREADSHEET_ID`: データ管理用スプレッドシートID
     - その他必要な設定値はスプレッドシートの「システム設定」シートから管理

## 使用方法

1. スプレッドシートの「システム設定」シートでAPI設定と処理設定を行う
2. トリガーを設定して定期実行（`setupTriggers()`関数を実行）
3. Zoom Phone APIおよびWebhook連携を設定（詳細は`docs/zoom_phone_integration_hybrid.md`を参照）
4. 自動的に録音ファイルが処理され、文字起こし・分析が実行され結果がスプレッドシートに保存
5. 指定時間に処理結果のサマリーがメール通知

## ライセンス

ISC

## 注意事項

- API使用量と料金に注意してください
- 個人情報を含む会話の処理には適切なセキュリティ対策を講じてください
- 長期間のデータ保持には`RetentionCleaner.js`の設定を適切に行ってください

## ドキュメント

主要なセットアップ・運用ガイドは docs/zoom_phone_integration_hybrid.md を参照してください。 