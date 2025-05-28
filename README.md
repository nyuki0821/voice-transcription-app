# Voice Transcription App

Zoom通話録音ファイルを文字起こしし、情報を抽出・分析するGoogleAppsScriptとクラウドファンクションを組み合わせたハイブリッドアプリケーション。

## 機能概要

- Zoom通話録音ファイルの自動処理
- Webhookによるリアルタイムメタ情報収集
- 定期的な録音ファイル取得
- AssemblyAIを使用した高精度な文字起こし
- OpenAIを使用した会話内容の分析・要約
- スプレッドシートへの結果一元管理

## アーキテクチャ全体像

```
                   recording_completed
    Zoom Phone -------------------------→ Cloud Functions
                                           (zoom-webhook-handler)
                                                   |
                                                   ↓
                                          Google Sheets
                                          (Recordings)

    [GAS Project]
    ┌───────────────────────────────────────────────┐
    │                                               │
    │  30 min Trigger                               │
    │  ↓                                            │
    │  ZoomphoneProcessor                           │
    │  └─→ download & save                          │
    │       ↓                                       │
    │  Google Drive                                 │
    │  (録音フォルダ)                               │
    │                                               │
    │  10 min Trigger                               │
    │  ↓                                            │
    │  Transcription Pipeline                       │
    │  └─→ Process & Analyze                        │
    │       └─→ Write to Sheets                     │
    │                                               │
    │  Weekly Retention Cleanup                     │
    │  └─→ Remove old files                         │
    └───────────────────────────────────────────────┘
```

| コンポーネント | 主な役割 | 実装/ファイル |
|---------------|----------|-------------|
| Cloud Functions | Webhook 受信／署名検証／Sheets 書込 | `index.js` (Node 18) |
| Google Sheets   | メタ情報一覧 & ログ | `Recordings` シート |
| GAS: ZoomphoneProcessor | 録音リスト取得 & Drive 保存 | `src/zoom/ZoomphoneProcessor.js` |
| GAS: ZoomphoneService   | Zoom API 呼び出しヘルパ | `src/zoom/ZoomphoneService.js` |
| GAS: Main & triggers    | 文字起こし・要約 | `src/main/*.js` |

## スケジュールとフロー別の役割分担

| タイミング | フロー | 実行場所 | 詳細 |
|------------|--------|----------|------|
| 録音完了時 即時 | Webhook → メタ記録 | **Cloud Functions** | HMAC 検証 / Sheets へ 1 行追記 (録音 ID, DL URL, 時刻, 電話番号, Duration等) |
| 30 分おき | 録音ファイル DL | **GAS `checkAndFetchZoomRecordings`** | Zoom API で 直近 2h をリスト → 未取得を Drive 保存 |
| 10 分おき | 文字起こし & 要約 | **GAS `processBatchOnSchedule`** | AssemblyAI → OpenAI (任意) → Sheets に貼付 |
| 日曜 03:00 | リテンション | GAS | Drive 内 90 日超ファイル削除 |

## ログステータス管理

Recordingsシートは以下のカラムで各ファイルのステータスを一元管理します:

| カラム | 説明 |
|-------|------|
| record_id | 一意の録音ID |
| timestamp_recording | 録音生成日時 |
| download_url | 録音ファイルURL |
| call_date | 通話日 |
| call_time | 通話時間 |
| duration | 通話時間長 |
| sales_phone_number | セールス側番号 |
| customer_phone_number | 顧客側番号 |
| timestamp_fetch | 取得処理日時 |
| status_fetch | 取得ステータス |
| timestamp_transcription | 文字起こし処理日時 |
| status_transcription | 文字起こしステータス |
| process_start | 処理開始時間 |
| process_end | 処理終了時間 |

## 処理フロー詳細

1. **Webhook受信と記録**
   - Zoom通話の録音ファイルが生成されると、クラウドファンクションによってRecordingsシートに記録
   - `record_id`, `download_url`, `timestamp_recording`等の情報を記録

2. **定期的なファイル取得**
   - 30分ごとにRecordingsシートから、timestampが直近2時間かつまだ取得していないrecording_idを対象に
   - ZoomAPIを使用してファイルをDrive上の未処理フォルダへfetch
   - 取得状況は`status_fetch`, `timestamp_fetch`に記録

3. **文字起こし処理**
   - 10分ごとに最大バッチサイズ10で未処理フォルダにあるファイルを文字起こし
   - 文字起こし状況は`status_transcription`, `timestamp_transcription`に記録
   - 処理時間は`process_start`, `process_end`に記録
   - 処理結果はcall_recordsシートに保存

4. **ファイル管理**
   - 文字起こし成功: 完了フォルダに移動
   - エラー発生: エラーフォルダに移動
   - 処理中断: 処理中フォルダから未処理フォルダへ自動復旧
   - 定期クリーニング: 90日超のファイルを自動削除

## プロジェクト構成

```
voice-transcription-app/
├── src/                        # ソースコードディレクトリ
│   ├── main/                   # メインコントローラー
│   │   ├── Main.js             # アプリケーションのエントリーポイント
│   │   ├── TriggerManager.js   # トリガー管理モジュール
│   │   └── ZoomphoneTriggersWrapper.js # 古いトリガー設定のラッパー
│   ├── core/                   # コア機能
│   │   ├── FileProcessor.js    # ファイル処理モジュール
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
└── zoom-webhook-signature-verification.js # Webhookシグネチャ検証ユーティリティ
```

## セットアップ方法

### 前提条件

- Google アカウント
- Google Cloud Platform プロジェクト
- Zoom開発者アカウント (`phone:read:recordings` スコープ 付き S2S OAuth アプリ)
- AssemblyAI APIキー
- OpenAI APIキー（オプション）
- Clasp CLI（ローカル開発用）

### セットアップ手順

1. **スプレッドシートセットアップ**
   - Google Drive → 新規 → Google スプレッドシート → タイトル `Zoom Voice Recordings` を作成
   - シートタブ名を `Recordings` に変更し、必要なヘッダーを設定
   - Recordingsシートのスプレッドシート ID を控える (`RECORDINGS_SHEET_ID` として使用)

2. **GCP プロジェクト作成**
   - 新規プロジェクト `zoom-webhook-relay` を作成
   - 以下のAPIを有効化:
     - Cloud Functions API
     - Cloud Build API
     - Google Sheets API

3. **サービスアカウント設定**
   - サービスアカウント `zoom-sheets-integration` を作成
   - ロール: Viewer (閲覧者)
   - スプレッドシートの共有設定でサービスアカウントに編集者権限を付与

4. **Cloud Functions 構築**
   - `index.js` と `package.json` を準備 (cloud_functionフォルダ参照)
   - Node.js 18, HTTP トリガーでデプロイ
   - 環境変数設定:
     - `ZOOM_WEBHOOK_SECRET`: Webhookシークレット
     - `RECORDINGS_SHEET_ID`: スプレッドシートID

5. **Zoom Webhook 設定**
   - Zoom Marketplace アプリ → Event Subscriptions にCloud FunctionsのURLを設定
   - Secret Tokenとして`ZOOM_WEBHOOK_SECRET`と同じ値を設定し検証

6. **Apps Script 配備**
   - `.clasp.json` ファイルを設定
   - `clasp push --force` でコードをデプロイ

7. **トリガー設定**
   - `clasp run TriggerManager.setupAllTriggers` で全てのトリガーを一括設定

8. **スクリプトプロパティ設定**
   - `CONFIG_SPREADSHEET_ID`: 設定用スプレッドシートID
   - スプレッドシートの `settings` シートに各種APIキーや設定を記述

### 必要な設定項目

| 設定項目 | 説明 |
|---------|------|
| RECORDINGS_SHEET_ID | 録音メタデータ管理シートID |
| ASSEMBLYAI_API_KEY | AssemblyAI API キー |
| OPENAI_API_KEY | OpenAI API キー（オプション） |
| SOURCE_FOLDER_ID | 未処理音声ファイル保存フォルダID |
| PROCESSING_FOLDER_ID | 処理中フォルダID |
| COMPLETED_FOLDER_ID | 完了フォルダID |
| ERROR_FOLDER_ID | エラーフォルダID |
| ZOOM_CLIENT_ID | Zoom API クライアントID |
| ZOOM_CLIENT_SECRET | Zoom API クライアントシークレット |
| ZOOM_ACCOUNT_ID | Zoom アカウントID |
| ZOOM_WEBHOOK_SECRET | Webhook検証用シークレット |
| RETENTION_DAYS | ファイル保持日数（デフォルト90日) |

## トリガー設定・管理

### 包括的トリガー設定

#### 1. **全トリガー設定（推奨）**
```javascript
// 復旧機能も含めた全トリガーを設定（デフォルト）
TriggerManager.setupAllTriggers();
// または明示的に
TriggerManager.setupAllTriggers(true);

// Apps Script直接実行用
setupAllTriggersWithRecovery();
```

**設定されるトリガー:**
- Zoom録音取得（30分ごと）
- 文字起こし処理（10分ごと）
- Zoomトークン更新（毎日6:00）
- Recordingsシート監視（1時間ごと）
- 部分的失敗検知（毎日22:00）
- 中断ファイル復旧（5分ごと）
- PENDING復旧（2時間ごと）
- エラーファイル復旧（4時間ごと）

#### 2. **基本トリガーのみ設定**
```javascript
// 復旧機能を除いた基本機能のみ
TriggerManager.setupBasicTriggers();

// Apps Script直接実行用
setupBasicTriggersOnly();
```

**設定されるトリガー:**
- Zoom録音取得（30分ごと）
- 文字起こし処理（10分ごと）
- Zoomトークン更新（毎日6:00）
- Recordingsシート監視（1時間ごと）
- 部分的失敗検知（毎日22:00）

#### 3. **復旧トリガーのみ追加**
```javascript
// 既存のトリガーに復旧機能を追加
TriggerManager.setupRecoveryTriggersOnly();

// Apps Script直接実行用
addRecoveryTriggersOnly();
```

**追加されるトリガー:**
- 中断ファイル復旧（5分ごと）
- PENDING復旧（2時間ごと）
- エラーファイル復旧（4時間ごと）

### スプレッドシートメニューからの設定

1. **設定スプレッドシートを開く**
2. **メニューバー** → **文字起こしシステム** → **トリガー設定**
3. 以下から選択：
   - **🔧 全トリガー設定（復旧機能込み）** ← **推奨**
   - **⚙️ 基本トリガーのみ設定**
   - **🔄 復旧トリガーのみ追加**

### トリガー削除

```javascript
// 全トリガー削除
TriggerManager.deleteAllTriggers();

// 復旧トリガーのみ削除
TriggerManager.removeRecoveryTriggers();
```

### 運用シナリオ別推奨設定

#### **通常運用時**
```javascript
// 全機能を有効にして安定運用
TriggerManager.setupAllTriggers();
```

#### **問題が少ない安定期**
```javascript
// 基本機能のみで軽量運用
TriggerManager.setupBasicTriggers();
```

#### **問題が多発している時**
```javascript
// 基本設定 + 復旧機能を強化
TriggerManager.setupBasicTriggers();
TriggerManager.setupRecoveryTriggersOnly();
```

#### **メンテナンス時**
```javascript
// 全復旧処理を一括実行
runFullRecoveryProcess();
```

## 料金モデルと運用コスト

| サービス | 無料枠 | 想定消費量/日 | 超過時単価 (USD) |
|----------|--------|--------------|------------------|
| Cloud Functions (1st Gen) | 2M calls, 400k GB-sec | 1k calls, <5k GB-sec | $0.40/1M calls |
| Cloud Build | 120 build-min/day | デプロイ時のみ | $0.0034/min |
| Sheets API | 無料 | 数十 write | – |
| Apps Script | 90 min exec/day | ≈60 min | – |
| Zoom API | 100 req/日/Account | ≈50 | 超過：1 min cool-down |

※ 1 日 1,000 Webhook & 100 録音でも各無料枠以内でほぼゼロコスト。

## 運用・監視のベストプラクティス

| 観点 | 推奨策 |
|------|--------|
| エラー検知 | Apps Script の例外を Gmail / Google Chat に送信 |
| リトライ | Cloud Functions は 200 を返し Zoom のリトライを防止 |
| レート制限 | `Utilities.sleep(200)` & pageSize 制御で緩和 |
| Drive 容量 | `RETENTION_DAYS` を 90 日→60 日へ短縮するなど柔軟運用 |
| コスト監視 | Billing Alert を $10/月 で設定 |

## 注意事項

- API使用量と料金に注意してください
- 個人情報を含む会話の処理には適切なセキュリティ対策を講じてください
- 長期間のデータ保持には `RetentionCleaner.js` の設定を適切に行ってください
- 文字起こしエラー発生時はログに記録されますが、call_recordsシートには保存されません 

## エラーハンドリング・復旧機能

### 改善されたエラー検知機能

システムは以下の方法でエラーを検知・処理します：

#### 1. リアルタイムエラー検知
- **OpenAI APIエラー**: クォータ制限、APIキーエラー、レスポンスエラーを即座に検知
- **文字起こしエラー**: GPT-4o-mini、AssemblyAIの処理エラーを検知
- **部分的失敗**: エラーが発生しても処理が続行されるケースを検知

#### 2. 部分的失敗の自動検知・復旧
毎日22:00に自動実行される機能：

```javascript
// 手動実行の場合
detectAndRecoverPartialFailures();
```

**検知対象のエラーパターン:**
- `insufficient_quota` (OpenAI APIクォータ制限)
- `GPT-4o-mini API呼び出しエラー`
- `OpenAI APIからのレスポンスエラー`
- `情報抽出に失敗しました`
- `JSONの解析に失敗しました`

#### 3. エラーステータスの適切な管理

**ステータスの種類:**
- `SUCCESS`: 正常完了
- `ERROR: [エラー内容]`: 処理エラー
- `ERROR_DETECTED: [検知された問題]`: 部分的失敗を後から検知
- `RETRY`: 復旧処理対象
- `PENDING`: 処理待ち

#### 4. 自動復旧機能

**復旧処理の種類:**
1. **エラーファイル復旧**: `recoverErrorFiles()`
2. **PENDING状態リセット**: `resetPendingTranscriptions()`
3. **部分的失敗復旧**: `detectAndRecoverPartialFailures()`
4. **強制復旧**: `forceRecoverAllErrorFiles()`

### 使用方法

#### 手動でエラー検知・復旧を実行

```javascript
// 部分的失敗を検知・復旧
var result = detectAndRecoverPartialFailures();
Logger.log(result);

// エラーファイルを復旧
var result = recoverErrorFiles();
Logger.log(result);

// PENDING状態をリセット
var result = resetPendingTranscriptions();
Logger.log(result);
```

#### トリガーの設定

```javascript
// 全トリガーを設定（部分的失敗検知を含む）
TriggerManager.setupAllTriggers();

// 部分的失敗検知トリガーのみ設定
TriggerManager.setupPartialFailureDetectionTrigger();
```

### 通知機能

部分的失敗が検知された場合、管理者に自動でメール通知が送信されます。

**通知内容:**
- 検知されたレコード数
- 復旧成功・失敗件数
- 各レコードの詳細情報（Record ID、問題の種類、復旧状況）

### 設定

環境設定スプレッドシートの `settings` シートで以下を設定：

| 設定項目 | 説明 |
|---------|------|
| ADMIN_EMAILS | 通知先メールアドレス（カンマ区切り） |
| ENHANCE_WITH_OPENAI | OpenAI機能の有効/無効（`true`/`false`） |

**OpenAI APIエラー時の緊急対処:**
```
ENHANCE_WITH_OPENAI = false
```
に設定すると、OpenAI APIを使わずにAssemblyAIのみで処理を続行できます。

### トラブルシューティング

#### よくある問題と対処法

1. **OpenAI APIクォータ制限**
   - 使用量ダッシュボードで確認: https://platform.openai.com/usage
   - 請求情報を確認: https://platform.openai.com/account/billing
   - 一時的に `ENHANCE_WITH_OPENAI = false` に設定

2. **部分的失敗が多発する場合**
   - `detectAndRecoverPartialFailures()` を手動実行
   - ログでエラーパターンを確認
   - 必要に応じてAPIキーや設定を見直し

3. **ファイルが見つからない場合**
   - 各フォルダ（SOURCE, PROCESSING, COMPLETED, ERROR）を確認
   - `moveFileToErrorFolder()` でファイルの場所を特定

## セットアップ方法 