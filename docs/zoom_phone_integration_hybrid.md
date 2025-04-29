# Zoom Phone 録音連携 ― ハイブリッド型（Cloud Functions + GAS）完全ガイド

> v2024-06-xx – 本ドキュメントは **`docs/zoom_webhook_via_cloudfunctions.md`** と **`docs/Zoomファイル連携_手順書.md`** を統合し、
> 一連のストーリーが途切れないよう再構成した決定版です。

## 目次

1. 目的と概要
2. スプレッドシートセットアップ
3. サービスアカウント設定
4. GCP プロジェクト作成
5. 必要 API の有効化
6. Cloud Functions 構築
7. Apps Script 配備
8. トリガー設定
9. 動作確認
10. アーキテクチャ全体像
11. フロー別の役割分担
12. 前提条件 & 準備チェックリスト
13. 料金モデルと運用コスト
14. 動作確認 & テストシナリオ
15. 運用・監視のベストプラクティス
16. 更新履歴

---

## 1. 目的と概要

Zoom Phone で生成される通話録音を、

1. **リアルタイムにメタ情報を収集（Webhook）**
2. **30 分バッチで録音ファイルを Google Drive へ保存**
3. **10 分バッチで文字起こし → 要約を生成**

という 3 段階に分離し、可観測性・冗長性・コスト最適化を両立したワークフローを実現する。

## 2. スプレッドシートセットアップ

1. Google Drive → 新規 → Google スプレッドシート → タイトル **`Zoom Voice Recordings`** を作成。
2. シートタブ名を `Recordings` に変更し A1〜G1 に以下ヘッダーを入力。
   | Timestamp | RecordingId | DownloadUrl | StartTime | PhoneNumber | Duration | Status |
3. URL 内 `/d/` と `/edit` に挟まれた **ID** を控えます (`SPREADSHEET_ID` として後続で使用)。

---
## 3. GCP プロジェクト作成

1. プロジェクトセレクタ → **新しいプロジェクト** → 名称 `zoom-webhook-relay` で作成。
2. 作成後、自動で当該プロジェクトに切り替わることを確認します。

---
## 4. サービスアカウント設定

1. Cloud Console → IAM と管理 → サービスアカウント → **サービスアカウントを作成**。
2. 名前: `zoom-sheets-integration`。
3. ロール: **閲覧者 (Viewer)** だけで OK。
4. **鍵ファイルは作成しません**（組織ポリシーで禁止されていても問題なし）。
5. スプレッドシートの共有設定で **`zoom-sheets-integration@zoom-webhook-relay.iam.gserviceaccount.com`** を **編集者** 権限で追加します。
6. Cloud Functions をデプロイする際、**Runtime service account** にこの SA を指定します。

---

## 5. 必要 API の有効化

`zoom-webhook-relay` プロジェクトを開いた状態で API ライブラリ画面に進み、下記 3 つを検索して **[有効にする]** を押します。

- Cloud Functions API
- Cloud Build API
- Google Sheets API

---

## 6. Cloud Functions 構築

> ここでは環境変数を **Secret Manager を使わず直接設定** する手順を示します。

### 6.1 ソースコード準備

```bash
mkdir cf && cd cf
# package.json
cat > package.json <<'EOF'
{
  "name": "zoom-webhook-handler",
  "type": "module",
  "main": "index.js",
  "engines": { "node": ">=18" },
  "dependencies": {
    "@google-cloud/functions-framework": "^3.3.0",
    "googleapis": "^126.0.0"
  }
}
EOF
# index.js（抜粋）
cat > index.js <<'EOF'
import { google } from 'googleapis';
import crypto from 'crypto';
import functions from '@google-cloud/functions-framework';
const { ZOOM_WEBHOOK_SECRET='', SPREADSHEET_ID='' } = process.env;
// ---- 省略: 完全版は元のコード参照 ----
EOF
zip -r ../webhook.zip .
cd ..
```

### 6.2 GUI デプロイ
1. Cloud Console → Cloud Functions → **Create Function**。
2. Node.js 18 / 1st Gen / HTTP。
3. ZIP アップロード、entry point `zoomWebhookHandler`。
4. Runtime variables:
   - `ZOOM_WEBHOOK_SECRET = myWebhookSecret`
   - `SPREADSHEET_ID = <ID>`
5. Deploy → URL をコピー。

### 6.3 Zoom Webhook 設定
Zoom Marketplace アプリ → Event Subscriptions に URL を貼付し Secret Token として `myWebhookSecret` を入力 → Validate。

---

## 7. Apps Script 配備

1. `.clasp.json` を作成:
```jsonc
{ "scriptId": "<scriptId>", "rootDir": "./src" }
```
2. Push:
```bash
clasp login   # 初回のみ
clasp push --force
```

---

## 8. トリガー設定

```bash
clasp run-function setupZoomTriggers            # 録音DLバッチ
clasp run-function setupTranscriptionTriggers   # 文字起こしバッチ
```

---

## 9. 動作確認
1. Zoom でテスト通話 → Cloud Functions で 200。
2. スプレッドシートに行追加。
3. 30 分以内に Drive 保存 → 10 分以内に文字起こし。

---

## 10. アーキテクチャ全体像

```mermaid
flowchart TB
    Z(Zoom Phone) -->|recording_completed| CF[Cloud Functions<br>(zoom-webhook-handler)]
    CF --> Sheets[(Google Sheets<br>Recordings)]
    subgraph GAS Project
      Trigger30((30 min Trigger)) --> ZP[ZoomphoneProcessor]
      ZP -->|download & save| GD[(Google Drive<br>録音フォルダ)]
      Trigger10((10 min Trigger)) --> TP[Transcription Pipeline]
      GD --> TP
      TP --> Sheets
      Retention((Weekly Clean)) --> GD
    end
```

| コンポーネント | 主な役割 | 実装/ファイル |
|---------------|----------|-------------|
| Cloud Functions | Webhook 受信／署名検証／Sheets 書込 | `index.js` (Node 18) |
| Google Sheets   | メタ情報一覧 & ログ | `Recordings` シート |
| GAS: ZoomphoneProcessor | 録音リスト取得 & Drive 保存 | `src/zoom/ZoomphoneProcessor.js` |
| GAS: ZoomphoneService   | Zoom API 呼び出しヘルパ | `src/zoom/ZoomphoneService.js` |
| GAS: Main & triggers    | 文字起こし・要約・メール | `src/main/*.js` |

---

## 11. フロー別の役割分担

| タイミング | フロー | 実行場所 | 詳細 |
|------------|--------|----------|------|
| 録音完了時 即時 | Webhook → メタ記録 | **Cloud Functions** | HMAC 検証 / Sheets へ 1 行追記 (録音 ID, DL URL, 時刻… )|
| 30 分おき | 録音ファイル DL | **GAS `checkAndFetchZoomRecordings`** | Zoom API で 直近 2h をリスト → 未取得を Drive 保存 |
| 10 分おき | 文字起こし & 要約 | **GAS `processBatchOnSchedule`** | AssemblyAI → OpenAI (任意) → Sheets に貼付 |
| 07:15 | 夜間バッチ | GAS | 前日 22:00〜当日 07:00 取得 |
| 月曜 09:10 | 週末バッチ | GAS | 金曜 21:00〜月曜 09:00 取得 |
| 日曜 03:00 | リテンション | GAS | Drive 内 90 日超ファイル削除 |

---

## 12. 前提条件 & 準備チェックリスト

*Zoom 側・Google Cloud 側の設定手順は従来ドキュメントの該当章をそのまま参照可能です。重複説明は割愛し、差分のみを列挙します。*

| 区分 | 必須 | 内容 |
|------|------|------|
| Google Cloud | ○ | Cloud Functions API / Cloud Build API / Sheets API 有効化 |
| Zoom | ○ | `phone:read:recordings` スコープ 付き S2S OAuth アプリ |
| サービスアカウント | ○ | Sheets 書込権限を付与 (`zoom-sheets-integration@…`) |
| Cloud Functions 環境変数 | ○ | `ZOOM_WEBHOOK_SECRET`, `SPREADSHEET_ID` |

---

## 13. 料金モデルと運用コスト

| サービス | 無料枠 | 想定消費量/日 | 超過時単価 (USD) |
|----------|--------|--------------|------------------|
| Cloud Functions (1st Gen) | 2M calls, 400k GB-sec | 1k calls, <5k GB-sec | $0.40/1M calls |
| Cloud Build | 120 build-min/day | デプロイ時のみ | $0.0034/min |
| Sheets API | 無料 | 数十 write | – |
| Apps Script | 90 min exec/day | ≈60 min | – |
| Zoom API | 100 req/日/Account | ≈50 | 超過：1 min cool-down |

> 💰 **結論**: 1 日 1,000 Webhook & 100 録音でも各無料枠以内でほぼゼロコスト。

---

## 14. 動作確認 & テストシナリオ

1. Zoom Phone でテスト通話 → 録音完了を待つ
2. Cloud Functions Logs に Webhook 受信ログが出る & 200 OK を返却
3. Sheets `Recordings` に新行追加を確認 (録音ID / URL …)
4. 30 分以内に Drive フォルダへ WAV/MP3 ファイルが出来る
5. その後 10 分以内に Sheets へ文字起こし & 要約列が更新される
6. 18:10 に日次レポートメールが届く

---

## 15. 運用・監視のベストプラクティス

| 観点 | 推奨策 |
|------|--------|
| エラー検知 | Apps Script の例外を Gmail / Google Chat に送信 |
| リトライ | Cloud Functions は 200 を返し Zoom のリトライを防止 |
| レート制限 | `Utilities.sleep(200)` & pageSize 制御で緩和 |
| Drive 容量 | `RETENTION_DAYS` を 90 日→60 日へ短縮するなど柔軟運用 |
| コスト監視 | Billing Alert を $10/月 で設定 |

---

## 16. 更新履歴

| 日付 | 変更者 | 変更内容 |
|------|-------|----------|
| 2024-06-XX | Dev Team | Cloud Functions + GAS ハイブリッド版 初版 |
