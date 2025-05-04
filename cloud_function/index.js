import { google } from 'googleapis';
import crypto from 'crypto';
import functions from '@google-cloud/functions-framework';

// --- 環境変数 ---------------------------------------------------------
const {
  ZOOM_WEBHOOK_SECRET = '',   // Zoom Dev Portal で発行した Secret Token
  RECORDINGS_SHEET_ID = '',  // Recordings シートのスプレッドシートID（名称変更）
} = process.env;

if (!ZOOM_WEBHOOK_SECRET || !RECORDINGS_SHEET_ID) {
  console.warn('[CF] 環境変数が未設定です');
}

// --- Google Sheets クライアント --------------------------------------
const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
let sheets; // lazy init
async function getSheets() {
  if (!sheets) {
    const authClient = await auth.getClient();
    sheets = google.sheets({ version: 'v4', auth: authClient });
  }
  return sheets;
}

/**
 * 日付を日本時間(JST)でフォーマットする関数
 * @param {string|Date} date - ISO形式の日付文字列またはDateオブジェクト
 * @return {string} - 「YYYY-MM-DD HH:MM:SS」形式の日本時間
 */
function formatToJST(date) {
  const d = typeof date === 'string' ? new Date(date) : date;

  // UTCの日付に9時間を加算して日本時間に変換
  const jstDate = new Date(d.getTime() + 9 * 60 * 60 * 1000);

  // YYYY-MM-DD HH:MM:SS 形式にフォーマット
  return jstDate.toISOString().replace('T', ' ').substring(0, 19);
}

// --- 関数本体 --------------------------------------------------------
functions.http('zoomWebhookHandler', async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  console.log('[CF] リクエスト受信:', {
    headers: JSON.stringify(req.headers),
    body: JSON.stringify(req.body).substring(0, 200) + '...' // 長いボディは省略
  });

  // 署名検証（plainToken検証以外）
  if (!(req.body?.payload?.plainToken)) {
    // 複数のヘッダー名に対応
    const signature = req.get('x-zm-signature') || req.get('x-zoom-signature');
    let ts = req.get('x-zm-request-timestamp') || req.get('x-zoom-request-timestamp');

    // event_tsからタイムスタンプを取得する場合（ミリ秒→秒に変換）
    if (!ts && req.body?.event_ts) {
      ts = Math.floor(req.body.event_ts / 1000).toString();
    }

    if (!signature || !ts) {
      console.warn('[CF] 署名またはタイムスタンプが見つかりません', {
        signature,
        ts,
        headers: JSON.stringify(req.headers)
      });
      return res.status(401).send('Unauthorized: Missing signature or timestamp');
    }

    // rawBodyがなければJSON.stringify(req.body)で代用
    let rawBody = req.rawBody;
    if (!rawBody) {
      rawBody = JSON.stringify(req.body);
    }
    if (Buffer.isBuffer(rawBody)) {
      rawBody = rawBody.toString('utf8');
    }

    // 正しいメッセージ形式: v0:タイムスタンプ:JSONボディ
    const msg = `v0:${ts}:${rawBody}`;
    console.log('[CF] 検証用メッセージ:', msg.substring(0, 100) + '...'); // 長いメッセージは省略

    const hmac = crypto.createHmac('sha256', ZOOM_WEBHOOK_SECRET);
    hmac.update(msg);
    const expected = hmac.digest('hex');

    // 署名比較（Zoomは v0= プレフィックスをつけて送信する場合がある）
    const receivedSignature = signature.startsWith('v0=') ? signature.substring(3) : signature;
    const match = receivedSignature === expected;

    console.log('[CF] 署名検証:', {
      received: signature,
      receivedWithoutPrefix: receivedSignature,
      calculated: expected,
      match
    });

    if (!match) {
      console.warn('[CF] Zoom署名検証失敗', {
        signature,
        receivedWithoutPrefix: receivedSignature,
        expected,
        ts,
        secretLength: ZOOM_WEBHOOK_SECRET.length,
        // 長すぎるデータはログ出力しない
        msgPrefix: msg.substring(0, 50) + '...',
        bodyPrefix: rawBody.substring(0, 50) + '...'
      });

      // デバッグモードから本番モードに変更 - 401 Unauthorizedを返す
      return res.status(401).send('Unauthorized: Invalid signature');
    }
  }

  try {
    const body = req.body || {};
    // 1. URL 検証 (plainToken → encryptedToken)
    if (body?.payload?.plainToken) {
      const plain = body.payload.plainToken;
      const hmac = crypto.createHmac('sha256', ZOOM_WEBHOOK_SECRET);
      hmac.update(plain);
      return res.status(200).json({ plainToken: plain, encryptedToken: hmac.digest('hex') });
    }

    // 2. 録音完了イベント
    if (body?.event === 'phone.recording_completed') {
      const p = body.payload || {};
      console.log('[CF] 録音完了ペイロード全体:', JSON.stringify(p));

      // 修正：recordings配列の最初の要素を取得
      const recording = p.object?.recordings?.[0] || {};

      // 電話番号を抽出してフォーマット
      let phoneNumber = recording.direction === 'outbound' ? recording.callee_number : recording.caller_number || '';

      // 電話番号の国際形式を国内形式に変換
      if (phoneNumber.startsWith('+81')) {
        // 日本の電話番号 (+81...) → 0...
        phoneNumber = '0' + phoneNumber.substring(3);
      } else if (phoneNumber.startsWith('+')) {
        // その他の国際電話番号は+を削除するだけ
        phoneNumber = phoneNumber.substring(1);
      }

      const rows = [[
        formatToJST(new Date()),            // Timestamp (システム時刻) - 日本時間
        recording.id || '',                 // RecordingId
        recording.download_url || '',       // DownloadUrl
        formatToJST(recording.date_time),   // StartTime (Zoom通話開始時刻) - 日本時間
        phoneNumber,                        // PhoneNumber (customer_phone_number形式)
        recording.duration || '',           // Duration
        'PENDING'                           // Status
      ]];
      console.log('[CF] スプレッドシートに追加するデータ:', rows);

      const api = await getSheets();
      await api.spreadsheets.values.append({
        spreadsheetId: RECORDINGS_SHEET_ID,
        range: 'Recordings!A:G',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: rows },
      });
      return res.status(200).send('ok');
    }

    // 3. 未ハンドル
    console.log('[CF] 未処理のイベント:', body.event);
    return res.status(200).send('ignored');
  } catch (err) {
    console.error('[CF] Error:', err);
    // Zoom の再送を避けるため 200 で返す
    return res.status(200).send('error');
  }
});