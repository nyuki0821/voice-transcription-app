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

/**
 * ISO形式の日時文字列から日付部分だけを抽出（YYYY-MM-DD形式）
 * @param {string} isoDateTime - ISO形式の日時文字列
 * @return {string} - YYYY-MM-DD形式の日付文字列
 */
function extractDatePart(isoDateTime) {
  return isoDateTime.split('T')[0]; // ISO形式（YYYY-MM-DDThh:mm:ss）から日付部分だけ抽出
}

/**
 * ISO形式の日時文字列から時間部分だけを抽出（HH:MM:SS形式）
 * @param {string} isoDateTime - ISO形式の日時文字列
 * @return {string} - HH:MM:SS形式の時間文字列
 */
function extractTimePart(isoDateTime) {
  const timePart = isoDateTime.split('T')[1];
  return timePart ? timePart.split('.')[0] : '00:00:00'; // 秒までの部分を抽出
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

      // 通話日時を処理（JSTに変換）
      const callDateTime = formatToJST(recording.date_time);
      // 日付と時間を分離（「YYYY-MM-DD HH:MM:SS」→「YYYY-MM-DD」と「HH:MM:SS」）
      const callDate = callDateTime.split(' ')[0];
      const callTime = callDateTime.split(' ')[1];

      // 電話番号を抽出してフォーマット
      const direction = recording.direction || 'unknown';
      let salesPhoneNumber = '';
      let customerPhoneNumber = '';

      if (direction === 'outbound') {
        // 発信コール：発信者（営業）と着信者（顧客）
        salesPhoneNumber = recording.caller_number || '';
        customerPhoneNumber = recording.callee_number || '';
      } else if (direction === 'inbound') {
        // 着信コール：発信者（顧客）と着信者（営業）
        customerPhoneNumber = recording.caller_number || '';
        salesPhoneNumber = recording.callee_number || '';
      } else {
        // direction が unknown の場合
        // 少なくとも1つの電話番号を取得できていれば処理
        if (recording.caller_number && recording.callee_number) {
          // 両方ある場合は caller を営業、callee を顧客と仮定
          salesPhoneNumber = recording.caller_number || '';
          customerPhoneNumber = recording.callee_number || '';
        } else if (recording.caller_number) {
          // caller_number のみの場合は営業の電話番号と仮定
          salesPhoneNumber = recording.caller_number || '';
          // 顧客番号は不明のため空文字
        } else if (recording.callee_number) {
          // callee_number のみの場合は顧客の電話番号と仮定
          customerPhoneNumber = recording.callee_number || '';
        }

        // ログに記録
        console.log('[CF] 通話方向不明: caller=' + recording.caller_number +
          ', callee=' + recording.callee_number +
          ' → sales=' + salesPhoneNumber +
          ', customer=' + customerPhoneNumber);
      }

      // 電話番号の国際形式を国内形式に変換
      if (customerPhoneNumber.startsWith('+81')) {
        // 日本の電話番号 (+81...) → 0...
        customerPhoneNumber = '0' + customerPhoneNumber.substring(3);
      }
      if (salesPhoneNumber.startsWith('+81')) {
        // 日本の電話番号 (+81...) → 0...
        salesPhoneNumber = '0' + salesPhoneNumber.substring(3);
      }

      const rows = [[
        recording.id || '',                 // record_id
        formatToJST(new Date()),            // timestamp_recording
        recording.download_url || '',       // download_url
        callDate,                           // call_date
        callTime,                           // call_time
        recording.duration || '',           // duration
        salesPhoneNumber,                   // sales_phone_number
        customerPhoneNumber,                // customer_phone_number
        '',                                 // timestamp_fetch
        'PENDING',                          // status_fetch
        '',                                 // timestamp_transcription
        '',                                 // status_transcription
        '',                                 // process_start
        ''                                  // process_end
      ]];
      console.log('[CF] スプレッドシートに追加するデータ:', rows);

      const api = await getSheets();
      await api.spreadsheets.values.append({
        spreadsheetId: RECORDINGS_SHEET_ID,
        range: 'Recordings!A:N', // 14列に変更
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