import { google } from 'googleapis';
import crypto from 'crypto';
import functions from '@google-cloud/functions-framework';

// --- 環境変数 ---------------------------------------------------------
const {
  ZOOM_WEBHOOK_SECRET = '',   // Zoom Dev Portal で発行した Secret Token
  SPREADSHEET_ID = '',       // Recordings シートのスプレッドシートID
} = process.env;

if (!ZOOM_WEBHOOK_SECRET || !SPREADSHEET_ID) {
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

// --- 関数本体 --------------------------------------------------------
functions.http('zoomWebhookHandler', async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
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
      const rows = [[
        new Date().toISOString(),
        p.recording_id || '',
        p.download_url || p.download_url_with_token || '',
        p.start_time || '',
        p.phone_number || '',
        p.duration || '',
        'PENDING'
      ]];
      const api = await getSheets();
      await api.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Recordings!A:G',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: rows },
      });
      return res.status(200).send('ok');
    }

    // 3. 未ハンドル
    return res.status(200).send('ignored');
  } catch (err) {
    console.error('[CF] Error:', err);
    // Zoom の再送を避けるため 200 で返す
    return res.status(200).send('error');
  }
});