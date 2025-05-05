/**
 * Zoom Webhook 署名検証の実装例
 * 
 * 本コードは Zoom から送信されるwebhookリクエストの署名を検証するための
 * 実装例を提供します。
 */

// 必要なモジュールのインポート
const crypto = require('crypto');

/**
 * Zoom Webhookの署名を検証する関数
 * 
 * @param {Object} req - HTTPリクエストオブジェクト
 * @param {Object} req.headers - リクエストヘッダー
 * @param {Object|string} req.body - リクエストボディ（JSONオブジェクトまたは文字列）
 * @param {string} secretToken - Zoom webhook secret token
 * @return {boolean} - 署名が有効な場合はtrue、それ以外はfalse
 */
function verifyZoomWebhookSignature(req, secretToken) {
  try {
    // 署名とタイムスタンプを取得
    // Zoomは複数の可能性のあるヘッダー名を使用する場合があるため、両方をチェック
    const signature = req.headers['x-zm-signature'] || req.headers['x-zoom-signature'];
    let timestamp = req.headers['x-zm-request-timestamp'] || req.headers['x-zoom-request-timestamp'];

    // タイムスタンプがヘッダーにない場合、event_tsフィールドから取得を試みる（ミリ秒→秒に変換）
    if (!timestamp && req.body && req.body.event_ts) {
      timestamp = Math.floor(req.body.event_ts / 1000).toString();
    }

    if (!signature || !timestamp) {
      console.warn('署名またはタイムスタンプが見つかりません:', { signature, timestamp });
      return false;
    }

    // リクエストボディを文字列化
    let rawBody = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);

    // 重要: 正しいメッセージ形式は "v0:タイムスタンプ:JSONボディ"
    const message = `v0:${timestamp}:${rawBody}`;
    console.log('検証用メッセージ:', message.substring(0, 100) + (message.length > 100 ? '...' : ''));

    // HMAC-SHA256で署名を計算
    const hmac = crypto.createHmac('sha256', secretToken);
    hmac.update(message);
    const expectedSignature = `v0=${hmac.digest('hex')}`;

    // 署名を比較（Zoomがv0=プレフィックスを付けない場合に対応）
    const receivedSignature = signature.startsWith('v0=') ? signature : `v0=${signature}`;

    console.log('受信した署名:', receivedSignature);
    console.log('計算した署名:', expectedSignature);

    return receivedSignature === expectedSignature;
  } catch (error) {
    console.error('署名検証中にエラーが発生しました:', error);
    return false;
  }
}

/**
 * Zoom エンドポイント検証トークンを生成する関数
 * 
 * @param {string} plainToken - Zoomから送信されたプレーンテキストトークン
 * @param {string} secretToken - Zoom webhook secret token
 * @return {string} - 暗号化されたトークン
 */
function generateZoomValidationToken(plainToken, secretToken) {
  const hmac = crypto.createHmac('sha256', secretToken);
  hmac.update(plainToken);
  return hmac.digest('hex');
}

/**
 * Express.jsでの使用例
 * 
 * 以下のコードはExpress.jsアプリでのZoom Webhook検証の実装例です。
 * このサンプルコードは実際に動作するコードではなく、実装方法を示すための例です。
 */
/*
const express = require('express');
const app = express();
const ZOOM_WEBHOOK_SECRET = process.env.ZOOM_WEBHOOK_SECRET || 'your_secret_token';

// JSONボディの解析
app.use(express.json({
  verify: (req, res, buf) => {
    // 生のリクエストボディをbuf.toString()で取得できるようにする
    req.rawBody = buf.toString();
  }
}));

// Zoom Webhookエンドポイント
app.post('/webhook', (req, res) => {
  // 署名検証
  if (!verifyZoomWebhookSignature(req, ZOOM_WEBHOOK_SECRET)) {
    console.warn('Webhook署名検証失敗');
    return res.status(401).send('Unauthorized: Invalid signature');
  }

  // エンドポイントURL検証リクエストの処理
  if (req.body.event === 'endpoint.url_validation') {
    const encryptedToken = generateZoomValidationToken(
      req.body.payload.plainToken,
      ZOOM_WEBHOOK_SECRET
    );
    return res.status(200).json({
      plainToken: req.body.payload.plainToken,
      encryptedToken: encryptedToken
    });
  }

  // 他のWebhookイベントの処理
  console.log('Webhook受信:', req.body.event);
  
  // イベントタイプに基づいて処理
  switch (req.body.event) {
    case 'phone.recording_completed':
      // 電話録音完了イベントの処理
      console.log('録音完了:', req.body.payload);
      break;
      
    case 'meeting.started':
      // ミーティング開始イベントの処理
      console.log('ミーティング開始:', req.body.payload);
      break;
      
    // 他のイベントタイプも同様に処理
    
    default:
      console.log('未処理のイベント:', req.body.event);
  }
  
  // 正常応答
  res.status(200).send('Event received');
});

// サーバー起動
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Zoom Webhook serverが起動しました: http://localhost:${PORT}`);
});
*/

// モジュールとしてエクスポート
module.exports = {
  verifyZoomWebhookSignature,
  generateZoomValidationToken
}; 