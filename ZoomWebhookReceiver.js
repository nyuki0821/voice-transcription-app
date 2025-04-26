/**
 * Zoom Webhook Receiver
 * Zoom の Webhook (phone.recording_completed) を受信し、録音ファイルを Drive に保存します。
 *
 * - URL 検証フェーズでは plainToken に対し HMAC-SHA256 を計算し encryptedToken を返却
 * - 本番イベントでは ZoomphoneProcessor.processWebhookRecording() を呼び出し
 *
 * スクリプトプロパティ
 *   WEBHOOK_TOKEN     : Zoom Marketplace の Event Subscription で設定した Secret Token
 */
function doPost(e) {
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_TOKEN');
    if (!secret) {
      throw new Error('WEBHOOK_TOKEN が設定されていません');
    }

    var data = {};
    try {
      data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (parseErr) {
      // JSON 解析に失敗した場合は 400
      return ContentService.createTextOutput('Invalid JSON').setMimeType(ContentService.MimeType.TEXT);
    }

    // --- URL 検証フェーズ -------------------------------------
    if (data && data.payload && data.payload.plainToken) {
      var plainToken = data.payload.plainToken;
      var hash = Utilities.computeHmacSha256Signature(plainToken, secret);
      var encrypted = Utilities.base64Encode(hash);
      var resp = {
        plainToken: plainToken,
        encryptedToken: encrypted
      };
      return ContentService.createTextOutput(JSON.stringify(resp))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- 本番イベント -------------------------------------------
    if (data.event === 'phone.recording_completed') {
      var p = data.payload || {};
      var recId = p.recording_id;
      var downloadUrl = p.download_url || p.download_url_with_token;
      var startTime = p.start_time;

      var result = ZoomphoneProcessor.processWebhookRecording(recId, downloadUrl, startTime);
      var status = (result && result.success) ? 'ok' : 'ng';
      return ContentService.createTextOutput(status).setMimeType(ContentService.MimeType.TEXT);
    }

    // それ以外のイベントは 200 OK を返す
    return ContentService.createTextOutput('ignored')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    // 例外時も 200 を返して Zoom 側のリトライループを防止
    Logger.log('[ZoomWebhookReceiver] Error: ' + err);
    return ContentService.createTextOutput('error').setMimeType(ContentService.MimeType.TEXT);
  }
} 