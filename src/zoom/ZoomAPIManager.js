/**
 * Zoom API マネージャーモジュール
 * Zoom API とのインタラクションを管理
 */
var ZoomAPIManager = (function () {
  // プライベート変数
  var accessToken = null;
  var tokenExpiry = null;

  /**
   * アクセストークンを取得する
   * @return {string} - 有効なアクセストークン
   */
  function getAccessToken() {
    try {
      // 必要なZoom認証情報を取得
      var settings = ConfigManager.getConfig();

      // 必要な認証情報が設定されているか確認
      if (!settings.ZOOM_CLIENT_ID || !settings.ZOOM_CLIENT_SECRET || !settings.ZOOM_ACCOUNT_ID) {
        throw new Error('Zoom API の認証情報が設定されていません');
      }

      // 既存のトークンがあり、有効期限内であればそれを使用
      var now = new Date();
      if (accessToken && tokenExpiry && now < tokenExpiry) {
        return accessToken;
      }

      // 新しいアクセストークンを取得
      // 認証URL
      var tokenUrl = 'https://zoom.us/oauth/token';

      // Basic 認証用の認証情報を Base64 エンコード
      var credentials = Utilities.base64Encode(settings.ZOOM_CLIENT_ID + ':' + settings.ZOOM_CLIENT_SECRET);

      // リクエストオプション
      var options = {
        method: 'post',
        headers: {
          'Authorization': 'Basic ' + credentials,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: 'grant_type=account_credentials&account_id=' + settings.ZOOM_ACCOUNT_ID,
        muteHttpExceptions: true
      };

      // API リクエスト実行
      var response = UrlFetchApp.fetch(tokenUrl, options);
      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('アクセストークンの取得に失敗しました。レスポンスコード: ' + responseCode);
      }

      // レスポンスからトークン情報を取得
      var responseJson = JSON.parse(response.getContentText());
      accessToken = responseJson.access_token;

      // トークンの有効期限を設定（1時間 - 5分の余裕を持たせる）
      var expiresIn = (responseJson.expires_in || 3600) - 300;
      tokenExpiry = new Date(now.getTime() + expiresIn * 1000);

      return accessToken;
    } catch (error) {
      throw error;
    }
  }

  /**
   * Zoom API にリクエストを送信する
   * @param {string} endpoint - API エンドポイント
   * @param {string} method - HTTPメソッド（get, postなど）
   * @param {Object} payload - リクエストボディ（オプション）
   * @return {Object} - APIレスポンス
   */
  function sendRequest(endpoint, method, payload) {
    try {
      // アクセストークンを取得
      var token = getAccessToken();

      // API ベースURL
      var baseUrl = 'https://api.zoom.us/v2';
      var url = baseUrl + endpoint;

      // リクエストオプション
      var options = {
        method: method || 'get',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };

      // ペイロードが指定されている場合は追加
      if (payload) {
        options.payload = JSON.stringify(payload);
      }

      // API リクエスト実行
      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();

      // エラーレスポンスの処理
      if (responseCode < 200 || responseCode >= 300) {
        var errorContent = response.getContentText();
        throw new Error('API リクエストに失敗しました。エンドポイント: ' + endpoint +
          ', レスポンスコード: ' + responseCode +
          ', 詳細: ' + errorContent);
      }

      // レスポンスを JSON としてパース
      var responseJson = JSON.parse(response.getContentText());
      return responseJson;
    } catch (error) {
      throw error;
    }
  }


  // 公開メソッド
  return {
    getAccessToken: getAccessToken,
    sendRequest: sendRequest
  };
})();

/**
 * Zoom APIのアクセストークンを更新する関数
 * 毎日5:00にトリガーによって実行され、新しいトークンを取得する
 * @return {string} 実行結果メッセージ
 */
function refreshZoomAPIToken() {
  try {
    Logger.log('Zoom APIトークンの更新を開始します...');

    // ZoomAPIManagerのgetAccessToken()を呼び出してトークンを強制的に再取得
    var token = ZoomAPIManager.getAccessToken();

    if (token) {
      Logger.log('Zoom APIトークンの更新に成功しました');
      return 'Zoom APIトークンの更新に成功しました';
    } else {
      throw new Error('トークンの取得に失敗しました');
    }
  } catch (error) {
    var errorMsg = error.toString();
    Logger.log('Zoom APIトークン更新中にエラー: ' + errorMsg);

    // エラー通知（もし設定されていれば）
    try {
      var settings = EnvironmentConfig.getConfig();
      if (settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
        var subject = "Zoom APIトークン更新エラー";
        var body = "Zoom APIトークンの更新中にエラーが発生しました。\n\n" +
          "日時: " + new Date().toLocaleString() + "\n" +
          "エラー: " + errorMsg + "\n\n" +
          "システム管理者に連絡してください。";

        for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
          GmailApp.sendEmail(settings.ADMIN_EMAILS[i], subject, body);
        }
      }
    } catch (notifyError) {
      Logger.log('エラー通知の送信中にエラー: ' + notifyError.toString());
    }

    return 'Zoom APIトークン更新中にエラー: ' + errorMsg;
  }
}