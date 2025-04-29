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
      // 必要なZoom認証情報を確実に取得するため、スクリプトプロパティも統合するloadSettingsを使用
      var settings = loadSettings();

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

  /**
   * 設定を読み込む
   * @return {Object} - 設定オブジェクト
   */
  function loadSettings() {
    try {
      // 基本設定を取得
      var settings = getSystemSettings();

      // スクリプトプロパティからの設定読み込み
      var scriptProperties = PropertiesService.getScriptProperties();

      // Zoom API 認証情報を設定
      settings.ZOOM_CLIENT_ID = settings.ZOOM_CLIENT_ID || scriptProperties.getProperty('ZOOM_CLIENT_ID');
      settings.ZOOM_CLIENT_SECRET = settings.ZOOM_CLIENT_SECRET || scriptProperties.getProperty('ZOOM_CLIENT_SECRET');
      settings.ZOOM_ACCOUNT_ID = settings.ZOOM_ACCOUNT_ID || scriptProperties.getProperty('ZOOM_ACCOUNT_ID');

      return settings;
    } catch (error) {
      return {
        ZOOM_CLIENT_ID: '',
        ZOOM_CLIENT_SECRET: '',
        ZOOM_ACCOUNT_ID: ''
      };
    }
  }

  // 公開メソッド
  return {
    getAccessToken: getAccessToken,
    sendRequest: sendRequest,
    loadSettings: loadSettings
  };
})();