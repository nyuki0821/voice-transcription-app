/**
 * 環境変数設定管理
 * すべての環境変数をスプレッドシートの「環境設定」シートから一元管理します
 */

var EnvironmentConfig = (function () {
  // キャッシュされた設定
  var cachedConfig = null;
  var lastFetchTime = null;
  // キャッシュの有効期限（30分）
  var CACHE_VALIDITY_MS = 30 * 60 * 1000;

  // settingsシート名
  var CONFIG_SHEET_NAME = 'settings';

  /**
   * 環境変数の設定を取得する
   * @param {boolean} forceRefresh - 強制的に再取得する場合はtrue
   * @return {Object} - 設定オブジェクト
   */
  function getConfig(forceRefresh) {
    var now = new Date().getTime();

    // キャッシュが有効な場合はキャッシュから返す
    if (!forceRefresh && cachedConfig && lastFetchTime && (now - lastFetchTime < CACHE_VALIDITY_MS)) {
      return cachedConfig;
    }

    try {
      // スプレッドシートIDを取得
      var configSpreadsheetId = getConfigSpreadsheetId();
      if (!configSpreadsheetId) {
        throw new Error('設定スプレッドシートIDが設定されていません');
      }

      // スプレッドシートを開く
      var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
      if (!spreadsheet) {
        throw new Error('設定スプレッドシートを開けませんでした');
      }

      // settingsシートを取得
      var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
      if (!configSheet) {
        throw new Error('settingsシートが見つかりません');
      }

      // 設定データを読み込む
      var configData = configSheet.getDataRange().getValues();
      if (!configData || configData.length < 2) {
        throw new Error('設定データが見つかりません');
      }

      // 設定マップを作成
      var configMap = {};
      for (var i = 1; i < configData.length; i++) {
        if (configData[i].length >= 2) {
          var key = String(configData[i][0] || '').trim();
          var value = configData[i][1];

          if (key) {
            configMap[key] = value;
          }
        }
      }

      // メールアドレスリストの処理（カンマ区切り文字列を配列に変換）
      if (configMap['ADMIN_EMAIL']) {
        configMap['ADMIN_EMAILS'] = String(configMap['ADMIN_EMAIL'])
          .split(',')
          .map(function (email) { return email.trim(); })
          .filter(function (email) { return email !== ''; });
      } else {
        configMap['ADMIN_EMAILS'] = [];
      }

      // 数値型の設定項目を変換
      if (configMap['MAX_BATCH_SIZE']) {
        configMap['MAX_BATCH_SIZE'] = parseInt(configMap['MAX_BATCH_SIZE'], 10) || 10;
      } else {
        configMap['MAX_BATCH_SIZE'] = 10;
      }

      // ブール値の設定項目を変換
      configMap['ENHANCE_WITH_OPENAI'] = configMap['ENHANCE_WITH_OPENAI'] !== false;

      // キャッシュを更新
      cachedConfig = configMap;
      lastFetchTime = now;

      return configMap;
    } catch (error) {
      Logger.log('環境設定の読み込み中にエラー: ' + error);
      // エラー時は空の設定を返す
      return getDefaultConfig();
    }
  }

  /**
   * 設定スプレッドシートIDを取得
   * @return {string} - 設定スプレッドシートID
   */
  function getConfigSpreadsheetId() {
    // 環境設定スプレッドシートIDをスクリプトプロパティから取得
    var configId = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
    return configId;
  }

  /**
   * デフォルト設定を返す
   * @return {Object} - デフォルト設定オブジェクト
   */
  function getDefaultConfig() {
    return {
      // スプレッドシートID (新設)
      RECORDINGS_SHEET_ID: '16wE2ebRsRSXDk0BHXawfziKsp3JDiKueTKOuWvGgJxo', // クラウドファンクションで使用しているスプレッドシートID
      PROCESSED_SHEET_ID: '',

      // 既存互換用
      SPREADSHEET_ID: '',
      ASSEMBLYAI_API_KEY: '',
      OPENAI_API_KEY: '',
      SOURCE_FOLDER_ID: '',
      PROCESSING_FOLDER_ID: '',
      COMPLETED_FOLDER_ID: '',
      ERROR_FOLDER_ID: '',
      ADMIN_EMAILS: [],
      MAX_BATCH_SIZE: 10,
      ENHANCE_WITH_OPENAI: true,
      ZOOM_CLIENT_ID: '',
      ZOOM_CLIENT_SECRET: '',
      ZOOM_ACCOUNT_ID: '',
      ZOOM_WEBHOOK_SECRET: '',
      RETENTION_DAYS: 90,
      PROCESSING_ENABLED: 'true'
    };
  }

  /**
   * 特定の設定値を取得
   * @param {string} key - 設定キー
   * @param {*} defaultValue - デフォルト値（設定が見つからない場合）
   * @return {*} - 設定値
   */
  function get(key, defaultValue) {
    var config = getConfig();
    return (key in config) ? config[key] : defaultValue;
  }

  /**
   * キャッシュを強制的にクリア
   */
  function clearCache() {
    cachedConfig = null;
    lastFetchTime = null;
  }

  // 公開API
  return {
    getConfig: getConfig,
    get: get,
    clearCache: clearCache
  };
})(); 