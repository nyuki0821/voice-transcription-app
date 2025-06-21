/**
 * 設定管理サービス
 * プロジェクト全体の設定を一元管理
 */
var ConfigManager = (function () {

  // 設定キャッシュ
  var _configCache = null;
  var _cacheTimestamp = null;
  var CACHE_DURATION = 5 * 60 * 1000; // 5分間キャッシュ

  /**
   * 設定を取得（キャッシュ機能付き）
   * @param {boolean} [forceRefresh] - 強制的にキャッシュを更新する場合true
   * @return {Object} - 設定オブジェクト
   */
  function getConfig(forceRefresh) {
    var now = new Date().getTime();

    // キャッシュが有効かチェック
    if (!forceRefresh && _configCache && _cacheTimestamp &&
      (now - _cacheTimestamp) < CACHE_DURATION) {
      return _configCache;
    }

    try {
      // 環境設定を取得
      var envConfig = EnvironmentConfig.getConfig();

      // 設定オブジェクトを構築
      _configCache = {
        // API設定
        OPENAI_API_KEY: envConfig.OPENAI_API_KEY || '',

        // フォルダ設定
        SOURCE_FOLDER_ID: envConfig.SOURCE_FOLDER_ID || '',
        PROCESSING_FOLDER_ID: envConfig.PROCESSING_FOLDER_ID || '',
        COMPLETED_FOLDER_ID: envConfig.COMPLETED_FOLDER_ID || '',
        ERROR_FOLDER_ID: envConfig.ERROR_FOLDER_ID || '',

        // スプレッドシート設定
        RECORDINGS_SHEET_ID: envConfig.RECORDINGS_SHEET_ID || '',
        PROCESSED_SHEET_ID: envConfig.PROCESSED_SHEET_ID || '',

        // Zoom設定
        ZOOM_CLIENT_ID: envConfig.ZOOM_CLIENT_ID || '',
        ZOOM_CLIENT_SECRET: envConfig.ZOOM_CLIENT_SECRET || '',
        ZOOM_ACCOUNT_ID: envConfig.ZOOM_ACCOUNT_ID || '',
        ZOOM_WEBHOOK_SECRET: envConfig.ZOOM_WEBHOOK_SECRET || '',

        // 処理設定
        MAX_BATCH_SIZE: envConfig.MAX_BATCH_SIZE || Constants.PROCESSING.MAX_BATCH_SIZE_DEFAULT,
        ENHANCE_WITH_OPENAI: envConfig.ENHANCE_WITH_OPENAI !== false,

        // 通知設定
        ADMIN_EMAILS: envConfig.ADMIN_EMAILS || []
      };

      _cacheTimestamp = now;
      return _configCache;

    } catch (error) {
      Logger.log('設定の読み込み中にエラー: ' + error.toString());
      return getDefaultConfig();
    }
  }

  /**
   * デフォルト設定を取得
   * @return {Object} - デフォルト設定オブジェクト
   */
  function getDefaultConfig() {
    return {
      OPENAI_API_KEY: '',
      SOURCE_FOLDER_ID: '',
      PROCESSING_FOLDER_ID: '',
      COMPLETED_FOLDER_ID: '',
      ERROR_FOLDER_ID: '',
      RECORDINGS_SHEET_ID: '',
      PROCESSED_SHEET_ID: '',
      ZOOM_CLIENT_ID: '',
      ZOOM_CLIENT_SECRET: '',
      ZOOM_ACCOUNT_ID: '',
      ZOOM_WEBHOOK_SECRET: '',
      MAX_BATCH_SIZE: Constants.PROCESSING.MAX_BATCH_SIZE_DEFAULT,
      ENHANCE_WITH_OPENAI: true,
      ADMIN_EMAILS: []
    };
  }

  /**
   * 特定の設定値を取得
   * @param {string} key - 設定キー
   * @param {*} [defaultValue] - デフォルト値
   * @return {*} - 設定値
   */
  function get(key, defaultValue) {
    var config = getConfig();
    return config.hasOwnProperty(key) ? config[key] : defaultValue;
  }

  /**
   * Recordingsシートのスプレッドシートを取得
   * @return {Spreadsheet} - スプレッドシートオブジェクト
   */
  function getRecordingsSpreadsheet() {
    var sheetId = get('RECORDINGS_SHEET_ID');
    if (!sheetId) {
      throw new Error('RECORDINGS_SHEET_IDが設定されていません');
    }
    return SpreadsheetApp.openById(sheetId);
  }

  /**
   * Recordingsシートを取得
   * @return {Sheet} - シートオブジェクト
   */
  function getRecordingsSheet() {
    var spreadsheet = getRecordingsSpreadsheet();
    var sheet = spreadsheet.getSheetByName(Constants.SHEET_NAMES.RECORDINGS);
    if (!sheet) {
      throw new Error(Constants.formatMessage(Constants.ERROR_MESSAGES.SHEET_NOT_FOUND, {
        sheetName: Constants.SHEET_NAMES.RECORDINGS
      }));
    }
    return sheet;
  }

  /**
   * Processedシートのスプレッドシートを取得
   * @return {Spreadsheet} - スプレッドシートオブジェクト
   */
  function getProcessedSpreadsheet() {
    var sheetId = get('PROCESSED_SHEET_ID');
    if (!sheetId) {
      throw new Error('PROCESSED_SHEET_IDが設定されていません');
    }
    return SpreadsheetApp.openById(sheetId);
  }

  /**
   * Call Recordsシートを取得
   * @return {Sheet} - シートオブジェクト
   */
  function getCallRecordsSheet() {
    var spreadsheet = getProcessedSpreadsheet();
    var sheet = spreadsheet.getSheetByName(Constants.SHEET_NAMES.CALL_RECORDS);
    if (!sheet) {
      throw new Error(Constants.formatMessage(Constants.ERROR_MESSAGES.SHEET_NOT_FOUND, {
        sheetName: Constants.SHEET_NAMES.CALL_RECORDS
      }));
    }
    return sheet;
  }

  /**
   * フォルダオブジェクトを取得
   * @param {string} folderType - フォルダタイプ（SOURCE, PROCESSING, COMPLETED, ERROR）
   * @return {Folder} - フォルダオブジェクト
   */
  function getFolder(folderType) {
    var folderId = get(folderType + '_FOLDER_ID');
    if (!folderId) {
      throw new Error(folderType + '_FOLDER_IDが設定されていません');
    }
    return DriveApp.getFolderById(folderId);
  }

  /**
   * 設定の妥当性をチェック
   * @return {Object} - チェック結果 {valid: boolean, errors: string[]}
   */
  function validateConfig() {
    var config = getConfig();
    var errors = [];

    // 必須設定のチェック
    var requiredKeys = [
      'OPENAI_API_KEY',
      'SOURCE_FOLDER_ID',
      'PROCESSING_FOLDER_ID',
      'COMPLETED_FOLDER_ID',
      'ERROR_FOLDER_ID',
      'RECORDINGS_SHEET_ID',
      'PROCESSED_SHEET_ID'
    ];

    for (var i = 0; i < requiredKeys.length; i++) {
      var key = requiredKeys[i];
      if (!config[key]) {
        errors.push(key + 'が設定されていません');
      }
    }

    // 管理者メールの設定チェック
    if (!config.ADMIN_EMAILS || config.ADMIN_EMAILS.length === 0) {
      errors.push('ADMIN_EMAILSが設定されていません');
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * キャッシュをクリア
   */
  function clearCache() {
    _configCache = null;
    _cacheTimestamp = null;
  }

  // 公開メソッド
  return {
    getConfig: getConfig,
    getDefaultConfig: getDefaultConfig,
    get: get,
    getRecordingsSpreadsheet: getRecordingsSpreadsheet,
    getRecordingsSheet: getRecordingsSheet,
    getProcessedSpreadsheet: getProcessedSpreadsheet,
    getCallRecordsSheet: getCallRecordsSheet,
    getFolder: getFolder,
    validateConfig: validateConfig,
    clearCache: clearCache
  };
})(); 