/**
 * トリガー管理モジュール
 * すべてのアプリケーショントリガーの設定・管理を一括して行う
 */

// TriggerManager名前空間を定義
var TriggerManager = (function () {
  return {
    setupZoomTriggers: setupZoomTriggers,
    setupTranscriptionTriggers: setupTranscriptionTriggers,
    setupDailyZoomApiTokenRefresh: setupDailyZoomApiTokenRefresh,
    setupRecordingsSheetTrigger: setupRecordingsSheetTrigger,
    setupAllTriggers: setupAllTriggers,
    deleteTriggersWithNameContaining: deleteTriggersWithNameContaining,
    deleteAllTriggers: deleteAllTriggers,
    fetchLastHourRecordings: fetchLastHourRecordings,
    fetchLast2HoursRecordings: fetchLast2HoursRecordings,
    fetchLast6HoursRecordings: fetchLast6HoursRecordings,
    fetchLast24HoursRecordings: fetchLast24HoursRecordings,
    fetchLast48HoursRecordings: fetchLast48HoursRecordings,
    fetchAllPendingRecordings: fetchAllPendingRecordings,
    setupRecoveryTriggers: setupRecoveryTriggers,
    removeRecoveryTriggers: removeRecoveryTriggers,
    setupPartialFailureDetectionTrigger: setupPartialFailureDetectionTrigger
  };
})();

/**
 * ZoomPhone 連携用のトリガー設定関数
 * ZoomPhone APIを使って自動的に録音ファイルを取得・処理するためのトリガーを設定
 */
function setupZoomTriggers() {
  // 既存のZoom関連のトリガーをすべて削除
  deleteTriggersWithNameContaining('fetchZoomRecordings');
  deleteTriggersWithNameContaining('checkAndFetchZoomRecordings');
  deleteTriggersWithNameContaining('purgeOldRecordings');
  deleteTriggersWithNameContaining('processRecordingsFromSheet');

  // 1. Recordingsシート処理の30分ごとのトリガー
  ScriptApp.newTrigger('processRecordingsFromSheet')
    .timeBased()
    .everyMinutes(30)
    .create();

  // 2. 朝6:15の夜間バッチ処理 - 毎日実行
  ScriptApp.newTrigger('fetchZoomRecordingsMorningBatch')
    .timeBased()
    .atHour(6)
    .nearMinute(15)
    .everyDays(1)
    .create();

  // 3. リテンションバッチ: 日曜 03:00 に古いファイルを削除
  ScriptApp.newTrigger('purgeOldRecordings')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .nearMinute(0)
    .create();

  return '以下のZoom録音取得トリガーを設定しました：\n' +
    '1. Recordingsシート処理: 30分ごとに実行\n' +
    '2. 夜間バッチ: 毎朝6:15\n' +
    '3. リテンションバッチ: 日曜 03:00 に古いファイルを削除';
}

/**
 * 文字起こし関連のトリガー設定関数
 */
function setupTranscriptionTriggers() {
  // 既存のトリガーをすべて削除
  deleteTriggersWithNameContaining('startDailyProcess');
  deleteTriggersWithNameContaining('processBatchOnSchedule');
  deleteTriggersWithNameContaining('sendDailySummary');

  // 6時にプロセスを開始するトリガー
  ScriptApp.newTrigger('startDailyProcess')
    .timeBased()
    .atHour(6)
    .nearMinute(0)
    .everyDays(1)
    .create();

  // 10分ごとの処理実行トリガー
  ScriptApp.newTrigger('processBatchOnSchedule')
    .timeBased()
    .everyMinutes(10)
    .create();

  // 19時の日次サマリーメール送信トリガー
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(19)
    .nearMinute(0)
    .everyDays(1)
    .create();

  return '文字起こし処理用トリガーを設定しました：\n' +
    '1. 6時開始トリガー\n' +
    '2. 10分ごとの処理トリガー（6:00～24:00の間のみ）\n' +
    '3. 19時の日次サマリーメール送信トリガー';
}

/**
 * Zoom APIトークンを毎日更新するためのトリガーを設定
 */
function setupDailyZoomApiTokenRefresh() {
  try {
    deleteTriggersWithNameContaining('refreshZoomAPIToken');

    // 毎日5:00にトークンを更新
    ScriptApp.newTrigger('refreshZoomAPIToken')
      .timeBased()
      .atHour(5)
      .nearMinute(0)
      .everyDays(1)
      .create();

    return "Zoom APIトークン更新トリガーを設定しました（毎日5:00）";
  } catch (error) {
    logAndNotifyError(error, "APIトークン更新トリガー設定");
    return "エラーが発生しました: " + error.toString();
  }
}

/**
 * Recordingsシートの未処理レコードを処理するトリガーを設定する
 * 30分ごとに実行（平日土日祝日問わず、6:00～24:00の間）
 */
function setupRecordingsSheetTrigger() {
  try {
    deleteTriggersWithNameContaining('processRecordingsFromSheet');

    // 平日土日祝日問わず30分ごとに実行するシンプルなトリガー
    ScriptApp.newTrigger('processRecordingsFromSheet')
      .timeBased()
      .everyMinutes(30)
      .create();

    return "Recordingsシート処理トリガーを設定しました（30分ごと）";
  } catch (error) {
    logAndNotifyError(error, "Recordingsシートトリガー設定");
    return "エラーが発生しました: " + error.toString();
  }
}

/**
 * 全てのトリガーをセットアップする
 * すべてのアプリケーショントリガーを一括で設定
 */
function setupAllTriggers() {
  // 既存のトリガーを削除
  removeAllProjectTriggers();

  // 各機能のトリガーを設定
  setupZoomTriggers();
  setupTranscriptionTriggers();
  setupDailyZoomApiTokenRefresh();
  setupRecordingsSheetTrigger();
  setupInterruptedFilesRecoveryTrigger();
  setupPartialFailureDetectionTrigger();

  Logger.log('全トリガーが正常に設定されました');
}

/**
 * 指定した関数名を含むトリガーを削除するヘルパー関数
 * @param {string} functionNamePart - 削除対象のトリガーに含まれる関数名の一部
 */
function deleteTriggersWithNameContaining(functionNamePart) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction.indexOf(functionNamePart) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * 全てのトリガーを一括削除する
 * 注意: すべてのトリガーがリセットされる
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  return "すべてのトリガーを削除しました。新しいトリガーを設定するには setupAllTriggers() を実行してください。";
}

/**
 * エラーをログ記録し通知するヘルパー関数
 * @param {Error} error - エラーオブジェクト
 * @param {string} processType - 処理タイプの説明
 */
function logAndNotifyError(error, processType) {
  var errorMsg = error.toString();
  Logger.log(processType + "処理中にエラー: " + errorMsg);

  var settings = getSystemSettings();
  if (settings && settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
    var subject = "トリガー設定エラー (" + processType + ")";
    var body = "トリガー設定中にエラーが発生しました。\n\n" +
      "処理タイプ: " + processType + "\n" +
      "日時: " + new Date().toLocaleString() + "\n" +
      "エラー: " + errorMsg + "\n\n" +
      "システム管理者に連絡してください。";

    for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
      GmailApp.sendEmail(settings.ADMIN_EMAILS[i], subject, body);
    }
  }
}

/**
 * 直近1時間の録音を取得する
 */
function fetchLastHourRecordings() {
  return fetchZoomRecordingsManually(1);
}

/**
 * 直近2時間の録音を取得する
 */
function fetchLast2HoursRecordings() {
  return fetchZoomRecordingsManually(2);
}

/**
 * 直近6時間の録音を取得する
 */
function fetchLast6HoursRecordings() {
  return fetchZoomRecordingsManually(6);
}

/**
 * 直近24時間の録音を取得する
 */
function fetchLast24HoursRecordings() {
  return fetchZoomRecordingsManually(24);
}

/**
 * 直近48時間の録音を取得する
 */
function fetchLast48HoursRecordings() {
  return fetchZoomRecordingsManually(48);
}

/**
 * 全ての未処理録音を取得する
 */
function fetchAllPendingRecordings() {
  return fetchZoomRecordingsManually();
}

/**
 * 設定を読み込む関数（参照用）
 */
function getSystemSettings() {
  // 1. EnvironmentConfigを直接使用
  try {
    return {
      ASSEMBLYAI_API_KEY: EnvironmentConfig.get('ASSEMBLYAI_API_KEY', ''),
      OPENAI_API_KEY: EnvironmentConfig.get('OPENAI_API_KEY', ''),
      SOURCE_FOLDER_ID: EnvironmentConfig.get('SOURCE_FOLDER_ID', ''),
      PROCESSING_FOLDER_ID: EnvironmentConfig.get('PROCESSING_FOLDER_ID', ''),
      COMPLETED_FOLDER_ID: EnvironmentConfig.get('COMPLETED_FOLDER_ID', ''),
      ERROR_FOLDER_ID: EnvironmentConfig.get('ERROR_FOLDER_ID', ''),
      ADMIN_EMAILS: EnvironmentConfig.get('ADMIN_EMAILS', []),
      MAX_BATCH_SIZE: EnvironmentConfig.get('MAX_BATCH_SIZE', 10),
      ENHANCE_WITH_OPENAI: EnvironmentConfig.get('ENHANCE_WITH_OPENAI', true),
      ZOOM_CLIENT_ID: EnvironmentConfig.get('ZOOM_CLIENT_ID', ''),
      ZOOM_CLIENT_SECRET: EnvironmentConfig.get('ZOOM_CLIENT_SECRET', ''),
      ZOOM_ACCOUNT_ID: EnvironmentConfig.get('ZOOM_ACCOUNT_ID', ''),
      ZOOM_WEBHOOK_SECRET: EnvironmentConfig.get('ZOOM_WEBHOOK_SECRET', ''),
      RECORDINGS_SHEET_ID: EnvironmentConfig.get('RECORDINGS_SHEET_ID', ''),
      PROCESSED_SHEET_ID: EnvironmentConfig.get('PROCESSED_SHEET_ID', '')
    };
  } catch (e) {
    Logger.log('EnvironmentConfigからの設定取得に失敗: ' + e);

    // 2. Mainモジュールの関数が参照できる場合はそちらを使用
    if (typeof getSystemSettings === 'function' && this !== getSystemSettings) {
      return getSystemSettings();
    }

    // 3. スクリプトプロパティから最低限の設定を読み込む
    var scriptProperties = PropertiesService.getScriptProperties();
    var adminEmail = scriptProperties.getProperty('ADMIN_EMAIL');
    var configSpreadsheetId = scriptProperties.getProperty('CONFIG_SPREADSHEET_ID');

    Logger.log('スクリプトプロパティからの設定取得: ADMIN_EMAIL=' + adminEmail + ', CONFIG_SPREADSHEET_ID=' + configSpreadsheetId);

    // スプレッドシートから設定を取得できる場合
    if (configSpreadsheetId) {
      try {
        var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
        var sheet = spreadsheet.getSheetByName('settings');
        if (sheet) {
          var data = sheet.getDataRange().getValues();
          var config = {};

          // ヘッダー行をスキップし、キーと値のペアを取得
          for (var i = 1; i < data.length; i++) {
            if (data[i][0]) {
              config[data[i][0]] = data[i][1];
            }
          }

          // ADMIN_EMAILSを適切に処理
          if (config.ADMIN_EMAIL) {
            config.ADMIN_EMAILS = String(config.ADMIN_EMAIL)
              .split(',')
              .map(function (email) { return email.trim(); })
              .filter(function (email) { return email !== ''; });
          } else {
            config.ADMIN_EMAILS = adminEmail ? [adminEmail] : [];
          }

          Logger.log('スプレッドシートから設定を取得しました: ' + JSON.stringify(config));
          return config;
        }
      } catch (sheetError) {
        Logger.log('スプレッドシートからの設定取得に失敗: ' + sheetError);
      }
    }

    // どれも失敗した場合は最小限の設定を返す
    return {
      ADMIN_EMAILS: adminEmail ? [adminEmail] : [],
      ASSEMBLYAI_API_KEY: scriptProperties.getProperty('ASSEMBLYAI_API_KEY') || '',
      OPENAI_API_KEY: scriptProperties.getProperty('OPENAI_API_KEY') || ''
    };
  }
}

/**
 * 中断ファイル復旧処理の定期実行トリガーを設定
 */
function setupInterruptedFilesRecoveryTrigger() {
  try {
    // 既存の該当トリガーを削除
    deleteTriggerByFunctionName('recoverInterruptedFiles');

    // 5分ごとに実行するトリガーを作成
    ScriptApp.newTrigger('recoverInterruptedFiles')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log('中断ファイル復旧処理トリガーが正常に設定されました');
  } catch (e) {
    Logger.log('中断ファイル復旧処理トリガーの設定でエラーが発生しました: ' + e.toString());
  }
}

/**
 * 指定した関数名と完全一致するトリガーを削除するヘルパー関数
 * @param {string} functionName - 削除対象のトリガーの関数名（完全一致）
 */
function deleteTriggerByFunctionName(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('トリガーを削除しました: ' + functionName);
    }
  }
}

/**
 * プロジェクトのすべてのトリガーを削除する
 */
function removeAllProjectTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('すべてのプロジェクトトリガーが削除されました');
}

/**
 * 特別復旧処理用のトリガーを設定（一時的な対応用）
 * PENDINGのままのファイルとエラーフォルダのファイルを定期的に再処理
 */
function setupRecoveryTriggers() {
  try {
    // 既存の該当トリガーを削除
    deleteTriggerByFunctionName('resetPendingTranscriptions');
    deleteTriggerByFunctionName('forceRecoverAllErrorFiles');

    // 2時間ごとにPENDINGファイル復旧処理を実行
    ScriptApp.newTrigger('resetPendingTranscriptions')
      .timeBased()
      .everyHours(2)
      .create();

    // 4時間ごとにエラーファイル強制復旧処理を実行
    ScriptApp.newTrigger('forceRecoverAllErrorFiles')
      .timeBased()
      .everyHours(4)
      .create();

    Logger.log('特別復旧処理トリガーが正常に設定されました');
    return '特別復旧処理トリガーが正常に設定されました';
  } catch (e) {
    var errorMsg = '特別復旧処理トリガーの設定でエラーが発生しました: ' + e.toString();
    Logger.log(errorMsg);
    return errorMsg;
  }
}

/**
 * 特別復旧処理用のトリガーを削除
 */
function removeRecoveryTriggers() {
  try {
    // 復旧処理トリガーを削除
    deleteTriggerByFunctionName('resetPendingTranscriptions');
    deleteTriggerByFunctionName('forceRecoverAllErrorFiles');

    Logger.log('特別復旧処理トリガーが正常に削除されました');
    return '特別復旧処理トリガーが正常に削除されました';
  } catch (e) {
    var errorMsg = '特別復旧処理トリガーの削除でエラーが発生しました: ' + e.toString();
    Logger.log(errorMsg);
    return errorMsg;
  }
}

/**
 * 部分的失敗検知・復旧のトリガーを設定する
 * 毎日22:00に実行して、SUCCESSになっているが実際にはエラーが含まれているレコードを検知・復旧
 */
function setupPartialFailureDetectionTrigger() {
  try {
    deleteTriggersWithNameContaining('detectAndRecoverPartialFailures');

    // 毎日22:00に部分的失敗検知・復旧を実行
    ScriptApp.newTrigger('detectAndRecoverPartialFailures')
      .timeBased()
      .atHour(22)
      .nearMinute(0)
      .everyDays(1)
      .create();

    return "部分的失敗検知・復旧トリガーを設定しました（毎日22:00）";
  } catch (error) {
    logAndNotifyError(error, "部分的失敗検知トリガー設定");
    return "エラーが発生しました: " + error.toString();
  }
} 