/**
 * トリガー管理モジュール
 * すべてのアプリケーショントリガーの設定・管理を一括して行う
 */

// TriggerManager名前空間を定義
var TriggerManager = (function () {
  return {
    // 基本トリガー設定
    setupZoomTriggers: setupZoomTriggers,
    setupTranscriptionTriggers: setupTranscriptionTriggers,
    setupDailyZoomApiTokenRefresh: setupDailyZoomApiTokenRefresh,
    setupRecordingsSheetTrigger: setupRecordingsSheetTrigger,
    setupPartialFailureDetectionTrigger: setupPartialFailureDetectionTrigger,

    // 包括的トリガー設定
    setupAllTriggers: setupAllTriggers,
    setupBasicTriggers: setupBasicTriggers,
    setupRecoveryTriggersOnly: setupRecoveryTriggersOnly,

    // 復旧トリガー設定（個別）
    setupRecoveryTriggers: setupRecoveryTriggers,
    removeRecoveryTriggers: removeRecoveryTriggers,
    setupInterruptedFilesRecoveryTrigger: setupInterruptedFilesRecoveryTrigger,

    // トリガー管理
    deleteTriggersWithNameContaining: deleteTriggersWithNameContaining,
    deleteAllTriggers: deleteAllTriggers,

    // 手動録音取得
    fetchLastHourRecordings: fetchLastHourRecordings,
    fetchLast2HoursRecordings: fetchLast2HoursRecordings,
    fetchLast6HoursRecordings: fetchLast6HoursRecordings,
    fetchLast24HoursRecordings: fetchLast24HoursRecordings,
    fetchLast48HoursRecordings: fetchLast48HoursRecordings,
    fetchAllPendingRecordings: fetchAllPendingRecordings
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
 * 全てのトリガーをセットアップする（復旧機能も含む）
 * すべてのアプリケーショントリガーを一括で設定
 * @param {boolean} includeRecovery - 復旧トリガーも含めるかどうか（デフォルト: true）
 */
function setupAllTriggers(includeRecovery) {
  // デフォルトで復旧トリガーも含める
  if (includeRecovery === undefined) {
    includeRecovery = true;
  }

  try {
    Logger.log('全トリガー設定開始（復旧トリガー含む: ' + includeRecovery + '）');

    // 既存のトリガーを削除
    removeAllProjectTriggers();

    // 基本機能のトリガーを設定
    setupZoomTriggers();
    setupTranscriptionTriggers();
    setupDailyZoomApiTokenRefresh();
    setupRecordingsSheetTrigger();
    setupPartialFailureDetectionTrigger();

    // 復旧機能のトリガーを設定（オプション）
    if (includeRecovery) {
      setupInterruptedFilesRecoveryTrigger();
      setupRecoveryTriggers();
      Logger.log('復旧トリガーも設定しました');
    }

    var message = '全トリガーが正常に設定されました（復旧トリガー含む: ' + includeRecovery + '）';
    Logger.log(message);
    return message;
  } catch (error) {
    var errorMessage = '全トリガー設定中にエラー: ' + error.toString();
    Logger.log(errorMessage);
    logAndNotifyError(error, "全トリガー設定");
    return errorMessage;
  }
}

/**
 * 基本トリガーのみをセットアップする（復旧機能は含まない）
 * 通常運用時の最小限のトリガー設定
 */
function setupBasicTriggers() {
  try {
    Logger.log('基本トリガー設定開始');

    // 既存のトリガーを削除
    removeAllProjectTriggers();

    // 基本機能のトリガーのみを設定
    setupZoomTriggers();
    setupTranscriptionTriggers();
    setupDailyZoomApiTokenRefresh();
    setupRecordingsSheetTrigger();
    setupPartialFailureDetectionTrigger();

    var message = '基本トリガーが正常に設定されました（復旧トリガーは除く）';
    Logger.log(message);
    return message;
  } catch (error) {
    var errorMessage = '基本トリガー設定中にエラー: ' + error.toString();
    Logger.log(errorMessage);
    logAndNotifyError(error, "基本トリガー設定");
    return errorMessage;
  }
}

/**
 * 復旧トリガーのみをセットアップする
 * 問題が多発している時の追加設定
 */
function setupRecoveryTriggersOnly() {
  try {
    Logger.log('復旧トリガーのみ設定開始');

    // 復旧関連のトリガーのみを削除
    deleteTriggerByFunctionName('recoverInterruptedFiles');
    deleteTriggerByFunctionName('resetPendingTranscriptions');
    deleteTriggerByFunctionName('forceRecoverAllErrorFiles');

    // 復旧機能のトリガーを設定
    setupInterruptedFilesRecoveryTrigger();
    setupRecoveryTriggers();

    var message = '復旧トリガーのみが正常に設定されました';
    Logger.log(message);
    return message;
  } catch (error) {
    var errorMessage = '復旧トリガー設定中にエラー: ' + error.toString();
    Logger.log(errorMessage);
    logAndNotifyError(error, "復旧トリガー設定");
    return errorMessage;
  }
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
  return ConfigManager.getConfig();
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
 * 見逃しエラー検知（毎日22:00）
 */
function setupPartialFailureDetectionTrigger() {
  try {
    deleteTriggersWithNameContaining('detectAndRecoverPartialFailures');

    // 見逃しエラー検知（毎日22:00）
    ScriptApp.newTrigger('detectAndRecoverPartialFailures')
      .timeBased()
      .everyDays(1)
      .atHour(22)
      .create();
    Logger.log('見逃しエラー検知トリガーを設定しました（毎日22:00）');

    return "見逃しエラー検知トリガーを設定しました（毎日22:00）";
  } catch (error) {
    logAndNotifyError(error, "見逃しエラー検知トリガー設定");
    return "エラーが発生しました: " + error.toString();
  }
} 