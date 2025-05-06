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
 * 全てのトリガーをセットアップする
 * すべてのアプリケーショントリガーを一括で設定
 */
function setupAllTriggers() {
  var results = [
    setupDailyZoomApiTokenRefresh(),
    setupRecordingsSheetTrigger(),
    setupZoomTriggers(),
    setupTranscriptionTriggers()
  ];

  return results.join("\n");
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
  // Mainモジュールの関数を参照
  if (typeof Main !== 'undefined' && typeof Main.getSystemSettings === 'function') {
    return Main.getSystemSettings();
  }

  // グローバルに定義されている場合
  if (typeof window !== 'undefined' && typeof window.getSystemSettings === 'function') {
    return window.getSystemSettings();
  }

  // スクリプトプロパティから最低限の設定を読み込む
  var scriptProperties = PropertiesService.getScriptProperties();
  var adminEmail = scriptProperties.getProperty('ADMIN_EMAIL');

  return {
    ADMIN_EMAILS: adminEmail ? [adminEmail] : []
  };
} 