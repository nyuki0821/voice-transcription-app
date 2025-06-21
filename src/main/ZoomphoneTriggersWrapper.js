/**
 * ZoomPhoneトリガーの設定関連機能は src/main/TriggerManager.js に移行されました。
 * このファイルは既存の依存関係を破壊しないためのラッパーです。
 * 今後は TriggerManager オブジェクトを使用してください。
 */

/**
 * 夜間バッチ: 前日深夜～当日朝までの録音を取得する関数
 * 毎朝6:15に実行（平日・休日問わず）
 */
function fetchZoomRecordingsMorningBatch() {
  try {
    // Recordingsシートから処理を行う
    Logger.log("夜間バッチ: Recordingsシートからの処理を開始します（タイムスタンプが古い順に処理）");
    var result = ZoomphoneProcessor.processRecordingsFromSheet();

    handleProcessingResult(result, "夜間バッチ (Recordingsシート)");

    // より詳細な結果を返す
    return "Zoom録音夜間バッチ完了: " +
      "シートから取得=" + (result.fetched || 0) + "件, " +
      "保存済=" + (result.saved || 0) + "件";
  } catch (error) {
    logAndNotifyError(error, "夜間バッチ (Recordingsシート)");
    return "エラーが発生しました: " + error.toString();
  }
}

/**
 * 過去の録音を手動で一括取得する関数
 * 任意のタイミングで実行可能
 * @param {number} [hours] - 取得する直近の時間（例: 1 = 直近1時間の録音）。省略時は時間フィルタなし
 */
function fetchZoomRecordingsManually(hours) {
  try {
    var isTimeFiltered = typeof hours === 'number' && hours > 0;
    var fromDate, toDate;

    // 時間範囲の設定
    if (isTimeFiltered) {
      var now = new Date();
      fromDate = new Date(now.getTime() - (hours * 60 * 60 * 1000)); // 指定時間前
      toDate = now; // 現在時刻

      // ログ記録
      Logger.log("手動取得: 直近" + hours + "時間の録音処理を開始します（" +
        Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + " 〜 " +
        Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + "）");
    } else {
      // ログ記録
      Logger.log("手動取得: すべての未処理録音を処理します（時間フィルタなし）");
    }

    // Recordingsシートから処理を行う（時間フィルタあり/なし）
    var result = ZoomphoneProcessor.processRecordingsFromSheet(isTimeFiltered ? fromDate : null, isTimeFiltered ? toDate : null);

    // 結果を詳細ログに記録
    Logger.log("取得結果: " + JSON.stringify(result));
    handleProcessingResult(result, isTimeFiltered ? "手動取得（直近" + hours + "時間）" : "手動取得（すべて）");

    // より詳細な結果を返す
    var resultMsg = "Zoom録音手動取得完了: " +
      "シートから取得=" + (result.fetched || 0) + "件, " +
      "保存済=" + (result.saved || 0) + "件";

    if (isTimeFiltered) {
      resultMsg += ", 期間=直近" + hours + "時間";
    }

    return resultMsg;
  } catch (error) {
    logAndNotifyError(error, "手動取得");
    return "エラーが発生しました: " + error.toString();
  }
}

/**
 * 処理結果を処理するヘルパー関数
 * @param {Object} result - 処理結果オブジェクト
 * @param {string} processType - 処理タイプの説明
 */
function handleProcessingResult(result, processType) {
  // エラーが発生した場合の通知
  if (result.error || (result.errorDetails && result.errorDetails.error > 0)) {
    var settings = getSystemSettings();
    if (settings && settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {

      // 統一フォーマットでエラー通知を作成
      var errorDetails;

      if (result.errorDetails) {
        // 既に詳細なエラー情報がある場合はそれを使用
        errorDetails = result.errorDetails;
        errorDetails.processType = processType;
      } else {
        // 簡単なエラー情報の場合は統一フォーマットに変換
        errorDetails = {
          processType: processType,
          total: 1,
          success: 0,
          error: 1,
          errors: [{
            message: result.error,
            errorCode: 'GENERAL_ERROR',
            timestamp: new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
          }]
        };
      }

      for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
        NotificationService.sendUnifiedErrorNotification(settings.ADMIN_EMAILS[i], errorDetails);
      }
    }
  }
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

    // 統一フォーマットでエラー通知を作成
    var errorDetails = {
      processType: processType,
      total: 1,
      success: 0,
      error: 1,
      errors: [{
        message: errorMsg,
        errorCode: 'CRITICAL_ERROR',
        timestamp: new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
      }]
    };

    for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
      NotificationService.sendUnifiedErrorNotification(settings.ADMIN_EMAILS[i], errorDetails);
    }
  }
}

/**
 * 設定を読み込む関数（既存の関数を使用）
 * この関数が存在しない場合は、Main.gsから読み込む
 */
function loadSettings() {
  // getSystemSettings関数があればそれを使用
  if (typeof getSystemSettings === 'function') {
    return getSystemSettings();
  }

  // Main.gsのgetSystemSettings関数を使用
  if (typeof Main !== 'undefined' && typeof Main.getSystemSettings === 'function') {
    return Main.getSystemSettings();
  }

  // スクリプトプロパティから最低限の設定を読み込む
  var scriptProperties = PropertiesService.getScriptProperties();
  var adminEmail = scriptProperties.getProperty('ADMIN_EMAIL');

  return {
    SOURCE_FOLDER_ID: scriptProperties.getProperty('SOURCE_FOLDER_ID') || '',
    ADMIN_EMAILS: adminEmail ? [adminEmail] : []
  };
}

// 以下のトリガー関連関数は下位互換性のために残しています
// 今後は TriggerManager オブジェクトの関数を使用してください

/**
 * ZoomPhone 連携用のトリガー設定関数
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function setupZoomTriggers() {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.setupZoomTriggers === 'function') {
    return TriggerManager.setupZoomTriggers();
  }

  // 下位互換性のために古い実装を残す
  // この部分は新しいTriggerManager.jsがデプロイされると使用されなくなります
  return "トリガー管理機能はTriggerManager.jsに移行されました。";
}

/**
 * Zoom APIトークンを毎日更新するためのトリガーを設定
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function setupDailyZoomApiTokenRefresh() {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.setupDailyZoomApiTokenRefresh === 'function') {
    return TriggerManager.setupDailyZoomApiTokenRefresh();
  }

  // 下位互換性のため
  return "トリガー管理機能はTriggerManager.jsに移行されました。";
}

/**
 * Recordingsシートの未処理レコードを処理するトリガーを設定する
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function setupRecordingsSheetTrigger() {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.setupRecordingsSheetTrigger === 'function') {
    return TriggerManager.setupRecordingsSheetTrigger();
  }

  // 下位互換性のため
  return "トリガー管理機能はTriggerManager.jsに移行されました。";
}

/**
 * 全てのトリガーをセットアップする
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function setupAllTriggers() {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.setupAllTriggers === 'function') {
    return TriggerManager.setupAllTriggers();
  }

  // 下位互換性のため
  return "トリガー管理機能はTriggerManager.jsに移行されました。";
}

/**
 * 指定した関数名を含むトリガーを削除するヘルパー関数
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function deleteTriggersWithNameContaining(functionNamePart) {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.deleteTriggersWithNameContaining === 'function') {
    return TriggerManager.deleteTriggersWithNameContaining(functionNamePart);
  }

  // 下位互換性のために古い実装を呼び出す場合の処理
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
 * 全てのトリガーを削除する
 * 互換性のために維持 - TriggerManager.jsの同名関数を使用します
 */
function deleteAllTriggers() {
  // TriggerManagerの関数を呼び出す
  if (typeof TriggerManager !== 'undefined' && typeof TriggerManager.deleteAllTriggers === 'function') {
    return TriggerManager.deleteAllTriggers();
  }

  // 下位互換性のために古い実装を呼び出す場合の処理
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  return "すべてのトリガーを削除しました。";
}

// 互換性のために以下の関数を残す（実装はそのまま）
function fetchLastHourRecordings() {
  return fetchZoomRecordingsManually(1);
}

function fetchLast2HoursRecordings() {
  return fetchZoomRecordingsManually(2);
}

function fetchLast6HoursRecordings() {
  return fetchZoomRecordingsManually(6);
}

function fetchLast24HoursRecordings() {
  return fetchZoomRecordingsManually(24);
}

function fetchLast48HoursRecordings() {
  return fetchZoomRecordingsManually(48);
}

function fetchAllPendingRecordings() {
  return fetchZoomRecordingsManually();
} 