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
  if (result.error) {
    var settings = getSystemSettings();
    if (settings && settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
      var subject = "ZoomPhone録音取得エラー (" + processType + ")";
      var body = "ZoomPhoneからの録音取得中にエラーが発生しました。\n\n" +
        "処理タイプ: " + processType + "\n" +
        "日時: " + new Date().toLocaleString() + "\n" +
        "エラー: " + result.error + "\n\n" +
        "システム管理者に連絡してください。";

      for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
        GmailApp.sendEmail(settings.ADMIN_EMAILS[i], subject, body);
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
    var subject = "ZoomPhone録音取得エラー (" + processType + ")";
    var body = "ZoomPhoneからの録音取得中にエラーが発生しました。\n\n" +
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
 */
function setupAllTriggers() {
  var results = [
    setupDailyZoomApiTokenRefresh(),
    setupRecordingsSheetTrigger()
  ];

  return results.join("\n");
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

/**
 * 直近1時間の録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchLastHourRecordings() {
  return fetchZoomRecordingsManually(1); // 直近1時間の録音を取得
}

/**
 * 直近2時間の録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchLast2HoursRecordings() {
  return fetchZoomRecordingsManually(2); // 直近2時間の録音を取得
}

/**
 * 直近6時間の録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchLast6HoursRecordings() {
  return fetchZoomRecordingsManually(6); // 直近6時間の録音を取得
}

/**
 * 直近24時間（1日）の録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchLast24HoursRecordings() {
  return fetchZoomRecordingsManually(24); // 直近24時間の録音を取得
}

/**
 * 直近48時間（2日）の録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchLast48HoursRecordings() {
  return fetchZoomRecordingsManually(48); // 直近48時間の録音を取得
}

/**
 * すべての未処理録音を取得する便利関数
 * スクリプトエディタから直接実行可能
 */
function fetchAllPendingRecordings() {
  return fetchZoomRecordingsManually(); // すべての未処理録音を取得
}

/**
 * プロジェクト内の全てのトリガーを削除する関数
 * ※注意: 全てのトリガーが削除されるため、実行後は手動で必要なトリガーを再設定する必要があります
 * @return {string} 削除したトリガーの数と情報
 */
function deleteAllTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var count = triggers.length;
    var deletedTriggers = [];

    // 全てのトリガーをループして削除
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      var handlerFunction = trigger.getHandlerFunction();

      // トリガー情報を記録
      deletedTriggers.push({
        function: handlerFunction,
        type: trigger.getEventType()
      });

      // トリガーを削除
      ScriptApp.deleteTrigger(trigger);
    }

    // 削除したトリガーの情報をログに記録
    Logger.log('削除したトリガー数: ' + count);
    Logger.log('削除したトリガー: ' + JSON.stringify(deletedTriggers));

    return '全てのトリガーを削除しました（' + count + '件）';
  } catch (error) {
    var errorMsg = error.toString();
    Logger.log('トリガー削除中にエラー: ' + errorMsg);
    return 'トリガー削除中にエラー: ' + errorMsg;
  }
}