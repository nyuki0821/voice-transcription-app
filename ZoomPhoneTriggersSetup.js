/**
 * ZoomPhone 連携用のトリガー設定関数
 * ZoomPhone APIを使って自動的に録音ファイルを取得・処理するためのトリガーを設定
 */
function setupZoomTriggers() {
  // 既存のZoom関連のトリガーをすべて削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction === 'fetchZoomRecordings' ||
      handlerFunction === 'fetchZoomRecordingsEvery30Min' ||
      handlerFunction === 'fetchZoomRecordingsMorningBatch' ||
      handlerFunction === 'fetchZoomRecordingsWeekendBatch' ||
      handlerFunction === 'checkAndFetchZoomRecordings' ||
      handlerFunction === 'purgeOldRecordings') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 1. 30分ごとの定期チェック（共通の時間チェック関数を使用）
  // 時間帯チェック機能を持った関数を30分ごとに実行するトリガーを1つだけ設定
  ScriptApp.newTrigger('checkAndFetchZoomRecordings')
    .timeBased()
    .everyMinutes(30)
    .create();

  // 2. 朝7:15の夜間バッチ処理（前日22時～当日朝までの録音）- 毎日実行
  ScriptApp.newTrigger('fetchZoomRecordingsMorningBatch')
    .timeBased()
    .atHour(7)
    .nearMinute(15)
    .everyDays(1)
    .create();

  // 3. 週末バッチ: 月曜9:10 (金曜21時～月曜朝) に実行
  ScriptApp.newTrigger('fetchZoomRecordingsWeekendBatch')
    .timeBased()
    .atHour(9)
    .nearMinute(10)
    .everyDays(1) // 月曜チェック処理内で曜日判定
    .create();

  // 4. リテンションバッチ: 日曜 03:00 に古いファイルを削除
  ScriptApp.newTrigger('purgeOldRecordings')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .nearMinute(0)
    .create();

  return '以下のZoom録音取得トリガーを設定しました：\n' +
    '1. 定期取得: 30分ごとに実行（毎日7:00～22:00の間のみ処理）\n' +
    '2. 夜間バッチ: 毎朝7:15（前日22時～当日朝までの録音）\n' +
    '3. 週末バッチ: 月曜9:10 (金曜21時～月曜朝)\n' +
    '4. リテンションバッチ: 日曜 03:00 に古いファイルを削除';
}

/**
 * 30分ごとに実行され、条件に合致する場合のみ録音取得する関数
 * トリガーは30分ごとに実行されるが、実際の処理は平日・休日の7:00～22:00の間のみ
 */
function checkAndFetchZoomRecordings() {
  var now = new Date();
  var hour = now.getHours();

  // 平日・休日問わず7:00～22:00の間のみ実行
  if (hour >= 7 && hour < 22) {
    try {
      // 2時間前の時刻
      var fromDate = new Date(now.getTime() - (2 * 60 * 60 * 1000));
      var toDate = new Date(); // 現在時刻

      // 処理実行と結果取得
      var result = ZoomphoneProcessor.processRecordings(fromDate, toDate);

      // ログ記録
      Logger.log("定期取得: " + Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
        " ～ " + Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
      Logger.log("取得結果: " + JSON.stringify(result));

      handleProcessingResult(result, "定期取得");
      return "定期取得が完了しました。取得件数: " + (result.saved || 0) + "件";
    } catch (error) {
      logAndNotifyError(error, "定期取得");
      return "エラーが発生しました: " + error.toString();
    }
  } else {
    // 営業時間外は何もしない（ログも残さない）
    return "営業時間外（7:00～22:00以外）のため処理をスキップしました";
  }
}

/**
 * 定期取得用: 30分ごとに過去2時間分の録音を取得する関数
 * 平日8:45～21:15の間に実行（30分おき）
 * ※この関数は下位互換性のために残しておく
 */
function fetchZoomRecordingsEvery30Min() {
  return checkAndFetchZoomRecordings();
}

/**
 * 夜間バッチ: 前日22時～当日朝までの録音を取得する関数
 * 毎朝7:15に実行（平日・休日問わず）
 */
function fetchZoomRecordingsMorningBatch() {
  try {
    var now = new Date();

    // 前日22時
    var fromDate = new Date(now);
    fromDate.setDate(fromDate.getDate() - 1);
    fromDate.setHours(22, 0, 0, 0);

    // 当日朝
    var toDate = new Date(now);
    toDate.setHours(7, 0, 0, 0);

    // 処理実行と結果取得
    var result = ZoomphoneProcessor.processRecordings(fromDate, toDate);

    // ログ記録
    Logger.log("夜間バッチ: " + Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
      " ～ " + Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
    Logger.log("取得結果: " + JSON.stringify(result));

    handleProcessingResult(result, "夜間バッチ");
    return "夜間バッチが完了しました。取得件数: " + (result.saved || 0) + "件";
  } catch (error) {
    logAndNotifyError(error, "夜間バッチ");
    return "エラーが発生しました: " + error.toString();
  }
}

/**
 * 週末バッチ: 金曜21時～月曜朝までの録音を取得する関数
 * 月曜9:10に実行
 */
function fetchZoomRecordingsWeekendBatch() {
  var now = new Date();
  var day = now.getDay(); // 0(日)～6(土)

  // 月曜日のみ実行
  if (day === 1) {
    try {
      // 先週金曜日の21時
      var fromDate = new Date(now);
      fromDate.setDate(fromDate.getDate() - 3); // 3日前=金曜日
      fromDate.setHours(21, 0, 0, 0);

      // 当日朝
      var toDate = new Date(now);
      toDate.setHours(9, 0, 0, 0);

      // 処理実行と結果取得
      var result = ZoomphoneProcessor.processRecordings(fromDate, toDate);

      // ログ記録
      Logger.log("週末バッチ: " + Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
        " ～ " + Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
      Logger.log("取得結果: " + JSON.stringify(result));

      handleProcessingResult(result, "週末バッチ");
      return "週末バッチが完了しました。取得件数: " + (result.saved || 0) + "件";
    } catch (error) {
      logAndNotifyError(error, "週末バッチ");
      return "エラーが発生しました: " + error.toString();
    }
  } else {
    return "月曜日以外は実行しません";
  }
}

/**
 * 過去x時間分の録音を一括取得する手動実行関数
 * 任意のタイミングで実行可能
 * @param {number} hours - 取得する過去の時間（デフォルト48時間）
 */
function fetchZoomRecordingsManually(hours) {
  try {
    var hoursToFetch = hours || 48; // デフォルト48時間

    // 指定時間前から現在までの録音を取得
    var now = new Date();
    var fromDate = new Date(now.getTime() - (hoursToFetch * 60 * 60 * 1000));
    var toDate = now;

    // ログ記録
    Logger.log("手動取得: 過去" + hoursToFetch + "時間分 " +
      Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
      " ～ " + Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));

    // 処理実行と結果取得
    var result = ZoomphoneProcessor.processRecordings(fromDate, toDate);

    handleProcessingResult(result, "手動取得(過去" + hoursToFetch + "時間)");
    return "手動取得が完了しました。取得件数: " + (result.saved || 0) + "件";
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