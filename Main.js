/**
 * メインコントローラー（OpenAI API対応版）
 * バッチ処理のエントリーポイント
 */

// グローバル変数
var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
var settings = getSystemSettings();
var NOTIFICATION_HOURS = [9, 12, 19]; // 通知を送信する時間（9時、12時、19時）

/**
 * システム設定を直接スプレッドシートから取得する関数
 * グローバル変数に依存せず、常に新しいインスタンスを返す
 */
function getSystemSettings() {
  try {
    // スプレッドシートIDをスクリプトプロパティから取得
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      Logger.log('SPREADSHEET_IDが設定されていません。');
      return getDefaultSettings();
    }

    // スプレッドシートを開く
    var spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      Logger.log('スプレッドシートを開けませんでした: ' + e);
      return getDefaultSettings();
    }

    // システム設定シートを取得
    var settingsSheet = spreadsheet.getSheetByName('システム設定');
    if (!settingsSheet) {
      Logger.log('システム設定シートが見つかりません。');
      return getDefaultSettings();
    }

    // 設定データを読み込む
    var settingsData;
    try {
      settingsData = settingsSheet.getDataRange().getValues();
    } catch (e) {
      Logger.log('設定データの読み込みに失敗しました: ' + e);
      return getDefaultSettings();
    }

    // 設定マップを作成
    var settingsMap = {};
    for (var i = 1; i < settingsData.length; i++) {
      if (settingsData[i].length >= 2) {
        var key = String(settingsData[i][0] || '').trim();
        var value = settingsData[i][1];

        if (key) {
          settingsMap[key] = value;
        }
      }
    }

    // 管理者メールアドレスを処理
    var adminEmails = [];
    if (settingsMap['ADMIN_EMAIL']) {
      var emailStr = String(settingsMap['ADMIN_EMAIL']);
      adminEmails = emailStr.split(',')
        .map(function (email) { return email.trim(); })
        .filter(function (email) { return email !== ''; });
    }

    // 結果オブジェクトを構築
    return {
      ASSEMBLYAI_API_KEY: settingsMap['ASSEMBLYAI_API_KEY'] || '',
      OPENAI_API_KEY: settingsMap['OPENAI_API_KEY'] || '',
      SOURCE_FOLDER_ID: settingsMap['SOURCE_FOLDER_ID'] || '',
      PROCESSING_FOLDER_ID: settingsMap['PROCESSING_FOLDER_ID'] || '',
      COMPLETED_FOLDER_ID: settingsMap['COMPLETED_FOLDER_ID'] || '',
      ERROR_FOLDER_ID: settingsMap['ERROR_FOLDER_ID'] || '',
      ADMIN_EMAILS: adminEmails,
      MAX_BATCH_SIZE: parseInt(settingsMap['MAX_BATCH_SIZE'] || '3', 10),
      ENHANCE_WITH_OPENAI: settingsMap['ENHANCE_WITH_OPENAI'] !== false
    };
  } catch (error) {
    Logger.log('設定の読み込み中にエラー: ' + error);
    return getDefaultSettings();
  }
}

/**
 * デフォルト設定を返す関数
 */
function getDefaultSettings() {
  return {
    ASSEMBLYAI_API_KEY: '',
    OPENAI_API_KEY: '',
    SOURCE_FOLDER_ID: '',
    PROCESSING_FOLDER_ID: '',
    COMPLETED_FOLDER_ID: '',
    ERROR_FOLDER_ID: '',
    ADMIN_EMAILS: [],
    MAX_BATCH_SIZE: 6,
    ENHANCE_WITH_OPENAI: true
  };
}

/**
 * 設定をロードする関数（互換性のために残す - 新しいgetSystemSettings関数を使用）
 */
function loadSettings() {
  return getSystemSettings();
}

/**
 * 通話記録をスプレッドシートに保存する関数
 * @param {Object} callData - 通話データオブジェクト
 */
function saveCallRecordToSheet(callData) {
  try {
    // ファイル名から日付と時間情報を抽出（日本時間に変換）
    var dateTimeInfo = SpreadsheetManager.extractDateTimeFromFileName(callData.fileName);
    var formattedDate = dateTimeInfo.formattedDate;
    var formattedTime = dateTimeInfo.formattedTime;

    // スプレッドシートIDを取得
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_IDが設定されていません');
    }

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('通話記録');

    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet('通話記録');
      // ヘッダー行を設定
      sheet.getRange(1, 1, 1, 12).setValues([
        ['record_id', 'call_date', 'call_time', 'sales_company', 'sales_person', 'customer_company', 'customer_name', 'call_status', 'reason_for_refusal', 'reason_for_appointment', 'summary', 'full_transcript']
      ]);
    }

    // 最終行の次の行に追加
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;

    // record_idをファイル名から抽出した元の時間情報から生成
    var recordId = dateTimeInfo.originalDateTime;

    // データを配列として準備
    var rowData = [
      recordId,
      formattedDate,
      formattedTime,
      callData.salesCompany,
      callData.salesPerson,
      callData.customerCompany,
      callData.customerName,
      callData.callStatus,
      callData.reasonForRefusal,
      callData.reasonForAppointment,
      callData.summary,
      callData.transcription
    ];

    // 行を追加
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    return true;
  } catch (error) {
    Logger.log('saveCallRecordToSheet関数でエラー: ' + error.toString());
    throw error;
  }
}

/**
 * バッチ処理を実行する関数
 */
function processBatch() {
  var startTime = new Date();
  Logger.log('バッチ処理開始: ' + startTime);

  try {
    // 設定を直接取得
    var localSettings = getSystemSettings();

    // APIキーと処理対象フォルダの確認
    if (!localSettings.ASSEMBLYAI_API_KEY) {
      throw new Error('Assembly AI APIキーが設定されていません');
    }

    if (!localSettings.SOURCE_FOLDER_ID) {
      throw new Error('処理対象フォルダIDが設定されていません');
    }

    // 未処理のファイルを取得
    var files = FileProcessor.getUnprocessedFiles(localSettings.SOURCE_FOLDER_ID, localSettings.MAX_BATCH_SIZE);
    Logger.log('処理対象ファイル数: ' + files.length);

    if (files.length === 0) {
      return '処理対象ファイルはありませんでした。';
    }

    // ファイル処理
    var results = {
      success: 0,
      error: 0,
      details: []
    };

    for (var i = 0; i < files.length; i++) {
      var file = files[i];
      var fileStartTime = new Date();
      var fileStartTimeStr = Utilities.formatDate(fileStartTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

      try {
        Logger.log('ファイル処理開始: ' + file.getName());

        // ファイルを処理中フォルダに移動
        FileProcessor.moveFileToFolder(file, localSettings.PROCESSING_FOLDER_ID);

        // 処理ログに記録
        var logId = 'log_' + new Date().getTime() + '_' + i;

        // 文字起こし処理
        var transcriptionResult = TranscriptionService.transcribe(
          file,
          localSettings.ASSEMBLYAI_API_KEY,
          localSettings.OPENAI_API_KEY
        );

        // OpenAI GPT-4o-miniによる情報抽出
        var extractedInfo;

        if (!localSettings.OPENAI_API_KEY) {
          throw new Error('OpenAI APIキーが設定されていないため、処理を続行できません');
        }

        extractedInfo = InformationExtractor.extract(
          transcriptionResult.text,
          transcriptionResult.utterances,
          localSettings.OPENAI_API_KEY
        );

        // 処理終了時間を取得
        var fileEndTime = new Date();
        var fileEndTimeStr = Utilities.formatDate(fileEndTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

        // スプレッドシートに書き込み
        var callData = {
          fileName: file.getName(),
          salesCompany: extractedInfo.salesCompany,
          salesPerson: extractedInfo.salesPerson,
          customerCompany: extractedInfo.customerCompany,
          customerName: extractedInfo.customerName,
          callStatus: extractedInfo.callStatus,
          reasonForRefusal: extractedInfo.reasonForRefusal,
          reasonForAppointment: extractedInfo.reasonForAppointment,
          summary: extractedInfo.summary,
          transcription: transcriptionResult.text
        };

        // 通話記録シートに出力
        saveCallRecordToSheet(callData);

        // 処理ログのみに処理時間を記録
        SpreadsheetManager.logProcessing(
          file.getName(),
          '保存完了',
          fileStartTimeStr,
          fileEndTimeStr
        );

        // ファイルを完了フォルダに移動
        FileProcessor.moveFileToFolder(file, localSettings.COMPLETED_FOLDER_ID);

        // 処理結果を記録
        results.success++;
        results.details.push({
          fileName: file.getName(),
          status: 'success',
          message: '処理が完了しました'
        });
      } catch (error) {
        Logger.log('ファイル処理エラー: ' + error.toString());

        // 処理終了時間を取得（エラー時）
        var fileEndTime = new Date();
        var fileEndTimeStr = Utilities.formatDate(fileEndTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

        // ファイルをエラーフォルダに移動
        try {
          FileProcessor.moveFileToFolder(file, localSettings.ERROR_FOLDER_ID);
        } catch (moveError) {
          Logger.log('ファイルの移動中にエラー: ' + moveError.toString());
        }

        // 処理結果を記録
        results.error++;
        results.details.push({
          fileName: file.getName(),
          status: 'error',
          message: error.toString()
        });

        // エラーログに処理時間も記録
        SpreadsheetManager.logProcessing(file.getName(), 'エラー: ' + error.toString(), fileStartTimeStr, fileEndTimeStr);
      }
    }

    // 処理結果の通知
    var endTime = new Date();
    var processingTime = (endTime - startTime) / 1000; // 秒単位

    var summary = '処理完了: ' + results.success + '件成功, ' + results.error + '件エラー, 処理時間: ' + processingTime + '秒';
    Logger.log(summary);

    return summary;
  } catch (error) {
    var errorMessage = 'バッチ処理エラー: ' + error.toString();
    Logger.log(errorMessage);
    return errorMessage;
  }
}

/**
 * トリガーを設定する関数
 */
function setupTriggers() {
  // 既存のトリガーをすべて削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 9時にプロセスを開始するトリガー
  ScriptApp.newTrigger('startDailyProcess')
    .timeBased()
    .atHour(9)
    .nearMinute(0)
    .everyDays(1)
    .create();

  // 8分ごとの処理実行トリガー
  ScriptApp.newTrigger('processBatchOnSchedule')
    .timeBased()
    .everyMinutes(8)
    .create();

  // 日次サマリー送信用のトリガー
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(18)
    .nearMinute(10)
    .everyDays(1)
    .create();

  return '9時開始トリガー、8分ごとの処理トリガー、18:10の日次サマリートリガーを設定しました。';
}

/**
 * 1日の処理開始関数（9時に実行）
 */
function startDailyProcess() {
  var now = new Date();
  var day = now.getDay(); // 0(日)～6(土)

  // 平日(月～金)のみ実行
  if (day >= 1 && day <= 5) {
    // スクリプトプロパティに処理開始フラグを設定
    PropertiesService.getScriptProperties().setProperty('PROCESSING_ENABLED', 'true');

    // 最初の処理を実行
    return processBatch();
  } else {
    // スクリプトプロパティに処理無効フラグを設定
    PropertiesService.getScriptProperties().setProperty('PROCESSING_ENABLED', 'false');

    return '休業日のため処理をスキップしました。';
  }
}

/**
 * スケジュール実行用の処理関数（時間条件付き）
 */
function processBatchOnSchedule() {
  var now = new Date();
  var day = now.getDay(); // 0(日)～6(土)
  var hour = now.getHours();

  // 処理有効フラグを確認
  var processingEnabled = PropertiesService.getScriptProperties().getProperty('PROCESSING_ENABLED');

  // 平日(月～金)の9時～21時の間のみ実行かつ処理フラグが有効
  if (day >= 1 && day <= 5 && hour >= 9 && hour < 21 && processingEnabled === 'true') {
    return processBatch();
  } else if (hour >= 21 || hour < 9) {
    // 業務時間外の場合は処理フラグをリセット
    if (hour >= 21) {
      PropertiesService.getScriptProperties().setProperty('PROCESSING_ENABLED', 'false');
    }

    return '業務時間外のため処理をスキップしました。';
  } else {
    return '処理フラグが無効または休業日のため処理をスキップしました。';
  }
}

/**
 * 日次サマリー集計と送信を行う関数
 */
function sendDailySummary() {
  try {
    // 設定を直接取得 (既存のloadSettings関数に依存しない)
    var settings = getSystemSettings();

    // スプレッドシートIDを取得
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      return 'SPREADSHEET_IDが設定されていません';
    }

    // スプレッドシートを開く
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var logSheet = spreadsheet.getSheetByName('処理ログ');

    if (!logSheet) {
      return '処理ログシートが見つかりません';
    }

    // 現在の日付を取得
    var today = new Date();
    var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

    // スプレッドシートの表示形式の日付（YYYY/MM/DD）
    var todayYMD = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

    // ヘッダー行とカラムインデックスを特定
    var allData = logSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      return '処理ログシートにデータがありません';
    }

    var headerRow = allData[0];

    var fileNameIdx = headerRow.indexOf('file_name');
    var statusIdx = headerRow.indexOf('status');
    var processStartIdx = headerRow.indexOf('process_start');
    var processEndIdx = headerRow.indexOf('process_end');

    // カラム名が異なる場合の対応
    if (fileNameIdx === -1) fileNameIdx = headerRow.indexOf('ファイル名');
    if (statusIdx === -1) statusIdx = headerRow.indexOf('ステータス');
    if (processStartIdx === -1) processStartIdx = headerRow.indexOf('処理開始');
    if (processEndIdx === -1) processEndIdx = headerRow.indexOf('処理終了');

    // 処理ログシートのすべてのセルの表示値を取得
    var displayValues = logSheet.getDataRange().getDisplayValues();

    // 当日のログを抽出（表示値を使用）
    var todayLogs = [];
    var todayValues = [];
    var todayDisplayValues = [];

    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      var displayRow = displayValues[i];
      var fileName = (fileNameIdx >= 0) ? row[fileNameIdx] : "不明なファイル";

      // プロセス開始時間の表示値からYYYY/MM/DD部分を抽出
      if (processStartIdx >= 0) {
        var displayDateStr = displayRow[processStartIdx];
        var dateOnly = '';

        // YYYY/MM/DD または YYYY-MM-DD 形式を抽出
        if (typeof displayDateStr === 'string') {
          var match = displayDateStr.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
          if (match) {
            dateOnly = match[0].replace(/-/g, '/'); // ハイフンをスラッシュに統一
          }
        }

        // 抽出した日付部分が今日の日付と一致するか確認
        if (dateOnly && (dateOnly === todayYMD || dateOnly === todayStr.replace(/-/g, '/'))) {
          todayLogs.push(row);
          todayValues.push(row);
          todayDisplayValues.push(displayRow);
        }
      }
    }

    // 集計結果を作成
    var successCount = 0;
    var errorCount = 0;
    var totalProcessingTime = 0;
    var fileDetails = [];
    var processedFiles = {};

    for (var i = 0; i < todayLogs.length; i++) {
      var row = todayValues[i];
      var displayRow = todayDisplayValues[i];
      var fileName = fileNameIdx >= 0 ? row[fileNameIdx] : '不明なファイル';
      var status = statusIdx >= 0 ? row[statusIdx] : '不明なステータス';

      // 処理時間の計算 - 表示値から直接時間を抽出して計算
      var processingTime = 0;
      if (processStartIdx >= 0 && processEndIdx >= 0) {
        var startTimeDisplay = displayRow[processStartIdx];
        var endTimeDisplay = displayRow[processEndIdx];

        if (startTimeDisplay && endTimeDisplay) {
          // 時間部分を抽出 (HH:MM:SS)
          var startTimeMatch = startTimeDisplay.match(/(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
          var endTimeMatch = endTimeDisplay.match(/(\d{1,2}):(\d{1,2}):(\d{1,2})$/);

          if (startTimeMatch && endTimeMatch) {
            // 秒に変換して差分を計算
            var startSeconds = parseInt(startTimeMatch[1]) * 3600 +
              parseInt(startTimeMatch[2]) * 60 +
              parseInt(startTimeMatch[3]);

            var endSeconds = parseInt(endTimeMatch[1]) * 3600 +
              parseInt(endTimeMatch[2]) * 60 +
              parseInt(endTimeMatch[3]);

            processingTime = endSeconds - startSeconds;

            // 日付が変わったケース（例: 23:59:59 → 00:00:05）の対応
            if (processingTime < 0) {
              processingTime += 24 * 3600; // 24時間分の秒数を追加
            }
          }
        }
      }

      // ステータス判定
      var isSuccess = false;
      var isError = false;

      if (status) {
        var statusStr = String(status).toLowerCase();
        isSuccess = statusStr.indexOf('完了') !== -1 ||
          statusStr.indexOf('成功') !== -1 ||
          statusStr.indexOf('保存') !== -1 ||
          statusStr === 'ok' ||
          statusStr === 'success';

        isError = statusStr.indexOf('エラー') !== -1 ||
          statusStr.indexOf('失敗') !== -1 ||
          statusStr === 'error' ||
          statusStr === 'fail' ||
          statusStr === 'failed';
      }

      if (isSuccess) {
        successCount++;
      } else if (isError) {
        errorCount++;
      }

      // ファイル詳細を追加/更新（同一ファイル名の場合は最新の情報を保持）
      if (!processedFiles[fileName]) {
        processedFiles[fileName] = {
          fileName: fileName,
          status: status,
          processingTime: processingTime
        };
        fileDetails.push(processedFiles[fileName]);
      } else {
        processedFiles[fileName].status = status;
        if (processingTime > 0) {
          processedFiles[fileName].processingTime = processingTime;
        }
      }

      // 処理時間の合計に加算
      if (processingTime > 0) {
        totalProcessingTime += processingTime;
      }
    }

    // 平均処理時間を計算
    var avgProcessingTime = 0;
    if (fileDetails.length > 0) {
      avgProcessingTime = totalProcessingTime / fileDetails.length;
    }

    // サマリー結果を作成
    var results = {
      success: successCount,
      error: errorCount,
      details: fileDetails,
      date: todayStr,
      totalProcessingTime: totalProcessingTime,
      avgProcessingTime: avgProcessingTime
    };

    // 管理者に通知
    if (settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
      // 各管理者に通知を送信
      settings.ADMIN_EMAILS.forEach(function (email) {
        sendEnhancedDailySummary(email, results, todayStr);
      });

      return '日次サマリーを送信しました: ' + todayStr + ' (成功=' + successCount +
        ', エラー=' + errorCount + ', 合計処理時間=' + formatTime(totalProcessingTime) + ')';
    } else {
      return '通知先メールアドレスが設定されていません';
    }
  } catch (error) {
    Logger.log('日次サマリー集計中にエラー: ' + error.toString());
    return 'エラーが発生しました: ' + error.toString();
  }
}

/**
 * 拡張日次サマリーメールを送信する関数
 * @param {string} email - 送信先メールアドレス
 * @param {Object} results - 集計結果
 * @param {string} dateStr - 日付文字列
 */
function sendEnhancedDailySummary(email, results, dateStr) {
  if (!email) {
    throw new Error('送信先メールアドレスが指定されていません');
  }

  var subject = '顧客会話自動文字起こしシステム - ' + dateStr + ' 日次処理結果サマリー';

  var body = dateStr + ' の処理結果サマリー\n\n' +
    '成功: ' + results.success + '件\n' +
    'エラー: ' + results.error + '件\n';

  // 処理時間情報を追加
  if (results.totalProcessingTime > 0) {
    body += '合計処理時間: ' + formatTime(results.totalProcessingTime) + '\n';
    body += '平均処理時間: ' + formatTime(results.avgProcessingTime) + '\n';
  }

  body += '\n';

  if (results.details && results.details.length > 0) {
    body += '処理ファイル一覧:\n';
    for (var i = 0; i < results.details.length; i++) {
      var detail = results.details[i];
      var detailText = '- ' + detail.fileName + ': ' + detail.status;

      // 処理時間情報を追加（存在する場合）
      if (detail.processingTime && detail.processingTime > 0) {
        detailText += ' (処理時間: ' + formatTime(detail.processingTime) + ')';
      }

      body += detailText + '\n';
    }
  } else {
    body += '本日処理されたファイルはありません。\n';
  }

  // メール送信
  GmailApp.sendEmail(email, subject, body);
}

/**
 * 秒数を読みやすい時間形式に変換する関数
 * @param {number} seconds - 秒数
 * @return {string} - フォーマットされた時間文字列
 */
function formatTime(seconds) {
  seconds = Math.round(seconds);

  if (seconds < 60) {
    return seconds + '秒';
  } else if (seconds < 3600) {
    var minutes = Math.floor(seconds / 60);
    var remainingSeconds = seconds % 60;
    return minutes + '分' + (remainingSeconds > 0 ? remainingSeconds + '秒' : '');
  } else {
    var hours = Math.floor(seconds / 3600);
    var remainingMinutes = Math.floor((seconds % 3600) / 60);
    var remainingSeconds = seconds % 60;

    var result = hours + '時間';
    if (remainingMinutes > 0) {
      result += remainingMinutes + '分';
    }
    if (remainingSeconds > 0) {
      result += remainingSeconds + '秒';
    }

    return result;
  }
}

/**
 * 新しいsendDailySummary関数をテストする関数
 */
function testNewSendDailySummary() {
  try {
    Logger.log("新しいsendDailySummaryのテスト開始");

    // 設定が正しく読み込めるか確認
    var settings = getSystemSettings();
    Logger.log("設定読み込み結果: ADMIN_EMAILS = " + JSON.stringify(settings.ADMIN_EMAILS));

    // 関数を呼び出し、結果を取得
    var result = sendDailySummary();

    // ログにテスト結果を出力
    Logger.log("テスト結果: " + result);

    return "テストが完了しました: " + result;
  } catch (error) {
    // エラー発生時の詳細なログ
    Logger.log("testNewSendDailySummaryでエラー: " + error.toString());

    if (error.stack) {
      Logger.log("エラースタック: " + error.stack);
    }

    return "テスト中にエラーが発生しました: " + error.toString();
  }
}
