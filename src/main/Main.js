/**
 * メインコントローラー（OpenAI API対応版）
 * バッチ処理のエントリーポイント
 */

// グローバル変数
var SPREADSHEET_ID = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
var settings = getSystemSettings();
var NOTIFICATION_HOURS = [9, 12, 19]; // 通知を送信する時間（9時、12時、19時）

/**
 * システム設定を取得する関数
 * 環境設定ファイルから一元的に取得
 */
function getSystemSettings() {
  try {
    // 環境設定を取得
    var config = EnvironmentConfig.getConfig();

    // 結果オブジェクトを構築
    return {
      ASSEMBLYAI_API_KEY: config.ASSEMBLYAI_API_KEY || '',
      OPENAI_API_KEY: config.OPENAI_API_KEY || '',
      SOURCE_FOLDER_ID: config.SOURCE_FOLDER_ID || '',
      PROCESSING_FOLDER_ID: config.PROCESSING_FOLDER_ID || '',
      COMPLETED_FOLDER_ID: config.COMPLETED_FOLDER_ID || '',
      ERROR_FOLDER_ID: config.ERROR_FOLDER_ID || '',
      ADMIN_EMAILS: config.ADMIN_EMAILS || [],
      MAX_BATCH_SIZE: config.MAX_BATCH_SIZE || 10,
      ENHANCE_WITH_OPENAI: config.ENHANCE_WITH_OPENAI !== false,
      ZOOM_CLIENT_ID: config.ZOOM_CLIENT_ID || '',
      ZOOM_CLIENT_SECRET: config.ZOOM_CLIENT_SECRET || '',
      ZOOM_ACCOUNT_ID: config.ZOOM_ACCOUNT_ID || '',
      ZOOM_WEBHOOK_SECRET: config.ZOOM_WEBHOOK_SECRET || '',
      RECORDINGS_SHEET_ID: config.RECORDINGS_SHEET_ID || '',
      PROCESSED_SHEET_ID: config.PROCESSED_SHEET_ID || '',
      LOG_SHEET_ID: config.LOG_SHEET_ID || ''
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
    MAX_BATCH_SIZE: 10,
    ENHANCE_WITH_OPENAI: true,
    ZOOM_CLIENT_ID: '',
    ZOOM_CLIENT_SECRET: '',
    ZOOM_ACCOUNT_ID: '',
    ZOOM_WEBHOOK_SECRET: '',
    RECORDINGS_SHEET_ID: '',
    PROCESSED_SHEET_ID: '',
    LOG_SHEET_ID: ''
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
 * @param {string} [targetSpreadsheetId] - 保存先のスプレッドシートID（省略時はRECORDINGS_SHEET_ID）
 * @param {string} [sheetName] - 保存先のシート名（省略時は'call_records'）
 */
function saveCallRecordToSheet(callData, targetSpreadsheetId, sheetName) {
  try {
    // デフォルト値の設定
    targetSpreadsheetId = targetSpreadsheetId || EnvironmentConfig.get('PROCESSED_SHEET_ID', '');
    sheetName = sheetName || 'call_records';

    Logger.log('saveCallRecordToSheet開始: ファイル名=' + callData.fileName);
    Logger.log('保存先スプレッドシートID=' + targetSpreadsheetId);
    Logger.log('保存先シート名=' + sheetName);

    // ファイル名から日付と時間情報を抽出（日本時間に変換）
    var dateTimeInfo = SpreadsheetManager.extractDateTimeFromFileName(callData.fileName);
    Logger.log('日時抽出結果: ' + JSON.stringify(dateTimeInfo));

    var formattedDate = dateTimeInfo.formattedDate;
    var formattedTime = dateTimeInfo.formattedTime;

    // スプレッドシートIDを取得
    if (!targetSpreadsheetId) {
      throw new Error('保存先スプレッドシートIDが設定されていません');
    }

    try {
      var spreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
      Logger.log('スプレッドシート取得: 成功 - 名前=' + spreadsheet.getName());
    } catch (ssError) {
      Logger.log('スプレッドシート取得エラー: ' + ssError.toString());
      throw new Error('スプレッドシート取得中にエラー: ' + ssError.toString());
    }

    var sheet = spreadsheet.getSheetByName(sheetName);
    Logger.log('シート取得: ' + (sheet ? '成功' : '失敗'));

    if (!sheet) {
      // シートが存在しない場合は作成
      Logger.log('シートが存在しないため新規作成します: ' + sheetName);
      sheet = spreadsheet.insertSheet(sheetName);
      // ヘッダー行を設定
      sheet.getRange(1, 1, 1, 12).setValues([
        ['record_id', 'call_date', 'call_time', 'sales_company', 'sales_person', 'customer_company', 'customer_name', 'call_status', 'reason_for_refusal', 'reason_for_appointment', 'summary', 'full_transcript']
      ]);
    }

    // 最終行の次の行に追加
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    Logger.log('挿入行: ' + newRow);

    // record_idをファイル名から抽出した元の時間情報から生成
    var recordId = dateTimeInfo.originalDateTime;

    // データを配列として準備
    var rowData = [
      recordId,
      formattedDate,
      formattedTime,
      callData.salesCompany || '',
      callData.salesPerson || '',
      callData.customerCompany || '',
      callData.customerName || '',
      callData.callStatus || '',
      callData.reasonForRefusal || '',
      callData.reasonForAppointment || '',
      callData.summary || '',
      callData.transcription || ''
    ];

    Logger.log('挿入データ準備完了: record_id=' + recordId);

    // 行を追加
    try {
      sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
      Logger.log('行の追加に成功しました');
    } catch (writeError) {
      Logger.log('行の追加中にエラー: ' + writeError.toString());
      // エラーをそのまま上位に投げる
      throw writeError;
    }

    Logger.log('saveCallRecordToSheet完了');
    return true;
  } catch (error) {
    Logger.log('saveCallRecordToSheet関数でエラー: ' + error.toString());
    // エラースタックがある場合は出力
    if (error.stack) {
      Logger.log('エラースタック: ' + error.stack);
    }
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

    // ファイル処理トラッキング
    var results = {
      total: files.length,      // 対象ファイル総数
      processed: 0,             // 処理試行数 
      success: 0,               // 成功数
      error: 0,                 // エラー数
      skipped: 0,               // スキップ数
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

        // 文字起こし結果のチェック
        Logger.log('文字起こし結果: 文字数=' + (transcriptionResult.text ? transcriptionResult.text.length : 0));
        Logger.log('発話数: ' + (transcriptionResult.utterances ? transcriptionResult.utterances.length : 0));

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

        // 抽出情報のチェック
        Logger.log('情報抽出結果: ' + JSON.stringify({
          salesCompany: extractedInfo.salesCompany,
          salesPerson: extractedInfo.salesPerson,
          customerCompany: extractedInfo.customerCompany,
          callStatus: extractedInfo.callStatus,
          summaryLength: extractedInfo.summary ? extractedInfo.summary.length : 0
        }));

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

        // call_recordsシートに出力
        var processedSheetId = localSettings.PROCESSED_SHEET_ID || '';
        Logger.log('文字起こし結果の保存先: PROCESSED_SHEET_ID=' + processedSheetId);
        saveCallRecordToSheet(callData, processedSheetId, 'call_records');

        // ファイル名から録音IDを抽出してみる
        var recordingId = null;
        var fileName = file.getName();

        // 複数のUUIDパターンを試す
        // パターン1: ハイフン区切りのUUID
        var uuidWithHyphens = fileName.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);

        // パターン2: アンダースコア区切りのUUID
        var uuidWithUnderscores = fileName.match(/([0-9a-f]{8}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{12})/i);

        // パターン3: 連続した32文字の16進数
        var uuidWithoutSeparator = fileName.match(/_([0-9a-f]{32})[^0-9a-f]/i);

        // どれかにマッチするものを使用
        if (uuidWithHyphens) {
          recordingId = uuidWithHyphens[1];
        } else if (uuidWithUnderscores) {
          recordingId = uuidWithUnderscores[1];
        } else if (uuidWithoutSeparator) {
          recordingId = uuidWithoutSeparator[1];
        } else {
          // 最後の手段: ファイル名から _YYYYMMDDHHMMSS_ の後の部分を抽出
          var lastPartMatch = fileName.match(/_(\d{14})_([0-9a-f]+)\./i);
          if (lastPartMatch) {
            recordingId = lastPartMatch[2];
          }
        }

        // 処理ログのみに処理時間を記録 (録音IDも含める)
        SpreadsheetManager.logProcessing(
          file.getName(),
          '保存完了',
          fileStartTimeStr,
          fileEndTimeStr,
          recordingId
        );

        // ファイルを完了フォルダに移動
        FileProcessor.moveFileToFolder(file, localSettings.COMPLETED_FOLDER_ID);

        // 処理カウント更新
        results.processed++;
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

        // 処理カウント更新
        results.processed++;
        // 処理結果を記録
        results.error++;
        results.details.push({
          fileName: file.getName(),
          status: 'error',
          message: error.toString()
        });

        // ファイル名から録音IDを抽出してみる
        var recordingId = null;
        var fileName = file.getName();

        // 複数のUUIDパターンを試す
        // パターン1: ハイフン区切りのUUID
        var uuidWithHyphens = fileName.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);

        // パターン2: アンダースコア区切りのUUID
        var uuidWithUnderscores = fileName.match(/([0-9a-f]{8}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{12})/i);

        // パターン3: 連続した32文字の16進数
        var uuidWithoutSeparator = fileName.match(/_([0-9a-f]{32})[^0-9a-f]/i);

        // どれかにマッチするものを使用
        if (uuidWithHyphens) {
          recordingId = uuidWithHyphens[1];
        } else if (uuidWithUnderscores) {
          recordingId = uuidWithUnderscores[1];
        } else if (uuidWithoutSeparator) {
          recordingId = uuidWithoutSeparator[1];
        } else {
          // 最後の手段: ファイル名から _YYYYMMDDHHMMSS_ の後の部分を抽出
          var lastPartMatch = fileName.match(/_(\d{14})_([0-9a-f]+)\./i);
          if (lastPartMatch) {
            recordingId = lastPartMatch[2];
          }
        }

        // エラーログに処理時間も記録
        SpreadsheetManager.logProcessing(file.getName(), 'エラー: ' + error.toString(), fileStartTimeStr, fileEndTimeStr, recordingId);
      }
    }

    // 処理結果の通知
    var endTime = new Date();
    var processingTime = (endTime - startTime) / 1000; // 秒単位

    var summary = '文字起こし処理完了: ' +
      '対象=' + results.total + '件, ' +
      '処理=' + results.processed + '件, ' +
      '成功=' + results.success + '件, ' +
      '失敗=' + results.error + '件, ' +
      '処理時間=' + processingTime + '秒';
    Logger.log(summary);

    return summary;
  } catch (error) {
    var errorMessage = 'バッチ処理エラー: ' + error.toString();
    Logger.log(errorMessage);
    return errorMessage;
  }
}

/**
 * 文字起こし処理用のトリガーを設定する関数
 */
function setupTranscriptionTriggers() {
  // 既存のトリガーをすべて削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var handlerFunction = trigger.getHandlerFunction();
    // 自分が管理するトリガーのみ削除
    if (handlerFunction === 'startDailyProcess' ||
      handlerFunction === 'processBatchOnSchedule' ||
      handlerFunction === 'sendDailySummary') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 7時にプロセスを開始するトリガー
  ScriptApp.newTrigger('startDailyProcess')
    .timeBased()
    .atHour(7)
    .nearMinute(0)
    .everyDays(1)
    .create();

  // 10分ごとの処理実行トリガー
  ScriptApp.newTrigger('processBatchOnSchedule')
    .timeBased()
    .everyMinutes(10)
    .create();

  // 日次サマリー送信用のトリガー
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(18)
    .nearMinute(10)
    .everyDays(1)
    .create();

  return '文字起こし処理用トリガーを設定しました：\n' +
    '1. 7時開始トリガー\n' +
    '2. 10分ごとの処理トリガー\n' +
    '3. 18:10の日次サマリートリガー';
}

/**
 * 処理中フォルダに残ったファイルを未処理フォルダに戻す関数
 * タイムアウトなどで処理が中断された場合のリカバリー用
 * @return {string} - 処理結果のメッセージ
 */
function recoverStuckFiles() {
  try {
    // 設定を取得
    var localSettings = getSystemSettings();

    // 処理中フォルダとソースフォルダのIDを確認
    if (!localSettings.PROCESSING_FOLDER_ID || !localSettings.SOURCE_FOLDER_ID) {
      return 'フォルダIDが設定されていません';
    }

    // 処理中フォルダ内のファイルを取得
    var processingFolder = DriveApp.getFolderById(localSettings.PROCESSING_FOLDER_ID);
    var files = processingFolder.getFiles();

    var movedCount = 0;
    var fileList = [];

    // 処理中フォルダ内のファイルをソースフォルダに移動
    while (files.hasNext()) {
      var file = files.next();
      fileList.push(file.getName());

      try {
        FileProcessor.moveFileToFolder(file, localSettings.SOURCE_FOLDER_ID);
        movedCount++;

        // ファイル名から録音IDを抽出してみる
        var recordingId = null;
        var fileName = file.getName();

        // 複数のUUIDパターンを試す
        // パターン1: ハイフン区切りのUUID
        var uuidWithHyphens = fileName.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);

        // パターン2: アンダースコア区切りのUUID
        var uuidWithUnderscores = fileName.match(/([0-9a-f]{8}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{4}_[0-9a-f]{12})/i);

        // パターン3: 連続した32文字の16進数
        var uuidWithoutSeparator = fileName.match(/_([0-9a-f]{32})[^0-9a-f]/i);

        // どれかにマッチするものを使用
        if (uuidWithHyphens) {
          recordingId = uuidWithHyphens[1];
        } else if (uuidWithUnderscores) {
          recordingId = uuidWithUnderscores[1];
        } else if (uuidWithoutSeparator) {
          recordingId = uuidWithoutSeparator[1];
        } else {
          // 最後の手段: ファイル名から _YYYYMMDDHHMMSS_ の後の部分を抽出
          var lastPartMatch = fileName.match(/_(\d{14})_([0-9a-f]+)\./i);
          if (lastPartMatch) {
            recordingId = lastPartMatch[2];
          }
        }

        // 処理ログに記録
        SpreadsheetManager.logProcessing(
          file.getName(),
          'タイムアウトリカバリー：未処理フォルダに戻しました',
          Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
          Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
          recordingId
        );
      } catch (moveError) {
        Logger.log('ファイルの移動中にエラー: ' + moveError.toString());
      }
    }

    if (movedCount > 0) {
      var message = '処理中フォルダから' + movedCount + '件のファイルを未処理フォルダに戻しました。ファイル: ' + fileList.join(', ');
      Logger.log(message);
      return message;
    } else {
      return '処理中フォルダに残っているファイルはありませんでした。';
    }
  } catch (error) {
    var errorMessage = '中断ファイルのリカバリー処理中にエラー: ' + error.toString();
    Logger.log(errorMessage);
    return errorMessage;
  }
}

/**
 * スケジュールされたバッチ処理の実行
 */
function processBatchOnSchedule() {
  var processingEnabled = EnvironmentConfig.get('PROCESSING_ENABLED', 'true');
  if (processingEnabled !== true && processingEnabled !== 'true') {
    Logger.log('バッチ処理が無効化されています。スクリプトプロパティ PROCESSING_ENABLED が true でない。');
    return 'バッチ処理が無効化されています';
  }

  // 処理を実行
  return processBatch();
}

/**
 * バッチ処理を停止する
 */
function stopDailyProcess() {
  PropertiesService.getScriptProperties().setProperty('PROCESSING_ENABLED', 'false');
  return 'バッチ処理を停止しました';
}

/**
 * バッチ処理を開始する
 */
function startDailyProcess() {
  PropertiesService.getScriptProperties().setProperty('PROCESSING_ENABLED', 'true');
  return 'バッチ処理を開始しました';
}

/**
 * 日次サマリーをメール送信
 */
function sendDailySummary() {
  try {
    var spreadsheetId = EnvironmentConfig.get('LOG_SHEET_ID', '');
    if (!spreadsheetId) {
      throw new Error('LOG_SHEET_ID が設定されていません');
    }

    // スプレッドシートを開く
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var logSheet = spreadsheet.getSheetByName('processing_log');

    if (!logSheet) {
      return 'Processing logシートが見つかりません';
    }

    // 現在の日付を取得
    var today = new Date();
    var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

    // スプレッドシートの表示形式の日付（YYYY/MM/DD）
    var todayYMD = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

    // ヘッダー行とカラムインデックスを特定
    var allData = logSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      return 'Processing logシートにデータがありません';
    }

    var headerRow = allData[0];

    var fileNameIdx = headerRow.indexOf('file_name');
    var statusIdx = headerRow.indexOf('status');
    var processStartIdx = headerRow.indexOf('process_start');
    var processEndIdx = headerRow.indexOf('process_end');

    // カラム名が異なる場合の対応（日本語/英語両方チェック）
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
