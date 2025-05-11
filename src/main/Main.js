/**
 * メインコントローラー（OpenAI API対応版）
 * バッチ処理のエントリーポイント
 *
 * 依存モジュール:
 * - EnvironmentConfig (src/config/EnvironmentConfig.js)
 * - TriggerManager (src/main/TriggerManager.js)
 * - ZoomphoneProcessor (src/zoom/ZoomphoneProcessor.js)
 * - NotificationService (src/core/NotificationService.js)
 * - ZoomphoneTriggersWrapper (src/main/ZoomphoneTriggersWrapper.js) - 下位互換性のため
 */

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
      PROCESSED_SHEET_ID: config.PROCESSED_SHEET_ID || ''
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
    PROCESSED_SHEET_ID: ''
  };
}

/**
 * 設定をロードする関数（互換性のために残す - 新しいgetSystemSettings関数を使用）
 */
function loadSettings() {
  return getSystemSettings();
}

/**
 * ファイルからメタデータを抽出する関数
 * @param {string} fileName - ファイル名
 * @return {Object} 抽出されたメタデータ
 */
function extractMetadataFromFile(fileName) {
  var result = {
    recordId: '',
    callDate: null,
    callTime: null,
    salesPhoneNumber: '',
    customerPhoneNumber: '',
    metadataFound: false
  };

  try {
    // ファイル名からファイルを検索
    var files = DriveApp.getFilesByName(fileName);

    if (files.hasNext()) {
      var file = files.next();
      var fileDescription = file.getDescription();

      // ファイルの説明からJSONメタデータを解析
      if (fileDescription && fileDescription.trim()) {
        try {
          var metadata = JSON.parse(fileDescription);
          Logger.log('ファイルからメタデータを読み込みました: ' + JSON.stringify(metadata));

          // メタデータから録音IDを取得
          if (metadata.recording_id) {
            result.recordId = metadata.recording_id;
            result.metadataFound = true;
          }

          // 開始時間から日付と時間を抽出
          if (metadata.start_time) {
            var startTime = new Date(metadata.start_time);
            // タイムゾーン考慮してフォーマット
            result.callDate = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
            result.callTime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm:ss');
            result.metadataFound = true;
          }

          // 電話番号情報を取得
          // 通話方向の確認
          var direction = metadata.call_direction || 'unknown';

          // 発信者番号（セールスの電話番号）
          if (direction === 'outbound') {
            // 発信コール：発信者は営業
            result.salesPhoneNumber = metadata.caller_number || '';
            result.customerPhoneNumber = metadata.called_number || '';
          } else if (direction === 'inbound') {
            // 着信コール：着信者は営業
            result.customerPhoneNumber = metadata.caller_number || '';
            result.salesPhoneNumber = metadata.called_number || '';
          } else {
            // 通話方向不明の場合は両方の番号を取得してみる
            if (metadata.caller_number) {
              result.salesPhoneNumber = metadata.caller_number;
            }

            if (metadata.called_number) {
              result.customerPhoneNumber = metadata.called_number;
            }

            // Recordingsシートから電話番号情報を補完
            if (result.recordId && (!result.customerPhoneNumber || !result.salesPhoneNumber)) {
              var phoneNumbers = getPhoneNumbersFromRecordingsSheet(result.recordId);
              if (phoneNumbers) {
                if (!result.salesPhoneNumber && phoneNumbers.salesPhoneNumber) {
                  result.salesPhoneNumber = phoneNumbers.salesPhoneNumber;
                }
                if (!result.customerPhoneNumber && phoneNumbers.customerPhoneNumber) {
                  result.customerPhoneNumber = phoneNumbers.customerPhoneNumber;
                }
              }
            }
          }

          Logger.log('顧客電話番号変換前の値: rawCalledNumber=' + result.customerPhoneNumber);

          // 国際電話形式（+81）から日本の一般的な形式（0XXX）に変換
          if (result.customerPhoneNumber && result.customerPhoneNumber.indexOf('+81') === 0) {
            result.customerPhoneNumber = '0' + result.customerPhoneNumber.substring(3);
            Logger.log('顧客電話番号変換後: +81形式から変換 → ' + result.customerPhoneNumber);
          } else {
            Logger.log('顧客電話番号変換後: そのまま使用 → ' + result.customerPhoneNumber);
          }

          Logger.log('ファイルメタデータから抽出: record_id=' + result.recordId
            + ', call_date=' + result.callDate
            + ', call_time=' + result.callTime
            + ', sales_phone=' + result.salesPhoneNumber
            + ', customer_phone=' + result.customerPhoneNumber);
        } catch (jsonError) {
          Logger.log('メタデータのJSONパース失敗: ' + jsonError.toString());
        }
      }
    }
  } catch (fileError) {
    Logger.log('ファイルメタデータ取得エラー: ' + fileError.toString());
  }

  return result;
}

/**
 * Recordingsシートから録音IDに一致する電話番号情報を取得する
 * @param {string} recordId - 録音ID
 * @return {Object|null} 電話番号情報 {salesPhoneNumber, customerPhoneNumber} または null
 */
function getPhoneNumbersFromRecordingsSheet(recordId) {
  try {
    var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
    if (!spreadsheetId) return null;

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Recordings');

    if (!sheet) {
      Logger.log('Recordingsシートが見つかりません');
      return null;
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // ヘッダー行をスキップして2行目から処理
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      // record_idカラム（1列目）をチェック
      if (row[0] === recordId) {
        return {
          salesPhoneNumber: row[6] || '', // sales_phone_number（7列目）
          customerPhoneNumber: row[7] || '' // customer_phone_number（8列目）
        };
      }
    }

    Logger.log('録音ID ' + recordId + ' に一致する行がRecordingsシートに見つかりません');
    return null;
  } catch (e) {
    Logger.log('Recordingsシートからの電話番号取得でエラー: ' + e.toString());
    return null;
  }
}

/**
 * 通話記録をスプレッドシートに保存する関数
 * @param {Object} callData - 通話データオブジェクト
 * @param {string} [targetSpreadsheetId] - 保存先のスプレッドシートID（省略時はPROCESSED_SHEET_ID）
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
      sheet.getRange(1, 1, 1, 17).setValues([
        ['record_id', 'call_date', 'call_time', 'sales_phone_number', 'sales_company', 'sales_person', 'customer_phone_number', 'customer_company', 'customer_name', 'call_status1', 'call_status2', 'reason_for_refusal', 'reason_for_refusal_category', 'reason_for_appointment', 'reason_for_appointment_category', 'summary', 'full_transcript']
      ]);
    }

    // 最終行の次の行に追加
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    Logger.log('挿入行: ' + newRow);

    // データを配列として準備（新しい順序に合わせる）
    var rowData = [
      callData.recordId,
      callData.callDate,
      callData.callTime,
      callData.salesPhoneNumber || '',
      callData.salesCompany || '',
      callData.salesPerson || '',
      callData.customerPhoneNumber || '',
      callData.customerCompany || '',
      callData.customerName || '',
      callData.callStatus1 || '',
      callData.callStatus2 || '',
      callData.reasonForRefusal || '',
      callData.reasonForRefusalCategory || '',
      callData.reasonForAppointment || '',
      callData.reasonForAppointmentCategory || '',
      callData.summary || '',
      callData.transcription || ''
    ];

    Logger.log('挿入データ準備完了: record_id=' + callData.recordId + ', call_date=' + callData.callDate + ', call_time=' + callData.callTime);

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
        var transcriptionResult;
        try {
          transcriptionResult = TranscriptionService.transcribe(
            file,
            localSettings.ASSEMBLYAI_API_KEY,
            localSettings.OPENAI_API_KEY
          );
        } catch (transcriptionError) {
          Logger.log('文字起こし処理でエラーが発生: ' + transcriptionError.toString());

          // 最低限の結果を生成して処理を続行
          transcriptionResult = {
            text: '【文字起こし失敗: ' + transcriptionError.toString() + '】',
            rawText: '',
            utterances: [],
            speakerInfo: {},
            speakerRoles: {},
            fileName: file.getName()
          };
        }

        // 文字起こし結果のチェック
        Logger.log('文字起こし結果: 文字数=' + (transcriptionResult.text ? transcriptionResult.text.length : 0));
        Logger.log('発話数: ' + (transcriptionResult.utterances ? transcriptionResult.utterances.length : 0));

        // OpenAI GPT-4o-miniによる情報抽出
        var extractedInfo;

        if (!localSettings.OPENAI_API_KEY) {
          throw new Error('OpenAI APIキーが設定されていないため、処理を続行できません');
        }

        try {
          extractedInfo = InformationExtractor.extract(
            transcriptionResult.text,
            transcriptionResult.utterances,
            localSettings.OPENAI_API_KEY
          );
        } catch (extractionError) {
          Logger.log('情報抽出処理でエラーが発生: ' + extractionError.toString());

          // 最低限の情報を生成して処理を続行
          extractedInfo = {
            sales_company: '不明（抽出エラー）',
            sales_person: '不明（抽出エラー）',
            customer_company: '不明（抽出エラー）',
            customer_name: '不明（抽出エラー）',
            call_status1: '',
            call_status2: '',
            reason_for_refusal: '',
            reason_for_appointment: '',
            summary: '情報抽出に失敗しました: ' + extractionError.toString()
          };
        }

        // 抽出情報のチェック
        Logger.log('情報抽出結果: ' + JSON.stringify({
          sales_company: extractedInfo.sales_company,
          sales_person: extractedInfo.sales_person,
          customer_company: extractedInfo.customer_company,
          call_status1: extractedInfo.call_status1,
          call_status2: extractedInfo.call_status2,
          summaryLength: extractedInfo.summary ? extractedInfo.summary.length : 0
        }));

        // 処理終了時間を取得
        var fileEndTime = new Date();
        var fileEndTimeStr = Utilities.formatDate(fileEndTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

        // ファイルからメタデータを抽出
        var metadata = extractMetadataFromFile(file.getName());

        // 必要なメタデータが取得できなかった場合はエラー扱い
        if (!metadata.recordId) {
          throw new Error('メタデータから録音IDを取得できませんでした: ' + file.getName());
        }

        // Recordingsシートの文字起こし状態を更新（成功）
        // 録音IDに一致する行を検索し、文字起こし状態を更新
        updateTranscriptionStatusByRecordId(metadata.recordId, 'SUCCESS', fileStartTimeStr, fileEndTimeStr);

        // スプレッドシートに書き込み
        var callData = {
          fileName: file.getName(),
          recordId: metadata.recordId,
          callDate: metadata.callDate,
          callTime: metadata.callTime,
          salesPhoneNumber: metadata.salesPhoneNumber,
          customerPhoneNumber: metadata.customerPhoneNumber,
          salesCompany: extractedInfo.sales_company,
          salesPerson: extractedInfo.sales_person,
          customerCompany: extractedInfo.customer_company,
          customerName: extractedInfo.customer_name,
          callStatus1: extractedInfo.call_status1,
          callStatus2: extractedInfo.call_status2,
          reasonForRefusal: extractedInfo.reason_for_refusal,
          reasonForRefusalCategory: extractedInfo.reason_for_refusal_category,
          reasonForAppointment: extractedInfo.reason_for_appointment,
          reasonForAppointmentCategory: extractedInfo.reason_for_appointment_category,
          summary: extractedInfo.summary,
          transcription: transcriptionResult.text
        };

        // call_recordsシートに出力
        var processedSheetId = localSettings.PROCESSED_SHEET_ID || '';
        Logger.log('文字起こし結果の保存先: PROCESSED_SHEET_ID=' + processedSheetId);
        saveCallRecordToSheet(callData, processedSheetId, 'call_records');

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

        // ファイルからメタデータを取得してみる（可能であれば）
        try {
          var errorMetadata = extractMetadataFromFile(file.getName());
          if (errorMetadata && errorMetadata.recordId) {
            // Recordingsシートの文字起こし状態を更新（エラー）
            updateTranscriptionStatusByRecordId(
              errorMetadata.recordId,
              'ERROR: ' + error.toString().substring(0, 100), // エラーメッセージは短く切り詰める
              fileStartTimeStr,
              fileEndTimeStr
            );
          }
        } catch (metaError) {
          Logger.log('エラー処理中のメタデータ取得に失敗: ' + metaError.toString());
        }

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
 * 処理中フォルダにある未完了ファイルを未処理フォルダに戻す
 * AppScriptの制限により中断されたファイルを復旧する
 */
function recoverInterruptedFiles() {
  var startTime = new Date();
  Logger.log('中断ファイル復旧処理開始: ' + startTime);

  try {
    // 設定を取得
    var localSettings = getSystemSettings();

    if (!localSettings.PROCESSING_FOLDER_ID) {
      throw new Error('処理中フォルダIDが設定されていません');
    }

    if (!localSettings.SOURCE_FOLDER_ID) {
      throw new Error('処理対象フォルダIDが設定されていません');
    }

    // 処理中フォルダのファイルを取得
    var processingFolder = DriveApp.getFolderById(localSettings.PROCESSING_FOLDER_ID);
    var files = processingFolder.getFiles();
    var interruptedFiles = [];

    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType() || "";

      // 音声ファイルのみを対象とする
      if (mimeType.indexOf('audio/') === 0 ||
        mimeType === 'application/octet-stream' ||
        file.getName().toLowerCase().indexOf('.mp3') !== -1) {
        interruptedFiles.push(file);
      }
    }

    Logger.log('中断された可能性のあるファイル数: ' + interruptedFiles.length);

    if (interruptedFiles.length === 0) {
      return '中断されたファイルはありませんでした。';
    }

    // 復旧結果のトラッキング
    var results = {
      total: interruptedFiles.length,
      recovered: 0,
      failed: 0,
      details: []
    };

    // ファイルを復旧
    for (var i = 0; i < interruptedFiles.length; i++) {
      var file = interruptedFiles[i];

      try {
        Logger.log('ファイル復旧処理: ' + file.getName());

        // ファイルからメタデータを取得
        var metadata = extractMetadataFromFile(file.getName());

        if (metadata && metadata.recordId) {
          // Recordingsシートの文字起こし状態を更新 (INTERRUPTED)
          updateTranscriptionStatusByRecordId(
            metadata.recordId,
            'INTERRUPTED',
            '',
            Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
          );
        }

        // ファイルを未処理フォルダに移動
        FileProcessor.moveFileToFolder(file, localSettings.SOURCE_FOLDER_ID);

        // 結果を記録
        results.recovered++;
        results.details.push({
          fileName: file.getName(),
          status: 'recovered',
          message: '未処理フォルダに復旧しました'
        });
      } catch (error) {
        Logger.log('ファイル復旧エラー: ' + error.toString());

        results.failed++;
        results.details.push({
          fileName: file.getName(),
          status: 'error',
          message: error.toString()
        });
      }
    }

    // 処理結果のログ出力
    var endTime = new Date();
    var processingTime = (endTime - startTime) / 1000; // 秒単位

    var summary = '中断ファイル復旧処理完了: ' +
      '対象=' + results.total + '件, ' +
      '復旧=' + results.recovered + '件, ' +
      '失敗=' + results.failed + '件, ' +
      '処理時間=' + processingTime + '秒';

    Logger.log(summary);
    return summary;

  } catch (error) {
    var errorMessage = '中断ファイル復旧処理でエラーが発生: ' + error.toString();
    Logger.log(errorMessage);
    return errorMessage;
  }
}

/**
 * 録音IDに一致するRecordingsシートの行を検索し、文字起こし状態を更新する
 * @param {string} recordId - 録音ID
 * @param {string} status - 更新するステータス
 * @param {string} processStart - 処理開始時間
 * @param {string} processEnd - 処理終了時間
 */
function updateTranscriptionStatusByRecordId(recordId, status, processStart, processEnd) {
  try {
    var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
    if (!spreadsheetId) return;

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Recordings');

    if (!sheet) {
      Logger.log('Recordingsシートが見つかりません');
      return;
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // ヘッダー行をスキップして2行目から処理
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      // record_idカラム（1列目）をチェック
      if (row[0] === recordId) {
        var rowIndex = i + 1; // スプレッドシートの行番号（1始まり）

        // ZoomphoneProcessor.updateRecordingStatus関数を使用
        ZoomphoneProcessor.updateRecordingStatus(rowIndex, status, 'transcription');

        // process_start（13列目）とprocess_end（14列目）を更新
        if (processStart) sheet.getRange(rowIndex, 13).setValue(processStart);
        if (processEnd) sheet.getRange(rowIndex, 14).setValue(processEnd);

        Logger.log('文字起こし状態を更新: record_id=' + recordId + ', status=' + status);
        return;
      }
    }

    Logger.log('録音ID ' + recordId + ' に一致する行がRecordingsシートに見つかりません');
  } catch (e) {
    Logger.log('文字起こし状態の更新でエラー: ' + e.toString());
  }
}

/**
 * スケジュールされたバッチ処理の実行
 */
function processBatchOnSchedule() {
  var now = new Date();
  var hour = now.getHours();

  // 6:00～24:00の間のみ実行
  if (hour >= 6 && hour < 24) {
    var processingEnabled = EnvironmentConfig.get('PROCESSING_ENABLED', 'true');
    if (processingEnabled !== true && processingEnabled !== 'true') {
      Logger.log('バッチ処理が無効化されています。スクリプトプロパティ PROCESSING_ENABLED が true でない。');
      return 'バッチ処理が無効化されています';
    }

    // 処理を実行
    return processBatch();
  } else {
    // 営業時間外は何もしない
    return "営業時間外（6:00～24:00以外）のため処理をスキップしました";
  }
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
 * メインエントリーポイント
 * このスクリプトのデフォルト実行関数
 */
function main() {
  try {
    Logger.log('メイン処理スタート: ' + new Date().toString());

    // 1. Zoom録音ファイルの取得処理
    Logger.log('Recordingsシートからの録音ファイル取得処理を開始します（タイムスタンプが古い順に処理）...');
    var recordingsResult = ZoomphoneProcessor.processRecordingsFromSheet();
    Logger.log('Recordingsシートからの録音取得結果: ' + JSON.stringify(recordingsResult));

    // 2. メイン処理バッチ（文字起こし、情報抽出）
    processBatch();

    Logger.log('メイン処理完了: ' + new Date().toString());
    return '処理完了';
  } catch (error) {
    // エラーハンドリング
    Logger.log('エラー発生: ' + error.toString());
    return 'エラー発生: ' + error.toString();
  }
}

/**
 * 日次処理結果サマリーを手動で送信する関数
 * 任意のタイミングで実行可能
 * @param {string} [dateStr] - 日付文字列（YYYY/MM/DD形式）。省略時は本日の日付
 * @return {string} 実行結果メッセージ
 */
function manualSendDailySummary(dateStr) {
  try {
    // 日付が指定されていない場合は本日の日付を使用
    if (!dateStr) {
      var today = new Date();
      dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
    }

    // 日付形式の確認（YYYY/MM/DD形式かチェック）
    var datePattern = /^\d{4}\/\d{2}\/\d{2}$/;
    if (!datePattern.test(dateStr)) {
      throw new Error('日付形式が正しくありません。YYYY/MM/DD形式で指定してください。');
    }

    Logger.log('指定された日付のサマリーを送信します: ' + dateStr);

    // Recordingsシートから指定日付のデータを集計（関数を直接記述）
    var summary = calculateSummaryFromSheet(dateStr);

    // 管理者メールアドレスを取得
    var settings = getSystemSettings();
    if (settings && settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
      // 各管理者にサマリーメールを送信
      for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
        NotificationService.sendDailyProcessingSummary(
          settings.ADMIN_EMAILS[i],
          summary,
          dateStr
        );
      }
      return dateStr + ' の日次処理サマリーを ' + settings.ADMIN_EMAILS.length + '名の管理者に送信しました。';
    } else {
      throw new Error('送信先のメールアドレスが設定されていません。');
    }
  } catch (error) {
    Logger.log('手動サマリー送信中にエラー: ' + error.toString());
    return 'エラーが発生しました: ' + error.toString();
  }
}

/**
 * 日次処理結果サマリーを自動で送信する関数（19:00トリガー用）
 * トリガーによって毎日19:00に自動実行される
 * @return {string} 実行結果メッセージ
 */
function sendDailySummary() {
  try {
    // 今日の日付を取得
    var today = new Date();
    var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

    Logger.log('本日(' + dateStr + ')の処理サマリーを自動送信します');

    // Recordingsシートから本日のデータを集計
    var summary = calculateSummaryFromSheet(dateStr);

    // 管理者メールアドレスを取得
    var settings = getSystemSettings();
    if (settings && settings.ADMIN_EMAILS && settings.ADMIN_EMAILS.length > 0) {
      // 各管理者にサマリーメールを送信
      for (var i = 0; i < settings.ADMIN_EMAILS.length; i++) {
        NotificationService.sendDailyProcessingSummary(
          settings.ADMIN_EMAILS[i],
          summary,
          dateStr
        );
      }
      return dateStr + ' の日次処理サマリーを ' + settings.ADMIN_EMAILS.length + '名の管理者に自動送信しました。';
    } else {
      throw new Error('送信先のメールアドレスが設定されていません。');
    }
  } catch (error) {
    Logger.log('自動サマリー送信中にエラー: ' + error.toString());
    return 'エラーが発生しました: ' + error.toString();
  }
}

/**
 * Recordingsシートから指定日付の処理結果を集計する
 * manualSendDailySummary関数用の内部関数
 * @param {string} dateStr - 日付文字列（YYYY/MM/DD形式）
 * @return {Object} 集計結果 { success: number, error: number, details: Array }
 */
function calculateSummaryFromSheet(dateStr) {
  try {
    var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
    if (!spreadsheetId) {
      return { success: 0, error: 0, details: [] };
    }

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Recordings');

    if (!sheet) {
      Logger.log('Recordingsシートが見つかりません');
      return { success: 0, error: 0, details: [] };
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // 集計用変数
    var summary = {
      success: 0,
      error: 0,
      details: []
    };

    // 日付のフォーマットを変換（YYYY/MM/DD → YYYY-MM-DD）
    // シート内の日付がハイフン形式になっているため
    var formattedDateStr = dateStr.replace(/\//g, '-');
    Logger.log('検索用に変換した日付: ' + formattedDateStr);

    // ヘッダー行をスキップして2行目から処理
    for (var i = 1; i < values.length; i++) {
      var row = values[i];

      // call_date列（4列目）をチェック
      var callDate = row[3]; // call_date

      // 日付の確認をログ出力
      Logger.log('レコード #' + (i + 1) + ' の日付: ' + callDate + ' (型: ' + typeof callDate + ')');

      // 日付文字列に変換
      var rowDateStr = '';
      if (callDate instanceof Date) {
        rowDateStr = Utilities.formatDate(callDate, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else if (typeof callDate === 'string') {
        // 既に文字列の場合はそのまま使用
        rowDateStr = callDate;
      }

      Logger.log('比較する日付: シート=' + rowDateStr + ', 指定=' + formattedDateStr);

      // 指定日付のレコードのみ集計
      if (rowDateStr === formattedDateStr) {
        Logger.log('日付一致！レコード #' + (i + 1) + ' を処理します');

        // ステータスに基づいて成功/エラーをカウント
        var status = row[9]; // status_fetch
        var transcStatus = row[11]; // status_transcription

        Logger.log('ステータス: fetch=' + status + ', transcription=' + transcStatus);

        // PROCESSEDは常に成功としてカウント
        if (status === 'PROCESSED') {
          summary.success++;
          Logger.log('成功としてカウント');
        } else if (status === 'ERROR' || (status && status.indexOf('ERROR') === 0)) {
          summary.error++;
          Logger.log('エラーとしてカウント');
        }
      }
    }

    Logger.log('集計結果: 成功=' + summary.success + '件, エラー=' + summary.error + '件');
    return summary;
  } catch (e) {
    Logger.log('サマリー集計でエラー: ' + e.toString());
    return { success: 0, error: 0, details: [] };
  }
}

/**
 * スプレッドシートが開かれたときに実行される関数
 * カスタムメニューを追加する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('文字起こしシステム')
    .addItem('今すぐバッチ処理実行', 'processBatch')
    .addItem('直近1時間の録音を取得', 'TriggerManager.fetchLastHourRecordings')
    .addItem('直近2時間の録音を取得', 'TriggerManager.fetchLast2HoursRecordings')
    .addItem('直近6時間の録音を取得', 'TriggerManager.fetchLast6HoursRecordings')
    .addItem('直近24時間の録音を取得', 'TriggerManager.fetchLast24HoursRecordings')
    .addItem('直近48時間の録音を取得', 'TriggerManager.fetchLast48HoursRecordings')
    .addItem('すべての未処理録音を取得', 'TriggerManager.fetchAllPendingRecordings')
    .addToUi();
}
