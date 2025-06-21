/**
 * メインコントローラー（OpenAI API対応版）
 * バッチ処理のエントリーポイント
 *
 * 依存モジュール:
 * - EnvironmentConfig (src/config/EnvironmentConfig.js)
 * - ConfigManager (src/core/ConfigManager.js)
 * - Constants (src/core/Constants.js)
 * - FileMovementService (src/core/FileMovementService.js)
 * - ErrorHandler (src/core/ErrorHandler.js)
 * - RecoveryService (src/core/RecoveryService.js)
 * - PartialFailureDetector (src/core/PartialFailureDetector.js)
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
var settings = ConfigManager.getConfig();

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
          // [RETRIED]や[RESET_RETRY]などのマークがある場合、JSON部分のみを抽出
          var jsonPart = fileDescription;
          var markerIndex = fileDescription.indexOf("[");

          if (markerIndex > 0) {
            // マーカーの前のJSONパートのみを使用
            jsonPart = fileDescription.substring(0, markerIndex).trim();
            Logger.log('マーカーを除去したJSONパート: ' + jsonPart);
          } else if (markerIndex === 0) {
            // 先頭からマーカーが始まる場合はJSONが見つからない
            throw new Error("JSON部分が見つかりません: " + fileDescription);
          }

          var metadata = JSON.parse(jsonPart);
          Logger.log('ファイルからメタデータを読み込みました: ' + JSON.stringify(metadata));

          // メタデータから録音IDを取得
          if (metadata.recording_id) {
            result.recordId = metadata.recording_id;
            result.metadataFound = true;

            // ここで重要: record_idが取得できたら、まずRecordingsシートからメタデータを取得
            var sheetMetadata = getMetadataFromRecordingsSheet(result.recordId);
            if (sheetMetadata) {
              Logger.log('Recordingsシートからメタデータを取得: ' + JSON.stringify(sheetMetadata));

              // Recordingsシートから取得した情報を優先的に使用
              if (sheetMetadata.callDate) {
                result.callDate = sheetMetadata.callDate;
              }

              if (sheetMetadata.callTime) {
                result.callTime = sheetMetadata.callTime;
              }

              if (sheetMetadata.salesPhoneNumber) {
                result.salesPhoneNumber = sheetMetadata.salesPhoneNumber;
              }

              if (sheetMetadata.customerPhoneNumber) {
                result.customerPhoneNumber = sheetMetadata.customerPhoneNumber;
              }

              // 必要な情報がすべて揃っている場合は、他のソースからの抽出をスキップ
              if (result.callDate && result.callTime &&
                result.salesPhoneNumber && result.customerPhoneNumber) {
                Logger.log('Recordingsシートから必要な情報をすべて取得しました。JSONメタデータからの抽出をスキップします。');
                return result;
              }

              // 一部の情報だけが取得できた場合は、残りをJSONメタデータから補完
              Logger.log('Recordingsシートから一部の情報を取得しました。残りの情報をJSONメタデータから補完します。');
            }
          }

          // Recordingsシートから取得できなかった情報をJSONメタデータから補完
          // 開始時間から日付と時間を抽出（Recordingsシートの情報がない場合のみ）
          if ((!result.callDate || !result.callTime) && metadata.start_time) {
            var startTime = new Date(metadata.start_time);
            // タイムゾーン考慮してフォーマット
            if (!result.callDate) {
              result.callDate = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
            }
            if (!result.callTime) {
              result.callTime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm:ss');
            }
            result.metadataFound = true;
          }

          // 電話番号情報を取得（Recordingsシートの情報がない場合のみ）
          if (!result.salesPhoneNumber || !result.customerPhoneNumber) {
            // 通話方向の確認
            var direction = metadata.call_direction || 'unknown';

            // 発信者番号（セールスの電話番号）
            if (direction === 'outbound') {
              // 発信コール：発信者は営業
              if (!result.salesPhoneNumber) result.salesPhoneNumber = metadata.caller_number || '';
              if (!result.customerPhoneNumber) result.customerPhoneNumber = metadata.called_number || '';
            } else if (direction === 'inbound') {
              // 着信コール：着信者は営業
              if (!result.customerPhoneNumber) result.customerPhoneNumber = metadata.caller_number || '';
              if (!result.salesPhoneNumber) result.salesPhoneNumber = metadata.called_number || '';
            } else {
              // 通話方向不明の場合は両方の番号を取得してみる
              if (!result.salesPhoneNumber && metadata.caller_number) {
                result.salesPhoneNumber = metadata.caller_number;
              }

              if (!result.customerPhoneNumber && metadata.called_number) {
                result.customerPhoneNumber = metadata.called_number;
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

          Logger.log('メタデータ抽出結果: record_id=' + result.recordId
            + ', call_date=' + result.callDate
            + ', call_time=' + result.callTime
            + ', sales_phone=' + result.salesPhoneNumber
            + ', customer_phone=' + result.customerPhoneNumber);
        } catch (jsonError) {
          Logger.log('メタデータのJSONパース失敗: ' + jsonError.toString());

          // JSONパース失敗時はファイル名からレコードIDの抽出を試みる
          var fileNameMatch = fileName.match(/zoom_call_(\d+)_([a-f0-9]+)\.mp3/i);
          if (fileNameMatch) {
            // タイムスタンプとレコードIDを抽出
            var timestamp = fileNameMatch[1];
            var recordId = fileNameMatch[2];

            if (recordId) {
              result.recordId = recordId;
              Logger.log('ファイル名からrecord_idを抽出: ' + recordId);

              // Recordingsシートからメタデータを取得
              var sheetMetadata = getMetadataFromRecordingsSheet(recordId);
              if (sheetMetadata) {
                Logger.log('Recordingsシートからメタデータを取得: ' + JSON.stringify(sheetMetadata));

                // Recordingsシートから取得した情報を優先的に使用
                if (sheetMetadata.callDate) result.callDate = sheetMetadata.callDate;
                if (sheetMetadata.callTime) result.callTime = sheetMetadata.callTime;
                if (sheetMetadata.salesPhoneNumber) result.salesPhoneNumber = sheetMetadata.salesPhoneNumber;
                if (sheetMetadata.customerPhoneNumber) result.customerPhoneNumber = sheetMetadata.customerPhoneNumber;

                // もし日付と時間が取得できた場合は、タイムスタンプからの抽出をスキップ
                if (result.callDate && result.callTime) {
                  Logger.log('Recordingsシートから日付と時間を取得しました。タイムスタンプからの抽出をスキップします。');
                  return result;
                }
              }

              // 日付と時間がRecordingsシートから取得できなかった場合のみタイムスタンプから抽出
              if ((!result.callDate || !result.callTime) && timestamp && timestamp.length >= 14) {
                // YYYYMMDDhhmmss 形式を解析
                var year = timestamp.substring(0, 4);
                var month = timestamp.substring(4, 6);
                var day = timestamp.substring(6, 8);
                var hour = timestamp.substring(8, 10);
                var minute = timestamp.substring(10, 12);
                var second = timestamp.substring(12, 14);

                var dateStr = year + '-' + month + '-' + day;
                var timeStr = hour + ':' + minute + ':' + second;

                if (!result.callDate) result.callDate = dateStr;
                if (!result.callTime) result.callTime = timeStr;

                Logger.log('ファイル名からタイムスタンプを抽出: ' + dateStr + ' ' + timeStr);
              }
            }
          }
        }
      } else {
        // description が空の場合もファイル名からの抽出を試みる
        var fileNameMatch = fileName.match(/zoom_call_(\d+)_([a-f0-9]+)\.mp3/i);
        if (fileNameMatch && fileNameMatch[2]) {
          result.recordId = fileNameMatch[2];
          Logger.log('ファイル名から直接record_idを抽出: ' + result.recordId);

          // Recordingsシートからメタデータを取得
          var sheetMetadata = getMetadataFromRecordingsSheet(result.recordId);
          if (sheetMetadata) {
            Logger.log('Recordingsシートからメタデータを取得: ' + JSON.stringify(sheetMetadata));

            // Recordingsシートから取得した情報を設定
            if (sheetMetadata.callDate) result.callDate = sheetMetadata.callDate;
            if (sheetMetadata.callTime) result.callTime = sheetMetadata.callTime;
            if (sheetMetadata.salesPhoneNumber) result.salesPhoneNumber = sheetMetadata.salesPhoneNumber;
            if (sheetMetadata.customerPhoneNumber) result.customerPhoneNumber = sheetMetadata.customerPhoneNumber;

            // タイムスタンプからの抽出が必要かチェック
            if (!result.callDate || !result.callTime) {
              var timestamp = fileNameMatch[1];
              if (timestamp && timestamp.length >= 14) {
                var year = timestamp.substring(0, 4);
                var month = timestamp.substring(4, 6);
                var day = timestamp.substring(6, 8);
                var hour = timestamp.substring(8, 10);
                var minute = timestamp.substring(10, 12);
                var second = timestamp.substring(12, 14);

                if (!result.callDate) result.callDate = year + '-' + month + '-' + day;
                if (!result.callTime) result.callTime = hour + ':' + minute + ':' + second;

                Logger.log('ファイル名からタイムスタンプを補完: ' + result.callDate + ' ' + result.callTime);
              }
            }
          } else {
            // Recordingsシートからデータが取得できない場合はファイル名のタイムスタンプから抽出
            var timestamp = fileNameMatch[1];
            if (timestamp && timestamp.length >= 14) {
              var year = timestamp.substring(0, 4);
              var month = timestamp.substring(4, 6);
              var day = timestamp.substring(6, 8);
              var hour = timestamp.substring(8, 10);
              var minute = timestamp.substring(10, 12);
              var second = timestamp.substring(12, 14);

              result.callDate = year + '-' + month + '-' + day;
              result.callTime = hour + ':' + minute + ':' + second;

              Logger.log('ファイル名から日付と時間を抽出: ' + result.callDate + ' ' + result.callTime);
            }
          }
        }
      }
    }
  } catch (fileError) {
    Logger.log('ファイルメタデータ取得エラー: ' + fileError.toString());

    // 最終手段としてファイル名からの抽出を試みる
    try {
      var fileNameMatch = fileName.match(/zoom_call_(\d+)_([a-f0-9]+)\.mp3/i);
      if (fileNameMatch && fileNameMatch[2]) {
        result.recordId = fileNameMatch[2];
        Logger.log('エラー後のファイル名からの直接抽出: record_id=' + result.recordId);

        // Recordingsシートからメタデータを取得試行
        var sheetMetadata = getMetadataFromRecordingsSheet(result.recordId);
        if (sheetMetadata) {
          Logger.log('エラー後にRecordingsシートからメタデータを取得: ' + JSON.stringify(sheetMetadata));

          // Recordingsシートから取得した情報を設定
          if (sheetMetadata.callDate) result.callDate = sheetMetadata.callDate;
          if (sheetMetadata.callTime) result.callTime = sheetMetadata.callTime;
          if (sheetMetadata.salesPhoneNumber) result.salesPhoneNumber = sheetMetadata.salesPhoneNumber;
          if (sheetMetadata.customerPhoneNumber) result.customerPhoneNumber = sheetMetadata.customerPhoneNumber;
        }
      }
    } catch (e) {
      Logger.log('ファイル名からのID抽出も失敗: ' + e.toString());
    }
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
 * Recordingsシートから録音IDに一致するメタデータを取得する
 * 電話番号に加えて、日付と時間も取得
 * @param {string} recordId - 録音ID
 * @return {Object|null} メタデータ {callDate, callTime, salesPhoneNumber, customerPhoneNumber} または null
 */
function getMetadataFromRecordingsSheet(recordId) {
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
        // 日付データを適切なフォーマットに変換
        var callDate = '';
        if (row[3]) { // call_date（4列目）
          if (row[3] instanceof Date) {
            callDate = Utilities.formatDate(row[3], 'Asia/Tokyo', 'yyyy-MM-dd');
          } else if (typeof row[3] === 'string') {
            callDate = row[3];
          }
        }

        // 時間データを適切なフォーマットに変換
        var callTime = '';
        if (row[4]) { // call_time（5列目）
          if (row[4] instanceof Date) {
            callTime = Utilities.formatDate(row[4], 'Asia/Tokyo', 'HH:mm:ss');
          } else if (typeof row[4] === 'string') {
            callTime = row[4];
          }
        }

        // 結果を返す
        return {
          callDate: callDate,
          callTime: callTime,
          salesPhoneNumber: row[6] || '', // sales_phone_number（7列目）
          customerPhoneNumber: row[7] || '' // customer_phone_number（8列目）
        };
      }
    }

    Logger.log('録音ID ' + recordId + ' に一致する行がRecordingsシートに見つかりません');
    return null;
  } catch (e) {
    Logger.log('Recordingsシートからのメタデータ取得でエラー: ' + e.toString());
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
      // ヘッダー行を設定（sales_personとcustomer_companyを削除）
      sheet.getRange(1, 1, 1, 15).setValues([
        ['record_id', 'call_date', 'call_time', 'sales_phone_number', 'sales_company', 'customer_phone_number', 'customer_name', 'call_status1', 'call_status2', 'reason_for_refusal', 'reason_for_refusal_category', 'reason_for_appointment', 'reason_for_appointment_category', 'summary', 'full_transcript']
      ]);
    } else {
      // 既存シートの列数を確認して不足があれば拡張
      var currentColumns = sheet.getMaxColumns();
      var requiredColumns = 15;
      if (currentColumns < requiredColumns) {
        Logger.log('シートの列数が不足しています（現在: ' + currentColumns + ', 必要: ' + requiredColumns + '）。列を追加します。');
        sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
        Logger.log('列を追加しました。新しい列数: ' + sheet.getMaxColumns());
      }
    }

    // 最終行の次の行に追加
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    Logger.log('挿入行: ' + newRow);

    // データを配列として準備（新しい順序に合わせる - sales_personとcustomer_companyを削除）
    var rowData = [
      callData.recordId,
      callData.callDate,
      callData.callTime,
      callData.salesPhoneNumber || '',
      callData.salesCompany || '',
      callData.customerPhoneNumber || '',
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

    // 行を追加前の安全チェック
    var maxColumns = sheet.getMaxColumns();
    if (rowData.length > maxColumns) {
      Logger.log('データ列数がシート列数を超えています（データ: ' + rowData.length + ', シート: ' + maxColumns + '）。');
      // 不足分の列を追加
      sheet.insertColumnsAfter(maxColumns, rowData.length - maxColumns);
      Logger.log('列を追加して調整しました。新しい列数: ' + sheet.getMaxColumns());
    }

    // 行を追加
    try {
      sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
      Logger.log('行の追加に成功しました');
    } catch (writeError) {
      Logger.log('行の追加中にエラー: ' + writeError.toString());
      Logger.log('シート情報: 最大行=' + sheet.getMaxRows() + ', 最大列=' + sheet.getMaxColumns());
      Logger.log('書き込み座標: 行=' + newRow + ', 列=1, データ列数=' + rowData.length);
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
    if (!localSettings.OPENAI_API_KEY) {
      throw new Error('OpenAI APIキーが設定されていません');
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
        var hasTranscriptionError = false;
        var transcriptionErrorMessage = '';

        try {
          transcriptionResult = TranscriptionService.transcribe(
            file,
            localSettings.OPENAI_API_KEY
          );

          // 文字起こし結果にエラーフィールドがある場合をチェック
          if (transcriptionResult && transcriptionResult.error) {
            hasTranscriptionError = true;
            transcriptionErrorMessage = transcriptionResult.error;
            Logger.log('文字起こし結果にエラーが含まれています: ' + transcriptionErrorMessage);
          }

          // OpenAI APIエラーの検知（文字起こしテキストにエラーメッセージが含まれている場合）
          if (transcriptionResult && transcriptionResult.text &&
            (transcriptionResult.text.includes('【文字起こし失敗:') ||
              transcriptionResult.text.includes('GPT-4o-mini API呼び出しエラー') ||
              transcriptionResult.text.includes('OpenAI APIからのレスポンスエラー') ||
              transcriptionResult.text.includes('insufficient_quota'))) {
            hasTranscriptionError = true;
            transcriptionErrorMessage = 'OpenAI API処理でエラーが発生しました';
            Logger.log('文字起こしテキストにエラーメッセージが検出されました');
          }

        } catch (transcriptionError) {
          Logger.log('文字起こし処理でエラーが発生: ' + transcriptionError.toString());
          hasTranscriptionError = true;
          transcriptionErrorMessage = transcriptionError.toString();

          // 最低限の結果を生成
          transcriptionResult = {
            text: '【文字起こし失敗: ' + transcriptionError.toString() + '】',
            rawText: '',
            utterances: [],
            speakerInfo: {},
            speakerRoles: {},
            fileName: file.getName(),
            error: transcriptionError.toString()
          };
        }

        // 文字起こし結果のチェック
        Logger.log('文字起こし結果: 文字数=' + (transcriptionResult.text ? transcriptionResult.text.length : 0));
        Logger.log('発話数: ' + (transcriptionResult.utterances ? transcriptionResult.utterances.length : 0));
        Logger.log('エラー状態: ' + hasTranscriptionError);

        // エラーが発生している場合は早期にエラー処理へ
        if (hasTranscriptionError) {
          throw new Error('文字起こし処理でエラーが発生: ' + transcriptionErrorMessage);
        }

        // OpenAI GPT-4o-miniによる情報抽出
        var extractedInfo;
        var hasExtractionError = false;
        var extractionErrorMessage = '';

        if (!localSettings.OPENAI_API_KEY) {
          hasExtractionError = true;
          extractionErrorMessage = 'OpenAI APIキーが設定されていません';
          Logger.log('OpenAI APIキーが設定されていないため、情報抽出をスキップします');
        } else {
          try {
            extractedInfo = InformationExtractor.extract(
              transcriptionResult.text,
              transcriptionResult.utterances,
              localSettings.OPENAI_API_KEY
            );
            Logger.log('情報抽出完了');
          } catch (extractionError) {
            Logger.log('情報抽出でエラーが発生: ' + extractionError.toString());
            hasExtractionError = true;
            extractionErrorMessage = extractionError.toString();

            // OpenAI APIクォータエラーの特別処理
            if (extractionError.toString().includes('insufficient_quota') ||
              extractionError.toString().includes('429')) {
              extractionErrorMessage = 'OpenAI APIクォータ制限に達しました';
              Logger.log('OpenAI APIクォータエラーを検出しました');
            }
          }
        }

        // 情報抽出エラーの場合もエラーとして扱う
        if (hasExtractionError) {
          throw new Error('情報抽出処理でエラーが発生: ' + extractionErrorMessage);
        }

        // 抽出情報のチェック
        Logger.log('情報抽出結果: ' + JSON.stringify({
          sales_company: extractedInfo.sales_company,
          call_status1: extractedInfo.call_status1,
          call_status2: extractedInfo.call_status2,
          summaryLength: extractedInfo.summary ? extractedInfo.summary.length : 0
        }));

        // 処理終了時間を取得
        var fileEndTime = new Date();
        var fileEndTimeStr = Utilities.formatDate(fileEndTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

        // ファイルからメタデータを抽出
        var metadata = extractMetadataFromFile(file.getName());

        // 必要なメタデータが取得できなかった場合はファイル名から直接抽出を試みる
        if (!metadata.recordId) {
          // ファイル名からrecordIdを直接抽出（例：zoom_call_20250519015737_96f847c2df6f46b684b046f2046efc8d.mp3）
          var fileNameMatch = file.getName().match(/zoom_call_\d+_([a-f0-9]+)\.mp3/i);
          if (fileNameMatch && fileNameMatch[1]) {
            metadata.recordId = fileNameMatch[1];
            Logger.log('メタデータ抽出失敗後、ファイル名から直接record_idを抽出: ' + metadata.recordId);

            // Recordingsシートから電話番号情報を補完
            var phoneNumbers = getPhoneNumbersFromRecordingsSheet(metadata.recordId);
            if (phoneNumbers) {
              metadata.salesPhoneNumber = phoneNumbers.salesPhoneNumber || '';
              metadata.customerPhoneNumber = phoneNumbers.customerPhoneNumber || '';
            }

            // ファイル名から日付と時間を抽出
            var timestampMatch = file.getName().match(/zoom_call_(\d+)_/i);
            if (timestampMatch && timestampMatch[1] && timestampMatch[1].length >= 14) {
              var timestamp = timestampMatch[1];
              var year = timestamp.substring(0, 4);
              var month = timestamp.substring(4, 6);
              var day = timestamp.substring(6, 8);
              var hour = timestamp.substring(8, 10);
              var minute = timestamp.substring(10, 12);
              var second = timestamp.substring(12, 14);

              metadata.callDate = year + '-' + month + '-' + day;
              metadata.callTime = hour + ':' + minute + ':' + second;
              Logger.log('ファイル名からタイムスタンプを抽出: ' + metadata.callDate + ' ' + metadata.callTime);
            } else {
              // 現在の日時をセット
              var now = new Date();
              metadata.callDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
              metadata.callTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
              Logger.log('タイムスタンプ抽出失敗のため現在時刻を使用: ' + metadata.callDate + ' ' + metadata.callTime);
            }
          } else {
            throw new Error('メタデータから録音IDを取得できませんでした: ' + file.getName());
          }
        }

        // metadata.recordId が取得できてる場合でも、時間情報の欠損に備えて
        // ファイル名からのタイムスタンプ抽出を試みる部分を追加
        if (metadata.recordId && (!metadata.callDate || !metadata.callTime)) {
          Logger.log('record_idはあるが時間情報が不足しているため、補完を試みます');

          // ファイル名から日付と時間を抽出
          var timestampMatch = file.getName().match(/zoom_call_(\d+)_/i);
          if (timestampMatch && timestampMatch[1] && timestampMatch[1].length >= 14) {
            var timestamp = timestampMatch[1];
            var year = timestamp.substring(0, 4);
            var month = timestamp.substring(4, 6);
            var day = timestamp.substring(6, 8);
            var hour = timestamp.substring(8, 10);
            var minute = timestamp.substring(10, 12);
            var second = timestamp.substring(12, 14);

            metadata.callDate = year + '-' + month + '-' + day;
            metadata.callTime = hour + ':' + minute + ':' + second;
            Logger.log('ファイル名から時間情報を補完: ' + metadata.callDate + ' ' + metadata.callTime);
          }
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

          // メタデータからrecordIdが取得できない場合はファイル名から直接抽出
          if (!errorMetadata || !errorMetadata.recordId) {
            var fileNameMatch = file.getName().match(/zoom_call_\d+_([a-f0-9]+)\.mp3/i);
            if (fileNameMatch && fileNameMatch[1]) {
              if (!errorMetadata) errorMetadata = {};
              errorMetadata.recordId = fileNameMatch[1];
              Logger.log('エラー処理中にファイル名から直接record_idを抽出: ' + errorMetadata.recordId);

              // ファイル名から日付と時間を抽出
              var timestampMatch = file.getName().match(/zoom_call_(\d+)_/i);
              if (timestampMatch && timestampMatch[1] && timestampMatch[1].length >= 14) {
                var timestamp = timestampMatch[1];
                var year = timestamp.substring(0, 4);
                var month = timestamp.substring(4, 6);
                var day = timestamp.substring(6, 8);
                var hour = timestamp.substring(8, 10);
                var minute = timestamp.substring(10, 12);
                var second = timestamp.substring(12, 14);

                errorMetadata.callDate = year + '-' + month + '-' + day;
                errorMetadata.callTime = hour + ':' + minute + ':' + second;
                Logger.log('エラー処理中にファイル名から時間情報を抽出: ' + errorMetadata.callDate + ' ' + errorMetadata.callTime);
              }
            }
          }

          // 時間情報が欠けている場合も補完
          if (errorMetadata && errorMetadata.recordId && (!errorMetadata.callDate || !errorMetadata.callTime)) {
            // ファイル名から日付と時間を抽出
            var timestampMatch = file.getName().match(/zoom_call_(\d+)_/i);
            if (timestampMatch && timestampMatch[1] && timestampMatch[1].length >= 14) {
              var timestamp = timestampMatch[1];
              var year = timestamp.substring(0, 4);
              var month = timestamp.substring(4, 6);
              var day = timestamp.substring(6, 8);
              var hour = timestamp.substring(8, 10);
              var minute = timestamp.substring(10, 12);
              var second = timestamp.substring(12, 14);

              errorMetadata.callDate = year + '-' + month + '-' + day;
              errorMetadata.callTime = hour + ':' + minute + ':' + second;
              Logger.log('エラー処理中に時間情報を補完: ' + errorMetadata.callDate + ' ' + errorMetadata.callTime);
            }
          }

          if (errorMetadata && errorMetadata.recordId) {
            // Recordingsシートの文字起こし状態を更新（エラー）
            updateTranscriptionStatusByRecordId(
              errorMetadata.recordId,
              'ERROR: ' + error.toString().substring(0, 100), // エラーメッセージは短く切り詰める
              fileStartTimeStr,
              fileEndTimeStr
            );

            // エラー情報をログに記録（call_recordsシートには出力しない）
            try {
              if (errorMetadata && errorMetadata.recordId && errorMetadata.callDate && errorMetadata.callTime) {
                // エラー情報の構築（ログ用）
                var errorCallData = {
                  fileName: file.getName(),
                  recordId: errorMetadata.recordId,
                  callDate: errorMetadata.callDate,
                  callTime: errorMetadata.callTime,
                  salesPhoneNumber: errorMetadata.salesPhoneNumber || '',
                  customerPhoneNumber: errorMetadata.customerPhoneNumber || '',
                  salesCompany: '【エラー】情報抽出失敗',
                  customerName: '【エラー】情報抽出失敗',
                  callStatus1: '',
                  callStatus2: '',
                  reasonForRefusal: '',
                  reasonForRefusalCategory: '',
                  reasonForAppointment: '',
                  reasonForAppointmentCategory: '',
                  summary: 'エラーが発生しました: ' + error.toString().substring(0, 100),
                  transcription: 'エラー発生：' + error.toString()
                };

                // エラー時はcall_recordsシートに出力しない（ログのみ）
                Logger.log('エラー情報: ' + JSON.stringify(errorCallData));
              } else {
                Logger.log('必須メタデータが不足しているためスプレッドシートへの書き込みをスキップ');
              }
            } catch (logError) {
              Logger.log('エラー情報のログ出力中にエラー: ' + logError.toString());
            }
          } else {
            Logger.log('エラー処理中にrecord_idが特定できませんでした: ' + file.getName());
          }
        } catch (metaError) {
          Logger.log('エラー処理中のメタデータ取得に失敗: ' + metaError.toString());

          // 最後の手段としてファイル名からの抽出を試みる
          try {
            var lastChanceMatch = file.getName().match(/zoom_call_\d+_([a-f0-9]+)\.mp3/i);
            if (lastChanceMatch && lastChanceMatch[1]) {
              var recordId = lastChanceMatch[1];
              Logger.log('エラー処理の最終手段でファイル名からrecord_idを抽出: ' + recordId);

              // ファイル名から日付と時間を抽出
              var timestampMatch = file.getName().match(/zoom_call_(\d+)_/i);
              var callDate = null;
              var callTime = null;

              if (timestampMatch && timestampMatch[1] && timestampMatch[1].length >= 14) {
                var timestamp = timestampMatch[1];
                var year = timestamp.substring(0, 4);
                var month = timestamp.substring(4, 6);
                var day = timestamp.substring(6, 8);
                var hour = timestamp.substring(8, 10);
                var minute = timestamp.substring(10, 12);
                var second = timestamp.substring(12, 14);

                callDate = year + '-' + month + '-' + day;
                callTime = hour + ':' + minute + ':' + second;
                Logger.log('最終手段でファイル名から時間情報を抽出: ' + callDate + ' ' + callTime);
              } else {
                // 現在時刻を使用
                var now = new Date();
                callDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
                callTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
                Logger.log('最終手段で現在時刻を使用: ' + callDate + ' ' + callTime);
              }

              updateTranscriptionStatusByRecordId(
                recordId,
                'ERROR: ' + error.toString().substring(0, 100), // エラーメッセージは短く切り詰める
                fileStartTimeStr,
                fileEndTimeStr
              );

              // 最終手段でもスプレッドシートに書き込み
              try {
                if (recordId && callDate && callTime) {
                  // スプレッドシートに書き込み
                  var lastChanceData = {
                    fileName: file.getName(),
                    recordId: recordId,
                    callDate: callDate,
                    callTime: callTime,
                    salesPhoneNumber: '',
                    customerPhoneNumber: '',
                    salesCompany: '【最終手段】情報抽出失敗',
                    customerName: '【最終手段】情報抽出失敗',
                    callStatus1: '',
                    callStatus2: '',
                    reasonForRefusal: '',
                    reasonForRefusalCategory: '',
                    reasonForAppointment: '',
                    reasonForAppointmentCategory: '',
                    summary: '最終手段による処理: ' + error.toString().substring(0, 100),
                    transcription: '最終手段による処理：' + error.toString()
                  };

                  // エラー時はcall_recordsシートに出力しない（ログのみ）
                  Logger.log('最終手段処理情報: ' + JSON.stringify(lastChanceData));
                }
              } catch (lastLogError) {
                Logger.log('最終手段処理のログ出力中にエラー: ' + lastLogError.toString());
              }
            }
          } catch (lastError) {
            Logger.log('すべての方法でrecord_id抽出に失敗: ' + lastError.toString());
          }
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

    // 管理者に通知メールを送信（エラーが発生した場合のみ）
    if (results.error > 0) {
      try {
        var settings = getSystemSettings();
        var adminEmails = settings.ADMIN_EMAILS || [];

        for (var i = 0; i < adminEmails.length; i++) {
          NotificationService.sendRealtimeErrorNotification(adminEmails[i], results);
        }
        Logger.log('リアルタイムエラー通知メールを送信しました');
      } catch (notificationError) {
        Logger.log('通知メール送信エラー: ' + notificationError.toString());
      }
    }

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
  return RecoveryService.recoverInterruptedFiles();
}

/**
 * Recordingsシートからステータスが"INTERRUPTED"のファイルを取得
 * @return {Array} 中断されたファイルの配列
 */
function getInterruptedFilesFromSheet() {
  try {
    var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
    if (!spreadsheetId) return [];

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Recordings');

    if (!sheet) {
      Logger.log('Recordingsシートが見つかりません');
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var interruptedFiles = [];
    var settings = getSystemSettings();
    var sourceFolderId = settings.SOURCE_FOLDER_ID;

    // ヘッダー行をスキップして2行目から処理
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      // status_transcription（12列目）をチェック
      if (row[11] === 'INTERRUPTED') {
        var recordId = row[0]; // record_idは1列目

        // record_idに対応するファイルを探す
        var filePattern = `zoom_call_*_${recordId}.mp3`;
        var files = DriveApp.searchFiles(`title contains "${recordId}" and mimeType contains "audio"`);

        while (files.hasNext()) {
          var file = files.next();
          // すでに処理対象フォルダにあるか確認
          var parents = file.getParents();
          var isInSourceFolder = false;

          while (parents.hasNext()) {
            var parent = parents.next();
            if (parent.getId() === sourceFolderId) {
              isInSourceFolder = true;
              break;
            }
          }

          // 処理対象フォルダにないファイルだけを追加
          if (!isInSourceFolder) {
            interruptedFiles.push(file);
            Logger.log('INTERRUPTEDステータスのファイルを追加: ' + file.getName());
            break; // 最初に見つかったファイルのみを使用
          }
        }
      }
    }

    return interruptedFiles;
  } catch (e) {
    Logger.log('中断ファイルの取得でエラー: ' + e.toString());
    return [];
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
        // 列数チェック
        var maxColumns = sheet.getMaxColumns();
        if (maxColumns < 14) {
          Logger.log('Recordingsシートの列数が不足しています（現在: ' + maxColumns + ', 必要: 14）。列を追加します。');
          sheet.insertColumnsAfter(maxColumns, 14 - maxColumns);
        }
        
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

    // 中断ファイル復旧処理を実行
    Logger.log('スケジュール実行: 中断ファイルの復旧処理を開始します...');
    var recoveryResult = recoverInterruptedFiles();
    Logger.log('スケジュール実行: 中断ファイル復旧結果: ' + recoveryResult);

    // エラーファイル復旧処理を実行（1回限りのリトライ）
    Logger.log('スケジュール実行: エラーファイルの復旧処理を開始します...');
    var errorRecoveryResult = recoverErrorFiles();
    Logger.log('スケジュール実行: エラーファイル復旧結果: ' + errorRecoveryResult);

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
  return ErrorHandler.safeExecute(function () {
    Logger.log('メイン処理スタート: ' + new Date().toString());

    // 0. 中断されたファイルの復旧処理
    Logger.log('中断されたファイルの復旧処理を開始します...');
    var recoveryResult = RecoveryService.recoverInterruptedFiles();
    Logger.log('中断ファイル復旧結果: ' + recoveryResult);

    // 0-2. エラーファイルの復旧処理（1回限りのリトライ）
    Logger.log('エラーファイルの復旧処理を開始します...');
    var errorRecoveryResult = RecoveryService.recoverErrorFiles();
    Logger.log('エラーファイル復旧結果: ' + errorRecoveryResult);

    // 1. Zoom録音ファイルの取得処理
    Logger.log('Recordingsシートからの録音ファイル取得処理を開始します（タイムスタンプが古い順に処理）...');
    var recordingsResult = ZoomphoneProcessor.processRecordingsFromSheet();
    Logger.log('Recordingsシートからの録音取得結果: ' + JSON.stringify(recordingsResult));

    // 2. メイン処理バッチ（文字起こし、情報抽出）
    processBatch();

    Logger.log('メイン処理完了: ' + new Date().toString());
    return '処理完了';
  }, 'メイン処理', {
    maxRetries: 0,
    category: ErrorHandler.ERROR_CATEGORIES.SYSTEM
  });
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
      fetchStatusCounts: {},
      transcriptionStatusCounts: {},
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
        var fetchStatus = row[9]; // status_fetch
        var transcStatus = row[11]; // status_transcription

        Logger.log('ステータス: fetch=' + fetchStatus + ', transcription=' + transcStatus);

        // PROCESSEDは常に成功としてカウント
        if (fetchStatus === 'PROCESSED') {
          summary.success++;
          Logger.log('成功としてカウント');
        } else if (fetchStatus === 'ERROR' || (fetchStatus && fetchStatus.indexOf('ERROR') === 0)) {
          summary.error++;
          Logger.log('エラーとしてカウント');
        }

        // status_fetchの集計
        if (fetchStatus) {
          if (!summary.fetchStatusCounts[fetchStatus]) {
            summary.fetchStatusCounts[fetchStatus] = 1;
          } else {
            summary.fetchStatusCounts[fetchStatus]++;
          }
        }

        // status_transcriptionの集計
        if (transcStatus) {
          if (!summary.transcriptionStatusCounts[transcStatus]) {
            summary.transcriptionStatusCounts[transcStatus] = 1;
          } else {
            summary.transcriptionStatusCounts[transcStatus]++;
          }
        }
      }
    }

    Logger.log('集計結果: 成功=' + summary.success + '件, エラー=' + summary.error + '件');
    Logger.log('取得ステータス別集計: ' + JSON.stringify(summary.fetchStatusCounts));
    Logger.log('文字起こしステータス別集計: ' + JSON.stringify(summary.transcriptionStatusCounts));

    return summary;
  } catch (e) {
    Logger.log('サマリー集計でエラー: ' + e.toString());
    return {
      success: 0,
      error: 0,
      fetchStatusCounts: {},
      transcriptionStatusCounts: {},
      details: []
    };
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
    .addSeparator()
    .addSubMenu(ui.createMenu('録音取得')
      .addItem('直近1時間の録音を取得', 'TriggerManager.fetchLastHourRecordings')
      .addItem('直近2時間の録音を取得', 'TriggerManager.fetchLast2HoursRecordings')
      .addItem('直近6時間の録音を取得', 'TriggerManager.fetchLast6HoursRecordings')
      .addItem('直近24時間の録音を取得', 'TriggerManager.fetchLast24HoursRecordings')
      .addItem('直近48時間の録音を取得', 'TriggerManager.fetchLast48HoursRecordings')
      .addItem('すべての未処理録音を取得', 'TriggerManager.fetchAllPendingRecordings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('トリガー設定')
      .addItem('🔧 全トリガー設定（復旧機能込み）', 'TriggerManager.setupAllTriggers')
      .addItem('⚙️ 基本トリガーのみ設定', 'TriggerManager.setupBasicTriggers')
      .addItem('🔄 復旧トリガーのみ追加', 'TriggerManager.setupRecoveryTriggersOnly')
      .addSeparator()
      .addItem('❌ 全トリガー削除', 'TriggerManager.deleteAllTriggers')
      .addItem('❌ 復旧トリガーのみ削除', 'TriggerManager.removeRecoveryTriggers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('復旧・メンテナンス')
      .addItem('🔍 部分的失敗を検知・復旧', 'detectAndRecoverPartialFailures')
      .addItem('🔄 PENDINGの文字起こしを再処理', 'resetPendingTranscriptions')
      .addItem('📁 エラーフォルダのファイルを強制復旧', 'forceRecoverAllErrorFiles')
      .addItem('⚠️ 中断ファイルを復旧', 'recoverInterruptedFiles')
      .addItem('🚨 エラーファイルを復旧', 'recoverErrorFiles'))
    .addToUi();
}

/**
 * エラーフォルダにあるファイルを未処理フォルダに戻す
 * エラーになったファイルを1回だけリトライする
 */
function recoverErrorFiles() {
  return RecoveryService.recoverErrorFiles();
}

/**
 * 特別復旧処理: PENDINGステータスのままになっている文字起こしを再処理
 */
function resetPendingTranscriptions() {
  return RecoveryService.resetPendingTranscriptions();
}

/**
 * 特別復旧処理: エラーフォルダに残っているファイルを一括で未処理フォルダに戻す
 * 通常のrecoverErrorFilesではスキップされている[RETRIED]マーク付きのファイルも対象
 */
function forceRecoverAllErrorFiles() {
  return RecoveryService.forceRecoverAllErrorFiles();
}

/**
 * 部分的失敗を検知して復旧する関数
 * 文字起こしが SUCCESS になっているが、実際にはエラーが含まれているケースを検知
 */
function detectAndRecoverPartialFailures() {
  return PartialFailureDetector.detectAndRecoverPartialFailures();
}

/**
 * ファイルをエラーフォルダに移動
 * @param {string} recordId - 録音ID
 * @return {boolean} - ファイルが見つかって移動できた場合true
 */
function moveFileToErrorFolder(recordId) {
  return PartialFailureDetector.moveFileToErrorFolder(recordId);
}

/**
 * 全トリガー設定（復旧機能込み）- Apps Script直接実行用
 * @return {string} 実行結果メッセージ
 */
function setupAllTriggersWithRecovery() {
  return TriggerManager.setupAllTriggers();
}

/**
 * 基本トリガーのみ設定 - Apps Script直接実行用
 * @return {string} 実行結果メッセージ
 */
function setupBasicTriggersOnly() {
  return TriggerManager.setupBasicTriggers();
}

/**
 * 復旧トリガーのみ追加 - Apps Script直接実行用
 * @return {string} 実行結果メッセージ
 */
function addRecoveryTriggersOnly() {
  return TriggerManager.addRecoveryTriggers();
}

/**
 * 全復旧処理を一括実行 - Apps Script直接実行用
 * 部分的失敗検知 → PENDING復旧 → エラーファイル復旧の順で実行
 * @return {string} 実行結果メッセージ
 */
function runFullRecoveryProcess() {
  return RecoveryService.runFullRecoveryProcess();
}

/**
 * call_recordsシートで文字起こし内容にエラーが含まれているかチェック
 * @param {string} recordId - 録音ID
 * @return {Object} - {found: boolean, issue: string}
 */
function checkForErrorInTranscription(recordId) {
  return PartialFailureDetector.checkForErrorInTranscription(recordId);
}
