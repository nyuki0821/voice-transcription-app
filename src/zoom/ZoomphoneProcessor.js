/**
 * Zoomphone Processor
 * - 指定期間内の録音を取得しGoogle Driveに保存
 */
var ZoomphoneProcessor = (function () {
  /**
   * 録音取得処理
   * @param {Date=} fromDate デフォルト = 24h 前
   * @param {Date=} toDate   デフォルト = 現在
   */
  function processRecordings(fromDate, toDate) {
    try {
      var processStart = new Date().getTime();
      var settings = loadModuleSettings();
      var folderId = settings.SOURCE_FOLDER_ID;
      if (!folderId) throw new Error('SOURCE_FOLDER_ID が未設定です');

      var now = new Date();
      fromDate = fromDate || new Date(now.getTime() - 24 * 60 * 60 * 1000);
      toDate = toDate || now;

      var page = 1, saved = 0, fetched = 0;
      while (true) {
        var res = ZoomphoneService.listCallRecordings(fromDate, toDate, page, 30);
        Utilities.sleep(200); // Zoom API rate limit 対策
        var recs = res.recordings || [];
        fetched += recs.length;
        if (recs.length === 0) break;

        recs.forEach(function (rec) {
          var id = rec.id;
          var url = rec.download_url || rec.download_url_with_token;
          if (!id || !url) return;
          if (getProcessedCallIds().indexOf(id) !== -1) return;

          var blob = ZoomphoneService.downloadBlob(url);
          Utilities.sleep(200); // Zoom API rate limit 対策
          var file = ZoomphoneService.saveRecordingToDrive(blob, rec, folderId);
          if (file) {
            markCallAsProcessed(id);
            logProcessedIdToSheet(id, rec);
            saved++;
          }
        });

        // --- 実行時間チェック: 4分を超える前に中断し後続トリガーを設定 ----------------
        if (new Date().getTime() - processStart > 4 * 60 * 1000) {
          scheduleContinuation(fromDate, toDate, page + 1);
          break;
        }

        if (!res.next_page_token && recs.length < 30) break;
        page++;
      }

      return { success: true, fetched: fetched, saved: saved };
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  }

  /** 旧関数名との互換 */
  function processCallRecordings(f, t) {
    return processRecordings(f, t);
  }

  /** 処理済み ID を取得 */
  function getProcessedCallIds() {
    var str = EnvironmentConfig.get('PROCESSED_CALL_IDS', '');
    return str ? JSON.parse(str) : [];
  }

  /** 処理済み ID をマーク */
  function markCallAsProcessed(id) {
    var ids = getProcessedCallIds();
    if (ids.indexOf(id) === -1) ids.push(id);
    while (ids.length > 1000) ids.shift();
    PropertiesService.getScriptProperties().setProperty('PROCESSED_CALL_IDS', JSON.stringify(ids));
  }

  /** 処理済みIDをスプレッドシートに記録する 
   * @param {string} id - 録音ID
   * @param {Object=} metadata - 追加情報（オプション）
   */
  function logProcessedIdToSheet(id, metadata) {
    try {
      var spreadsheetId = EnvironmentConfig.get('PROCESSED_IDS_SHEET_ID', '');
      if (!spreadsheetId) return;

      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('processed_ids');

      if (!sheet) {
        sheet = spreadsheet.insertSheet('processed_ids');
        sheet.getRange(1, 1, 1, 6).setValues([
          ['timestamp', 'recording_id', 'call_datetime', 'duration_seconds', 'caller_number', 'called_number']
        ]);
        // カラム幅を設定
        sheet.setColumnWidth(1, 180); // timestamp
        sheet.setColumnWidth(2, 300); // recording_id
        sheet.setColumnWidth(3, 180); // call_datetime
        sheet.setColumnWidth(4, 120); // duration_seconds
        sheet.setColumnWidth(5, 150); // caller_number
        sheet.setColumnWidth(6, 150); // called_number
      }

      var now = new Date();
      var formattedDateTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

      // メタデータから追加情報を取得
      var callDatetime = '';
      var durationSeconds = '';
      var callerNumber = '';
      var calledNumber = '';

      if (metadata) {
        Logger.log('録音メタデータ: ' + JSON.stringify(metadata));

        // 開始時間の取得（複数の可能性を考慮）
        if (metadata.start_time) {
          var startTime = new Date(metadata.start_time);
          callDatetime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        } else if (metadata.date_time) {
          var dateTime = new Date(metadata.date_time);
          callDatetime = Utilities.formatDate(dateTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        } else if (metadata.created_at) {
          var createdAt = new Date(metadata.created_at);
          callDatetime = Utilities.formatDate(createdAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        }

        // 通話時間
        durationSeconds = metadata.duration || metadata.duration_seconds || '';

        // 電話番号情報
        if (metadata.caller) {
          callerNumber = metadata.caller.phone_number || '';
        } else {
          callerNumber = metadata.caller_number || metadata.from || '';
        }

        if (metadata.callee) {
          calledNumber = metadata.callee.phone_number || '';
        } else {
          // 国際電話形式（+81）から日本の一般的な形式（0XXX）に変換（ハイフンなし）
          var rawNumber = metadata.callee_number || metadata.to || '';
          if (rawNumber.startsWith('+81')) {
            // +81 を 0 に置き換える（文字列として扱う）
            calledNumber = '0' + rawNumber.substring(3);
          } else {
            calledNumber = rawNumber;
          }
        }
      }

      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, 6).setValues([
        [formattedDateTime, id, callDatetime, durationSeconds, callerNumber, calledNumber]
      ]);

      Logger.log('録音情報をスプレッドシートに保存: ID=' + id +
        ', 通話時間=' + durationSeconds +
        ', 発信=' + callerNumber +
        ', 着信=' + calledNumber);
    } catch (error) {
      // エラーは無視して処理続行
      Logger.log('処理済みID記録エラー: ' + error.toString());
    }
  }

  /**
   * Recordingsシートから未処理の録音情報を取得して処理する
   * @param {Date} [fromDate] 取得する開始日時（オプション）- 指定した場合、この時刻以降の録音のみを処理
   * @param {Date} [toDate] 取得する終了日時（オプション）- 指定した場合、この時刻以前の録音のみを処理
   * @returns {Object} 処理結果 { success: boolean, fetched: number, saved: number }
   */
  function processRecordingsFromSheet(fromDate, toDate) {
    try {
      var processStart = new Date().getTime();
      var settings = loadModuleSettings();
      var folderId = settings.SOURCE_FOLDER_ID;
      if (!folderId) throw new Error('SOURCE_FOLDER_ID が未設定です');

      // 時間フィルターの情報をログに記録
      if (fromDate instanceof Date && toDate instanceof Date) {
        Logger.log('指定された時間範囲: ' +
          Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 〜 ' +
          Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
      }

      // Recordingsシートから未処理のレコードを取得
      var pendingRecordings = getRecordingsToProcess(fromDate, toDate);
      Logger.log('Recordingsシートから取得した未処理録音: ' + pendingRecordings.length + '件');

      var saved = 0, fetched = pendingRecordings.length;

      for (var i = 0; i < pendingRecordings.length; i++) {
        var rec = pendingRecordings[i];
        var id = rec.recordingId;
        var url = rec.downloadUrl;

        if (!id || !url) {
          Logger.log('録音ID/URLがありません: ' + JSON.stringify(rec));
          continue;
        }

        if (getProcessedCallIds().indexOf(id) !== -1) {
          // 既に処理済みのIDはスキップ
          Logger.log('既に処理済みのID: ' + id);
          updateRecordingStatus(rec.rowIndex, 'DUPLICATE');
          continue;
        }

        var blob = ZoomphoneService.downloadBlob(url);
        Utilities.sleep(200); // Zoom API rate limit 対策

        if (!blob) {
          Logger.log('録音ファイルの取得に失敗: ' + id);
          updateRecordingStatus(rec.rowIndex, 'DOWNLOAD_ERROR');
          continue;
        }

        var recMeta = {
          id: id,
          date_time: rec.startTime,
          duration: rec.duration,
          caller_number: rec.phoneNumber,
          direction: 'unknown'
        };

        var file = ZoomphoneService.saveRecordingToDrive(blob, recMeta, folderId);
        if (file) {
          markCallAsProcessed(id);
          logProcessedIdToSheet(id, recMeta);
          updateRecordingStatus(rec.rowIndex, 'PROCESSED');
          saved++;
        } else {
          updateRecordingStatus(rec.rowIndex, 'SAVE_ERROR');
        }

        // 実行時間チェック: 4分を超える前に中断
        if (new Date().getTime() - processStart > 4 * 60 * 1000) {
          Logger.log('処理時間制限に達したため中断します。次回トリガーで続きを処理します。');
          break;
        }
      }

      return { success: true, fetched: fetched, saved: saved };
    } catch (e) {
      Logger.log('Recordingsシートからの処理でエラー: ' + e.toString());
      return { success: false, error: e.toString() };
    }
  }

  /**
   * Recordingsシートから未処理(PENDING)の録音情報を取得する
   * @param {Date} [fromDate] 取得する開始日時（オプション）
   * @param {Date} [toDate] 取得する終了日時（オプション）
   * @returns {Array} 未処理の録音情報配列
   */
  function getRecordingsToProcess(fromDate, toDate) {
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

      // ヘッダー行を考慮
      if (values.length <= 1) {
        return [];
      }

      var result = [];
      var hasTimeFilter = fromDate instanceof Date && toDate instanceof Date;
      if (hasTimeFilter) {
        Logger.log('時間範囲フィルターを適用: ' +
          Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 〜 ' +
          Utilities.formatDate(toDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
      }

      // fromDateとtoDateがJST（Asia/Tokyo）であることを前提とします
      var fromDateMs = hasTimeFilter ? fromDate.getTime() : 0;
      var toDateMs = hasTimeFilter ? toDate.getTime() : Number.MAX_SAFE_INTEGER;

      // ヘッダー行をスキップして2行目から処理
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        // [Timestamp, RecordingId, DownloadUrl, StartTime, PhoneNumber, Duration, Status]
        if (row[6] === 'PENDING') {
          var recordTimestamp = row[0]; // Timestamp
          var recordStartTime = row[3]; // StartTime

          // 時間範囲フィルターがあれば適用
          if (hasTimeFilter) {
            // タイムスタンプと開始時間のいずれか有効なものを使用
            var recordTimeObj = recordStartTime instanceof Date ? recordStartTime :
              (recordTimestamp instanceof Date ? recordTimestamp : new Date(recordStartTime || recordTimestamp));
            var recordTimeMs = recordTimeObj.getTime();

            // 指定時間範囲外ならスキップ
            if (recordTimeMs < fromDateMs || recordTimeMs > toDateMs) {
              continue;
            }
          }

          result.push({
            rowIndex: i + 1, // スプレッドシートの行番号（1始まり）
            timestamp: row[0],
            recordingId: row[1],
            downloadUrl: row[2],
            startTime: row[3],
            phoneNumber: row[4],
            duration: row[5]
          });
        }
      }

      // timestampが古い順にソートする
      result.sort(function (a, b) {
        // timestampが日付オブジェクトの場合
        if (a.timestamp instanceof Date && b.timestamp instanceof Date) {
          return a.timestamp.getTime() - b.timestamp.getTime();
        }
        // timestampが文字列の場合は日付オブジェクトに変換
        else {
          var dateA = new Date(a.timestamp);
          var dateB = new Date(b.timestamp);
          return dateA.getTime() - dateB.getTime();
        }
      });

      Logger.log('Recordingsシートから取得した未処理録音: ' + result.length + '件、古い順にソート済み');
      return result;
    } catch (e) {
      Logger.log('Recordingsシートの読み取りでエラー: ' + e.toString());
      return [];
    }
  }

  /**
   * Recordingsシートの録音ステータスを更新する
   * @param {number} rowIndex シートの行番号（1始まり）
   * @param {string} status 新しいステータス
   */
  function updateRecordingStatus(rowIndex, status) {
    try {
      var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
      if (!spreadsheetId) return;

      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('Recordings');

      if (!sheet) return;

      // Statusカラム（G列=7列目）を更新
      sheet.getRange(rowIndex, 7).setValue(status);
    } catch (e) {
      Logger.log('録音ステータスの更新でエラー: ' + e.toString());
    }
  }

  /** モジュール設定を読み込む */
  function loadModuleSettings() {
    var config = EnvironmentConfig.getConfig();
    return {
      SOURCE_FOLDER_ID: config.SOURCE_FOLDER_ID || '',
      PROCESSED_IDS_PROP: 'PROCESSED_CALL_IDS'
    };
  }

  /** 追加の継続トリガーを設定する（最短で1分後実行） */
  function scheduleContinuation(fromDate, toDate, nextPage) {
    try {
      var triggerFn = 'continueZoomProcessing';
      var existing = ScriptApp.getProjectTriggers().filter(function (t) {
        return t.getHandlerFunction() === triggerFn;
      });
      // すでにトリガーが存在する場合はスキップ
      if (existing.length === 0) {
        ScriptApp.newTrigger(triggerFn)
          .timeBased()
          .after(1 * 60 * 1000) // 1分後
          .create();
      }
      // 次ページ情報をスクリプトプロパティに一時保存
      var key = 'ZOOMPHONE_CONTINUATION';
      var data = {
        from: fromDate.getTime(),
        to: toDate.getTime(),
        page: nextPage
      };
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(data));
    } catch (e) {
      Logger.log('[ZoomphoneProcessor] Failed to schedule continuation: ' + e);
    }
  }

  /** 継続処理用のラッパー（内部実装） */
  function _continueZoomProcessing() {
    var key = 'ZOOMPHONE_CONTINUATION';
    var raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return;
    PropertiesService.getScriptProperties().deleteProperty(key);
    try {
      var obj = JSON.parse(raw);
      var fromDate = new Date(obj.from);
      var toDate = new Date(obj.to);
      var page = obj.page || 1;

      // processRecordings を継続ページから開始するために while 内ロジックを抽出するのが難しいので
      // とりあえず再度 processRecordings を呼び出す。次ページ処理は内部でスキップ済みページ数計算を行うか
      // もしくは listCallRecordings に page を渡して最初の呼び出しを調整する。
      // 簡易実装として from/to を再利用して processRecordings を呼び出す。
      processRecordings(fromDate, toDate);
    } catch (e) {
      Logger.log('[ZoomphoneProcessor] continuation error: ' + e);
    }
  }

  return {
    processRecordings: processRecordings,
    processCallRecordings: processCallRecordings,
    processRecordingsFromSheet: processRecordingsFromSheet,
    getProcessedCallIds: getProcessedCallIds,
    markCallAsProcessed: markCallAsProcessed,
    logProcessedIdToSheet: logProcessedIdToSheet,
    continueZoomProcessing: _continueZoomProcessing
  };
})();

/**
 * 継続処理用のラッパー（トリガーから呼び出すためグローバルに公開）
 */
function continueZoomProcessing() {
  if (typeof ZoomphoneProcessor !== 'undefined' && ZoomphoneProcessor.continueZoomProcessing) {
    ZoomphoneProcessor.continueZoomProcessing();
  }
}

/**
 * Recordingsシートの未処理レコードを処理するグローバル関数
 */
function processRecordingsFromSheet() {
  var result = ZoomphoneProcessor.processRecordingsFromSheet();
  Logger.log('Recordingsシート処理結果: ' + JSON.stringify(result));
  return result;
}