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
      var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
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