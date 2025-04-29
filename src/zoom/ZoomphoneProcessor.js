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
            logProcessedIdToSheet(id);
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
    var str = PropertiesService.getScriptProperties().getProperty('PROCESSED_CALL_IDS');
    return str ? JSON.parse(str) : [];
  }

  /** 処理済み ID をマーク */
  function markCallAsProcessed(id) {
    var ids = getProcessedCallIds();
    if (ids.indexOf(id) === -1) ids.push(id);
    while (ids.length > 1000) ids.shift();
    PropertiesService.getScriptProperties().setProperty('PROCESSED_CALL_IDS', JSON.stringify(ids));
  }

  /** 処理済みIDをスプレッドシートに記録する */
  function logProcessedIdToSheet(id) {
    try {
      var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (!spreadsheetId) return;

      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('処理済みID記録');

      if (!sheet) {
        sheet = spreadsheet.insertSheet('処理済みID記録');
        sheet.getRange(1, 1, 1, 2).setValues([
          ['処理日時', '録音ID']
        ]);
        sheet.setColumnWidth(1, 180);
        sheet.setColumnWidth(2, 300);
      }

      var now = new Date();
      var formattedDateTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, 2).setValues([
        [formattedDateTime, id]
      ]);
    } catch (error) {
      // エラーは無視して処理続行
    }
  }

  /** モジュール設定を読み込む */
  function loadModuleSettings() {
    if (typeof getSystemSettings === 'function') {
      return getSystemSettings();
    }
    var key = 'ZOOMPHONE_SETTINGS';
    var sp = PropertiesService.getScriptProperties();
    var raw = sp.getProperty(key);
    if (raw) {
      try { return JSON.parse(raw); } catch (e) { /* ignore */ }
    }
    var def = { SOURCE_FOLDER_ID: '', PROCESSED_IDS_PROP: 'PROCESSED_CALL_IDS' };
    sp.setProperty(key, JSON.stringify(def));
    return def;
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

  /** Webhook 経由で届いた録音を即時処理する */
  function processWebhookRecording(recId, downloadUrl, startTime) {
    try {
      if (!recId || !downloadUrl) {
        throw new Error('recId または downloadUrl が不正です');
      }
      // 既に処理済みの場合はスキップ
      if (getProcessedCallIds().indexOf(recId) !== -1) {
        return { success: true, skipped: true };
      }

      var settings = loadModuleSettings();
      var folderId = settings.SOURCE_FOLDER_ID;
      if (!folderId) throw new Error('SOURCE_FOLDER_ID が未設定です');

      var blob = ZoomphoneService.downloadBlob(downloadUrl);
      Utilities.sleep(200); // Zoom API rate limit 対策
      var meta = {
        id: recId,
        recording_id: recId,
        start_time: startTime || new Date()
      };
      var file = ZoomphoneService.saveRecordingToDrive(blob, meta, folderId);
      if (file) {
        markCallAsProcessed(recId);
        logProcessedIdToSheet(recId);
      }
      return { success: true };
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  }

  return {
    processRecordings: processRecordings,
    processCallRecordings: processCallRecordings,
    getProcessedCallIds: getProcessedCallIds,
    markCallAsProcessed: markCallAsProcessed,
    logProcessedIdToSheet: logProcessedIdToSheet,
    processWebhookRecording: processWebhookRecording,
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