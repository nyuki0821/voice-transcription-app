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
      var settings = loadModuleSettings();
      var folderId = settings.SOURCE_FOLDER_ID;
      if (!folderId) throw new Error('SOURCE_FOLDER_ID が未設定です');

      var now = new Date();
      fromDate = fromDate || new Date(now.getTime() - 24 * 60 * 60 * 1000);
      toDate   = toDate   || now;

      var page = 1, saved = 0;
      while (true) {
        var res  = ZoomphoneService.listCallRecordings(fromDate, toDate, page, 30);
        var recs = res.recordings || [];
        if (recs.length === 0) break;

        recs.forEach(function (rec) {
          var id  = rec.id;
          var url = rec.download_url || rec.download_url_with_token;
          if (!id || !url) return;
          if (getProcessedCallIds().indexOf(id) !== -1) return;

          var blob = ZoomphoneService.downloadBlob(url);
          var file = ZoomphoneService.saveRecordingToDrive(blob, rec, folderId);
          if (file) {
            markCallAsProcessed(id);
            logProcessedIdToSheet(id);
            saved++;
          }
        });

        if (!res.next_page_token && recs.length < 30) break;
        page++;
      }

      return { success: true, saved: saved };
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
    var sp  = PropertiesService.getScriptProperties();
    var raw = sp.getProperty(key);
    if (raw) {
      try { return JSON.parse(raw); } catch (e) { /* ignore */ }
    }
    var def = { SOURCE_FOLDER_ID: '', PROCESSED_IDS_PROP: 'PROCESSED_CALL_IDS' };
    sp.setProperty(key, JSON.stringify(def));
    return def;
  }

  return {
    processRecordings: processRecordings,
    processCallRecordings: processCallRecordings,
    getProcessedCallIds: getProcessedCallIds,
    markCallAsProcessed: markCallAsProcessed,
    logProcessedIdToSheet: logProcessedIdToSheet
  };
})();