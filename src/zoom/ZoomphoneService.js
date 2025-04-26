/**
 * Zoom Phone サービスモジュール
 *  ├─ 録音一覧取得   : listCallRecordings()
 *  ├─ 録音ファイル取得 : getRecordingBlob()
 *  ├─ URL直接ダウンロード : downloadBlob()
 *  └─ Drive 保存     : saveRecordingToDrive()
 */
var ZoomphoneService = (function () {

  /**
   * 指定期間の録音一覧を取得する（録音が存在する通話のみ返る）
   * @param {Date}   fromDate   取得開始日
   * @param {Date}   toDate     取得終了日
   * @param {number} [page=1]   ページ番号 (1‑based)
   * @param {number} [pageSize=30] ページサイズ
   * @return {Object} Zoom API レスポンス ({ recordings: [...] })
   */
  function listCallRecordings(fromDate, toDate, page, pageSize) {
    try {
      var fromStr = Utilities.formatDate(fromDate, 'UTC', 'yyyy-MM-dd');
      var toStr   = Utilities.formatDate(toDate,   'UTC', 'yyyy-MM-dd');

      var params = [
        'from=' + fromStr,
        'to='   + toStr,
        'page_size='   + (pageSize || 30),
        'page_number=' + (page     || 1)
      ];
      var endpoint = '/phone/recordings?' + params.join('&');

      var res = ZoomAPIManager.sendRequest(endpoint, 'get');
      return res;
    } catch (err) {
      throw err;
    }
  }

  /**
   * recordingId から録音ファイル (Blob) を取得する
   * @param  {string} recordingId
   * @return {Blob|null} 取得できなければ null
   */
  function getRecordingBlob(recordingId) {
    try {
      if (!recordingId) throw new Error('recordingId がありません');

      /* --- Step 1: recordingId → file_id を取得 ---------------- */
      var info = ZoomAPIManager.sendRequest('/phone/recordings/' + recordingId, 'get');
      var fileId =
        info.file_id ||
        (info.recording_files && info.recording_files.length && info.recording_files[0].file_id);

      if (!fileId) {
        return null;
      }

      /* --- Step 2: file_id でダウンロード ----------------------- */
      var blob = ZoomAPIManager.sendRequest(
        '/phone/recording/download/' + fileId,
        'get',
        null,
        {
          responseType: 'blob',
          muteHttpExceptions: true,
          contentType: 'application/octet-stream'
        }
      );
      return blob;
    } catch (err) {
      // 404（録音削除済み）などは null 扱い
      if (err && err.toString().indexOf('404') !== -1) {
        return null;
      }
      throw err;
    }
  }

  /**
   * URL から録音ファイル (Blob) を直接ダウンロードする
   * @param  {string} url ダウンロードURL
   * @return {Blob|null} 取得できなければ null
   */
  function downloadBlob(url) {
    try {
      if (!url) throw new Error('ダウンロードURLがありません');

      var token = ZoomAPIManager.getAccessToken();
      var options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();

      if (responseCode < 200 || responseCode >= 300) {
        return null;
      }

      return response.getBlob();
    } catch (err) {
      return null;
    }
  }

  /**
   * @param {Blob}   blob        録音ファイル
   * @param {Object} recMeta     listCallRecordings() の recordings 配列要素
   * @param {string} folderId    Drive フォルダ ID
   * @return {GoogleAppsScript.Drive.File|null}
   */
  function saveRecordingToDrive(blob, recMeta, folderId) {
    try {
      if (!blob) return null; // 取得失敗時
      if (!folderId) throw new Error('保存先フォルダ ID がありません');
      
      // 日時情報の取得（複数の可能性を考慮）
      var timestamp = null;
      if (recMeta.start_time) {
        timestamp = new Date(recMeta.start_time);
      } else if (recMeta.date_time) {
        timestamp = new Date(recMeta.date_time);
      } else if (recMeta.created_at) {
        timestamp = new Date(recMeta.created_at);
      } else {
        // どのフィールドも存在しない場合は現在時刻を使用
        timestamp = new Date();
      }
      
      // 録音IDの取得
      var recordingId = recMeta.recording_id || recMeta.id || 'unknown';
      
      // 日付をフォーマット
      var dateStr = Utilities.formatDate(timestamp, 'UTC', 'yyyyMMddHHmmss');
      var fileName = 'zoom_call_' + dateStr + '_' + recordingId + '.mp3';
      
      var folder = DriveApp.getFolderById(folderId);
      var file = folder.createFile(blob.setName(fileName));
      return file;
    } catch (err) {
      throw err;
    }
  }

  /* ------------------------------------------------------------------
   * 公開メソッド
   * ------------------------------------------------------------------*/
  return {
    listCallRecordings : listCallRecordings,
    getRecordingBlob   : getRecordingBlob,
    downloadBlob       : downloadBlob,
    saveRecordingToDrive: saveRecordingToDrive
  };
})();