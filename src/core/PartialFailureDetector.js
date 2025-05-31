/**
 * 見逃しエラー検知サービス
 * 文字起こしが SUCCESS になっているが、実際にはエラーが含まれているケースを検知・復旧
 */
var PartialFailureDetector = (function () {

  /**
   * 見逃しエラーを検知して復旧する関数
   * 文字起こしが SUCCESS になっているが、実際にはエラーが含まれているケースを検知
   */
  function detectAndRecoverPartialFailures() {
    Logger.log('見逃しエラー検知・復旧処理を開始します');

    try {
      var sheet = ConfigManager.getRecordingsSheet();
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var idIndex = headers.indexOf(Constants.COLUMNS.RECORDINGS.ID);
      var statusTranscriptionIndex = headers.indexOf(Constants.COLUMNS.RECORDINGS.STATUS_TRANSCRIPTION);
      var updatedAtIndex = headers.indexOf(Constants.COLUMNS.RECORDINGS.UPDATED_AT);

      if (idIndex === -1 || statusTranscriptionIndex === -1) {
        Logger.log(Constants.ERROR_MESSAGES.REQUIRED_COLUMNS_NOT_FOUND);
        return createEmptyResult();
      }

      var results = FileMovementService.createResultObject();

      // SUCCESSステータスのレコードを検査
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var recordId = row[idIndex];
        var status = row[statusTranscriptionIndex];

        if (status === Constants.STATUS.SUCCESS) {
          results.total++;

          // call_recordsシートで実際の文字起こし内容をチェック
          var hasError = checkForErrorInTranscription(recordId);

          if (hasError.found) {
            Logger.log('見逃しエラーを検知: Record ID ' + recordId + ' - ' + hasError.issue);

            // ステータスをERROR_DETECTEDに更新
            sheet.getRange(i + 1, statusTranscriptionIndex + 1).setValue(Constants.STATUS.ERROR_DETECTED);

            // updated_atを更新
            if (updatedAtIndex !== -1) {
              sheet.getRange(i + 1, updatedAtIndex + 1).setValue(new Date());
            }

            // ファイルをエラーフォルダに移動
            var fileFound = moveFileToErrorFolder(recordId);

            results.recovered++;
            results.details.push({
              recordId: recordId,
              status: 'recovered',
              issue: hasError.issue,
              message: 'ステータスをERROR_DETECTEDに更新し、エラーフォルダに移動',
              fileFound: fileFound
            });
          }
        }
      }

      Logger.log('見逃しエラー検知・復旧処理完了: 検査=' + results.total + '件, 検知・復旧=' + results.recovered + '件');

      // 管理者に通知メールを送信（見逃しエラーが検知された場合のみ）
      if (results.total > 0 && results.recovered > 0) {
        try {
          var config = ConfigManager.getConfig();
          var adminEmails = config.ADMIN_EMAILS || [];

          for (var i = 0; i < adminEmails.length; i++) {
            NotificationService.sendPartialFailureDetectionSummary(adminEmails[i], results);
          }
          Logger.log('見逃しエラー検知結果の通知メールを送信しました');
        } catch (notificationError) {
          Logger.log('通知メール送信エラー: ' + notificationError.toString());
        }
      }

      return results;

    } catch (error) {
      Logger.log('見逃しエラー検知・復旧処理中にエラー: ' + error.toString());
      return createErrorResult(error.toString());
    }
  }

  /**
   * call_recordsシートで文字起こし内容にエラーが含まれているかチェック
   * @param {string} recordId - 録音ID
   * @return {Object} - {found: boolean, issue: string}
   */
  function checkForErrorInTranscription(recordId) {
    try {
      var sheet = ConfigManager.getCallRecordsSheet();
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var recordIdIndex = headers.indexOf(Constants.COLUMNS.CALL_RECORDS.RECORD_ID);
      var transcriptionIndex = headers.indexOf(Constants.COLUMNS.CALL_RECORDS.TRANSCRIPTION);

      if (recordIdIndex === -1 || transcriptionIndex === -1) {
        return { found: false, issue: Constants.ERROR_MESSAGES.REQUIRED_COLUMNS_NOT_FOUND };
      }

      // 該当レコードを検索
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row[recordIdIndex] === recordId) {
          var transcription = row[transcriptionIndex] || '';

          // エラーパターンをチェック
          for (var j = 0; j < Constants.ERROR_PATTERNS.length; j++) {
            var errorPattern = Constants.ERROR_PATTERNS[j];
            if (transcription.includes(errorPattern.pattern)) {
              return { found: true, issue: errorPattern.issue };
            }
          }

          return { found: false, issue: null };
        }
      }

      return { found: false, issue: Constants.ERROR_MESSAGES.RECORD_NOT_FOUND };

    } catch (error) {
      Logger.log('文字起こし内容エラーチェック中にエラー: ' + error.toString());
      return { found: false, issue: 'チェック処理エラー' };
    }
  }

  /**
   * レコードIDに基づいてファイルをエラーフォルダに移動
   * @param {string} recordId - 録音ID
   * @return {boolean} - ファイルが見つかって移動できた場合true
   */
  function moveFileToErrorFolder(recordId) {
    try {
      var config = ConfigManager.getConfig();
      var foldersToCheck = [
        { id: config.SOURCE_FOLDER_ID, name: '未処理フォルダ' },
        { id: config.PROCESSING_FOLDER_ID, name: '処理中フォルダ' },
        { id: config.COMPLETED_FOLDER_ID, name: '完了フォルダ' }
      ];

      // 各フォルダをチェック
      for (var i = 0; i < foldersToCheck.length; i++) {
        var folderInfo = foldersToCheck[i];
        if (!folderInfo.id) continue;

        try {
          var folder = DriveApp.getFolderById(folderInfo.id);
          var files = folder.getFiles();

          while (files.hasNext()) {
            var file = files.next();
            var fileName = file.getName();

            // ファイル名にレコードIDが含まれているかチェック
            if (fileName.indexOf(recordId) !== -1) {
              Logger.log('見逃しエラー対象ファイルを発見: ' + fileName + ' (' + folderInfo.name + ')');

              // エラーフォルダに移動
              FileMovementService.moveFileWithFallback(file, config.ERROR_FOLDER_ID);
              Logger.log('ファイルをエラーフォルダに移動しました: ' + fileName);
              return true;
            }
          }
        } catch (folderError) {
          Logger.log(folderInfo.name + 'のチェック中にエラー: ' + folderError.toString());
        }
      }

      Logger.log('対応するファイルが見つかりませんでした: record_id=' + recordId);
      return false;
    } catch (error) {
      Logger.log('ファイル移動中にエラー: ' + error.toString());
      return false;
    }
  }

  /**
   * 空の結果オブジェクトを作成
   * @return {Object} 空の結果オブジェクト
   */
  function createEmptyResult() {
    return {
      total: 0,
      recovered: 0,
      failed: 0,
      details: []
    };
  }

  /**
   * エラー結果オブジェクトを作成
   * @param {string} errorMessage - エラーメッセージ
   * @return {Object} エラー結果オブジェクト
   */
  function createErrorResult(errorMessage) {
    return {
      total: 0,
      recovered: 0,
      failed: 1,
      details: [{
        status: 'error',
        message: errorMessage
      }]
    };
  }

  /**
   * 見逃しエラー検知の統計情報を取得
   * @return {Object} 統計情報
   */
  function getDetectionStatistics() {
    try {
      var sheet = ConfigManager.getRecordingsSheet();
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var statusIndex = headers.indexOf(Constants.COLUMNS.RECORDINGS.STATUS_TRANSCRIPTION);

      if (statusIndex === -1) {
        return { error: 'ステータス列が見つかりません' };
      }

      var stats = {
        total: data.length - 1, // ヘッダー行を除く
        success: 0,
        error: 0,
        errorDetected: 0,
        pending: 0,
        other: 0
      };

      // ステータス別にカウント
      for (var i = 1; i < data.length; i++) {
        var status = data[i][statusIndex];

        if (status === Constants.STATUS.SUCCESS) {
          stats.success++;
        } else if (status && status.toString().indexOf('ERROR') === 0) {
          if (status === Constants.STATUS.ERROR_DETECTED) {
            stats.errorDetected++;
          } else {
            stats.error++;
          }
        } else if (status === Constants.STATUS.PENDING) {
          stats.pending++;
        } else {
          stats.other++;
        }
      }

      return stats;
    } catch (error) {
      Logger.log('統計情報取得中にエラー: ' + error.toString());
      return { error: error.toString() };
    }
  }

  // 公開メソッド
  return {
    detectAndRecoverPartialFailures: detectAndRecoverPartialFailures,
    checkForErrorInTranscription: checkForErrorInTranscription,
    moveFileToErrorFolder: moveFileToErrorFolder,
    getDetectionStatistics: getDetectionStatistics
  };
})(); 