/**
 * ファイル移動サービス
 * 重複していたファイル移動処理を統一管理
 */
var FileMovementService = (function () {

  /**
   * ファイルを指定フォルダに移動（フォールバック機能付き）
   * @param {File} file - 移動対象ファイル
   * @param {string} targetFolderId - 移動先フォルダID
   * @param {string} [description] - ファイルに追加する説明
   * @return {boolean} - 移動成功時true
   */
  function moveFileWithFallback(file, targetFolderId, description) {
    if (!file) {
      throw new Error('移動対象ファイルが指定されていません');
    }

    if (!targetFolderId) {
      throw new Error('移動先フォルダIDが指定されていません');
    }

    try {
      var fileName = file.getName();
      var originalDescription = file.getDescription() || "";

      // 説明を更新（指定されている場合）
      if (description) {
        file.setDescription(originalDescription + " " + description);
      }

      // 通常の移動を試行
      try {
        FileProcessor.moveFileToFolder(file, targetFolderId);
        Logger.log('ファイル移動成功: ' + fileName);
        return true;
      } catch (moveError) {
        Logger.log('通常の移動方法で失敗: ' + moveError.toString());

        // 代替移動方法を試行（コピー → 削除）
        return performFallbackMove(file, targetFolderId, originalDescription, description);
      }

    } catch (error) {
      throw new Error('ファイル移動処理中にエラー: ' + error.toString());
    }
  }

  /**
   * 代替移動方法（コピー → 削除）
   * @param {File} file - 移動対象ファイル
   * @param {string} targetFolderId - 移動先フォルダID
   * @param {string} originalDescription - 元の説明
   * @param {string} [additionalDescription] - 追加する説明
   * @return {boolean} - 移動成功時true
   */
  function performFallbackMove(file, targetFolderId, originalDescription, additionalDescription) {
    try {
      Logger.log('代替移動方法を試行: コピー → 削除');

      var targetFolder = DriveApp.getFolderById(targetFolderId);
      var fileName = file.getName();
      var fileBlob = file.getBlob();

      // ファイルをコピー
      var copiedFile = targetFolder.createFile(fileBlob);
      copiedFile.setName(fileName);

      // 説明を設定
      var finalDescription = originalDescription;
      if (additionalDescription) {
        finalDescription += " " + additionalDescription;
      }
      finalDescription += "[COPY_RECOVERED]";
      copiedFile.setDescription(finalDescription);

      // 元のファイルを削除試行（失敗してもOK）
      try {
        file.setTrashed(true);
        Logger.log('元ファイルの削除成功');
      } catch (trashError) {
        Logger.log('元ファイルの削除に失敗（コピーは成功）: ' + trashError.toString());
      }

      Logger.log('代替方法でファイル移動完了: ' + fileName);
      return true;

    } catch (copyError) {
      throw new Error('代替移動方法でも失敗: ' + copyError.toString());
    }
  }

  /**
   * 処理結果のログを統一形式で出力
   * @param {Date} startTime - 処理開始時刻
   * @param {Object} results - 処理結果オブジェクト
   * @param {string} processName - 処理名
   * @return {string} - サマリー文字列
   */
  function logProcessingResult(startTime, results, processName) {
    var endTime = new Date();
    var processingTime = (endTime - startTime) / 1000; // 秒単位

    var summary = processName + '完了: ' +
      '対象=' + results.total + '件, ' +
      '成功=' + (results.recovered || results.success || 0) + '件, ' +
      '失敗=' + (results.failed || results.error || 0) + '件';

    // 追加情報があれば含める
    if (results.recordIdFound !== undefined) {
      summary += ', ID特定=' + results.recordIdFound + '件';
    }
    if (results.recordIdNotFound !== undefined) {
      summary += ', ID不明=' + results.recordIdNotFound + '件';
    }

    summary += ', 処理時間=' + processingTime + '秒';

    Logger.log(summary);
    return summary;
  }

  /**
   * 標準的な処理結果オブジェクトを作成
   * @return {Object} - 初期化された結果オブジェクト
   */
  function createResultObject() {
    return {
      total: 0,
      recovered: 0,
      failed: 0,
      details: []
    };
  }

  /**
   * 処理結果に成功ケースを追加
   * @param {Object} results - 結果オブジェクト
   * @param {string} fileName - ファイル名
   * @param {string} [recordId] - レコードID
   * @param {string} message - 成功メッセージ
   */
  function addSuccessResult(results, fileName, recordId, message) {
    results.recovered++;
    results.details.push({
      fileName: fileName,
      recordId: recordId || 'unknown',
      status: 'recovered',
      message: message
    });
  }

  /**
   * 処理結果に失敗ケースを追加
   * @param {Object} results - 結果オブジェクト
   * @param {string} fileName - ファイル名
   * @param {string} [recordId] - レコードID
   * @param {string} errorMessage - エラーメッセージ
   */
  function addFailureResult(results, fileName, recordId, errorMessage) {
    results.failed++;
    results.details.push({
      fileName: fileName,
      recordId: recordId || 'unknown',
      status: 'error',
      message: errorMessage
    });
  }

  // 公開メソッド
  return {
    moveFileWithFallback: moveFileWithFallback,
    logProcessingResult: logProcessingResult,
    createResultObject: createResultObject,
    addSuccessResult: addSuccessResult,
    addFailureResult: addFailureResult
  };
})(); 