/**
 * 復旧処理サービス
 * 各種エラー・中断ファイルの復旧処理を統合管理
 */
var RecoveryService = (function () {

  /**
   * 処理中フォルダにある未完了ファイルを未処理フォルダに戻す
   * AppScriptの制限により中断されたファイルを復旧する
   */
  function recoverInterruptedFiles() {
    var startTime = new Date();
    Logger.log('中断ファイル復旧処理開始: ' + startTime);

    try {
      // 設定を取得
      var config = ConfigManager.getConfig();

      if (!config.PROCESSING_FOLDER_ID) {
        throw new Error('処理中フォルダIDが設定されていません');
      }

      if (!config.SOURCE_FOLDER_ID) {
        throw new Error('処理対象フォルダIDが設定されていません');
      }

      // 処理中フォルダのファイルを取得
      var processingFolder = ConfigManager.getFolder('PROCESSING');
      var files = processingFolder.getFiles();
      var interruptedFiles = [];

      while (files.hasNext()) {
        var file = files.next();

        // 音声ファイルのみを対象とする
        if (Constants.isAudioFile(file)) {
          interruptedFiles.push(file);
        }
      }

      // Recordingsシートで既にINTERRUPTEDステータスになっているファイルも復旧対象に追加
      var recordingsInterruptedFiles = getInterruptedFilesFromSheet();
      if (recordingsInterruptedFiles && recordingsInterruptedFiles.length > 0) {
        Logger.log('Recordingsシートから追加の中断ファイル数: ' + recordingsInterruptedFiles.length);
        interruptedFiles = interruptedFiles.concat(recordingsInterruptedFiles);
      }

      Logger.log('中断された可能性のあるファイル数: ' + interruptedFiles.length);

      if (interruptedFiles.length === 0) {
        return '中断されたファイルはありませんでした。';
      }

      // 復旧結果のトラッキング
      var results = FileMovementService.createResultObject();
      results.total = interruptedFiles.length;

      // ファイルを復旧
      for (var i = 0; i < interruptedFiles.length; i++) {
        var file = interruptedFiles[i];

        try {
          Logger.log('ファイル復旧処理: ' + file.getName());

          // ファイル名からrecord_idを抽出
          var recordId = Constants.extractRecordIdFromFileName(file.getName());

          if (!recordId) {
            // より広いパターンで再試行
            var fileNamePattern = /.*_([a-f0-9]{32}|[a-f0-9-]{36})\.mp3$/i;
            var matches = file.getName().match(fileNamePattern);
            recordId = matches && matches.length > 1 ? matches[1] : null;
          }

          // 状態を更新
          if (recordId) {
            // Recordingsシートの文字起こし状態を更新 (PENDING - 再処理対象として)
            updateTranscriptionStatusByRecordId(
              recordId,
              Constants.STATUS.PENDING,
              '',
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
            );
          } else {
            Logger.log('警告: record_idが特定できないため、Recordingsシートの状態更新をスキップします');
          }

          // ファイルを未処理フォルダに移動
          FileMovementService.moveFileWithFallback(file, config.SOURCE_FOLDER_ID);

          // 結果を記録
          FileMovementService.addSuccessResult(
            results,
            file.getName(),
            recordId || 'unknown',
            '未処理フォルダに復旧しました' + (recordId ? '' : '（record_id不明）')
          );

        } catch (error) {
          Logger.log('ファイル復旧エラー: ' + error.toString());

          FileMovementService.addFailureResult(
            results,
            file.getName(),
            null,
            error.toString()
          );
        }
      }

      // 処理結果のログ出力
      return FileMovementService.logProcessingResult(startTime, results, '中断ファイル復旧処理');

    } catch (error) {
      var errorMessage = '中断ファイル復旧処理でエラーが発生: ' + error.toString();
      Logger.log(errorMessage);
      return errorMessage;
    }
  }

  /**
   * エラーフォルダにあるファイルを未処理フォルダに戻す
   * エラーになったファイルを1回だけリトライする
   */
  function recoverErrorFiles() {
    var startTime = new Date();
    Logger.log('エラーファイル復旧処理開始: ' + startTime);

    try {
      // 設定を取得
      var config = ConfigManager.getConfig();

      if (!config.ERROR_FOLDER_ID) {
        throw new Error('エラーフォルダIDが設定されていません');
      }

      if (!config.SOURCE_FOLDER_ID) {
        throw new Error('処理対象フォルダIDが設定されていません');
      }

      // エラーフォルダのファイルを取得
      var errorFolder = ConfigManager.getFolder('ERROR');
      var files = errorFolder.getFiles();
      var errorFiles = [];

      while (files.hasNext()) {
        var file = files.next();

        // 音声ファイルのみを対象とする
        if (Constants.isAudioFile(file)) {
          errorFiles.push(file);

          // ファイルの説明を取得してリトライ済みかチェック
          var description = file.getDescription() || "";
          var hasRetried = description.indexOf(Constants.RETRY_MARKS.RETRIED) >= 0;

          if (hasRetried) {
            Logger.log('リトライ済みのファイルも含めて処理: ' + file.getName());
          }
        }
      }

      Logger.log('リトライ対象のエラーファイル数: ' + errorFiles.length);

      if (errorFiles.length === 0) {
        return 'リトライ対象のエラーファイルはありませんでした。';
      }

      // 復旧結果のトラッキング
      var results = FileMovementService.createResultObject();
      results.total = errorFiles.length;

      // ファイルを復旧
      for (var i = 0; i < errorFiles.length; i++) {
        var file = errorFiles[i];

        try {
          Logger.log('エラーファイル復旧処理: ' + file.getName());

          // ファイルからメタデータを取得
          var recordId = Constants.extractRecordIdFromFileName(file.getName());

          if (recordId) {
            // Recordingsシートの文字起こし状態を更新 (RETRY)
            updateTranscriptionStatusByRecordId(
              recordId,
              Constants.STATUS.RETRY,
              '',
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
            );
          }

          // ファイルを未処理フォルダに移動（リトライマーク付き）
          FileMovementService.moveFileWithFallback(
            file,
            config.SOURCE_FOLDER_ID,
            Constants.RETRY_MARKS.RETRIED
          );

          // 結果を記録
          FileMovementService.addSuccessResult(
            results,
            file.getName(),
            recordId,
            'エラーファイルを未処理フォルダに復旧しました（1回限りのリトライ）'
          );

        } catch (error) {
          Logger.log('エラーファイル復旧処理エラー: ' + error.toString());

          FileMovementService.addFailureResult(
            results,
            file.getName(),
            null,
            error.toString()
          );
        }
      }

      // 処理結果のログ出力
      return FileMovementService.logProcessingResult(startTime, results, 'エラーファイル復旧処理');

    } catch (error) {
      var errorMessage = 'エラーファイル復旧処理でエラーが発生: ' + error.toString();
      Logger.log(errorMessage);
      return errorMessage;
    }
  }

  /**
   * 特別復旧処理: PENDINGステータスのままになっている文字起こしを再処理
   */
  function resetPendingTranscriptions() {
    var startTime = new Date();
    Logger.log('PENDING状態の文字起こし復旧処理開始: ' + startTime);

    try {
      var sheet = ConfigManager.getRecordingsSheet();
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      var pendingRecords = [];

      // ヘッダー行をスキップして2行目から処理
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        // status_transcription（12列目）をチェック
        if (row[11] === Constants.STATUS.PENDING) {
          pendingRecords.push({
            rowIndex: i + 1, // スプレッドシートの行番号（1始まり）
            recordId: row[0], // record_id
            timestamp: row[1] // timestamp_recording
          });
        }
      }

      Logger.log('PENDING状態の文字起こしレコード数: ' + pendingRecords.length);

      if (pendingRecords.length === 0) {
        return 'PENDING状態の文字起こしレコードはありませんでした。';
      }

      // 処理対象の記録
      var results = FileMovementService.createResultObject();
      results.total = pendingRecords.length;

      // PENDINGレコードを処理
      for (var i = 0; i < pendingRecords.length; i++) {
        var record = pendingRecords[i];

        try {
          Logger.log('PENDING文字起こし処理: record_id=' + record.recordId);

          // タイムスタンプを更新して再処理をトリガー
          var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

          // 処理状態を一旦RESETにして再処理を促す
          updateTranscriptionStatusByRecordId(
            record.recordId,
            Constants.STATUS.RESET_PENDING,
            now,
            now
          );

          // ドライブ内でファイルを検索して移動
          var fileFound = findAndMoveFileToSource(record.recordId);

          FileMovementService.addSuccessResult(
            results,
            'record_' + record.recordId,
            record.recordId,
            fileFound ? 'ステータスをリセットし、ファイルを移動しました' : 'ファイルが見つからず、ステータスのみリセット'
          );

        } catch (error) {
          Logger.log('PENDING文字起こしリセットエラー: ' + error.toString());

          FileMovementService.addFailureResult(
            results,
            'record_' + record.recordId,
            record.recordId,
            error.toString()
          );
        }
      }

      // 処理結果のログ出力
      return FileMovementService.logProcessingResult(startTime, results, 'PENDING状態の文字起こしリセット処理');

    } catch (error) {
      var errorMessage = 'PENDING文字起こしリセット処理でエラー: ' + error.toString();
      Logger.log(errorMessage);
      return errorMessage;
    }
  }

  /**
   * 特別復旧処理: エラーフォルダに残っているファイルを一括で未処理フォルダに戻す
   * 通常のrecoverErrorFilesではスキップされている[RETRIED]マーク付きのファイルも対象
   */
  function forceRecoverAllErrorFiles() {
    var startTime = new Date();
    Logger.log('エラーフォルダ内ファイル強制復旧処理開始: ' + startTime);

    try {
      // 設定を取得
      var config = ConfigManager.getConfig();

      if (!config.ERROR_FOLDER_ID) {
        throw new Error('エラーフォルダIDが設定されていません');
      }

      if (!config.SOURCE_FOLDER_ID) {
        throw new Error('処理対象フォルダIDが設定されていません');
      }

      // エラーフォルダのファイルを取得（すべて対象）
      var errorFolder = ConfigManager.getFolder('ERROR');
      var files = errorFolder.getFiles();
      var errorFiles = [];

      while (files.hasNext()) {
        var file = files.next();

        // 音声ファイルのみを対象とする
        if (Constants.isAudioFile(file)) {
          errorFiles.push(file);
        }
      }

      Logger.log('強制復旧対象のエラーファイル数: ' + errorFiles.length);

      if (errorFiles.length === 0) {
        return '復旧対象のエラーファイルはありませんでした。';
      }

      // 復旧結果のトラッキング
      var results = FileMovementService.createResultObject();
      results.total = errorFiles.length;

      // ファイルを復旧
      for (var i = 0; i < errorFiles.length; i++) {
        var file = errorFiles[i];

        try {
          Logger.log('エラーファイル強制復旧処理: ' + file.getName());

          // ファイルからメタデータを取得
          var recordId = Constants.extractRecordIdFromFileName(file.getName());

          if (recordId) {
            // Recordingsシートの文字起こし状態を更新 (FORCE_RETRY)
            updateTranscriptionStatusByRecordId(
              recordId,
              Constants.STATUS.FORCE_RETRY,
              '',
              Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
            );
          } else {
            Logger.log('警告: record_idが特定できないため、Recordingsシートの状態更新をスキップします');
          }

          // 元の説明を保持し、強制リトライマークを追加
          var originalDescription = file.getDescription() || "";
          file.setDescription(originalDescription + " " + Constants.RETRY_MARKS.FORCE_RETRY);

          // ファイルを未処理フォルダに移動
          FileMovementService.moveFileWithFallback(file, config.SOURCE_FOLDER_ID);

          // 結果を記録
          FileMovementService.addSuccessResult(
            results,
            file.getName(),
            recordId,
            'エラーファイルを強制復旧しました'
          );

        } catch (error) {
          Logger.log('エラーファイル強制復旧処理エラー: ' + error.toString());

          FileMovementService.addFailureResult(
            results,
            file.getName(),
            null,
            error.toString()
          );
        }
      }

      // 処理結果のログ出力
      return FileMovementService.logProcessingResult(startTime, results, 'エラーファイル強制復旧処理');

    } catch (error) {
      var errorMessage = 'エラーファイル強制復旧処理でエラー: ' + error.toString();
      Logger.log(errorMessage);
      return errorMessage;
    }
  }

  /**
   * 全復旧処理を一括実行
   * 見逃しエラー検知 → PENDING復旧 → エラーファイル復旧の順で実行
   * @return {string} 実行結果メッセージ
   */
  function runFullRecoveryProcess() {
    var startTime = new Date();
    Logger.log('全復旧処理開始: ' + startTime);

    var results = [];

    try {
      // 1. 見逃しエラー検知・復旧
      Logger.log('1. 見逃しエラー検知・復旧を実行中...');
      var partialResult = detectAndRecoverPartialFailures();
      results.push('見逃しエラー検知: ' + JSON.stringify(partialResult));

      // 2. PENDING状態リセット
      Logger.log('2. PENDING状態リセットを実行中...');
      var pendingResult = resetPendingTranscriptions();
      results.push('PENDING復旧: ' + pendingResult);

      // 3. エラーファイル復旧
      Logger.log('3. エラーファイル復旧を実行中...');
      var errorResult = recoverErrorFiles();
      results.push('エラーファイル復旧: ' + errorResult);

      // 4. 中断ファイル復旧
      Logger.log('4. 中断ファイル復旧を実行中...');
      var interruptedResult = recoverInterruptedFiles();
      results.push('中断ファイル復旧: ' + interruptedResult);

      var endTime = new Date();
      var processingTime = (endTime - startTime) / 1000;

      var summary = '全復旧処理完了（処理時間: ' + processingTime + '秒）\n\n' + results.join('\n');
      Logger.log(summary);

      return summary;
    } catch (error) {
      var errorMessage = '全復旧処理中にエラー: ' + error.toString();
      Logger.log(errorMessage);
      return errorMessage;
    }
  }

  /**
   * レコードIDに基づいてファイルを検索し、未処理フォルダに移動
   * @param {string} recordId - 録音ID
   * @return {boolean} - ファイルが見つかって移動できた場合true
   */
  function findAndMoveFileToSource(recordId) {
    try {
      var config = ConfigManager.getConfig();
      var foldersToCheck = [
        { id: config.ERROR_FOLDER_ID, name: 'エラーフォルダ' },
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
              Logger.log('ファイルを発見: ' + fileName + ' (' + folderInfo.name + ')');

              // エラーフォルダや処理中フォルダにある場合は未処理フォルダに移動
              if (folderInfo.id === config.ERROR_FOLDER_ID ||
                folderInfo.id === config.PROCESSING_FOLDER_ID) {
                FileMovementService.moveFileWithFallback(file, config.SOURCE_FOLDER_ID);
                Logger.log('ファイルを未処理フォルダに移動しました: ' + fileName);
                return true;
              } else {
                Logger.log('ファイルは ' + folderInfo.name + ' にあります。移動は不要です。');
                return true;
              }
            }
          }
        } catch (folderError) {
          Logger.log(folderInfo.name + 'のチェック中にエラー: ' + folderError.toString());
        }
      }

      return false;
    } catch (error) {
      Logger.log('ファイル検索・移動中にエラー: ' + error.toString());
      return false;
    }
  }

  // 公開メソッド
  return {
    recoverInterruptedFiles: recoverInterruptedFiles,
    recoverErrorFiles: recoverErrorFiles,
    resetPendingTranscriptions: resetPendingTranscriptions,
    forceRecoverAllErrorFiles: forceRecoverAllErrorFiles,
    runFullRecoveryProcess: runFullRecoveryProcess,
    findAndMoveFileToSource: findAndMoveFileToSource
  };
})(); 