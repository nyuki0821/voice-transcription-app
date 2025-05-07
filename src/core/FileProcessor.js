/**
 * ファイル処理モジュール（共有ドライブ対応版）
 */
var FileProcessor = (function () {
  /**
   * 未処理のファイルを取得する
   * @param {string} sourceFolderId - 処理対象フォルダID
   * @param {number} maxFiles - 最大取得ファイル数
   * @return {Array} - 未処理ファイルの配列
   */
  function getUnprocessedFiles(sourceFolderId, maxFiles) {
    if (!sourceFolderId) {
      throw new Error('処理対象フォルダIDが指定されていません');
    }

    try {
      var sourceFolder = DriveApp.getFolderById(sourceFolderId);
      var files = sourceFolder.getFiles();
      var fileArray = [];

      // 全ての音声ファイルを取得
      while (files.hasNext()) {
        var file = files.next();
        var mimeType = file.getMimeType() || "";

        // 音声ファイルのみを対象とする
        if (mimeType.indexOf('audio/') === 0 ||
          mimeType === 'application/octet-stream') { // 一部の音声ファイルはoctet-streamとして認識される場合がある
          fileArray.push(file);
        }
      }

      // Recordingsシートを取得して情報を関連付ける
      var enrichedFiles = getFilesWithRecordingsInfo(fileArray);

      // timestamp_fetchが古い順にソートする
      enrichedFiles.sort(function (a, b) {
        // timestamp_fetchが存在しない場合は最後に配置
        if (!a.metadata || !a.metadata.timestamp_fetch) return 1;
        if (!b.metadata || !b.metadata.timestamp_fetch) return -1;

        // 両方存在する場合はtimestamp_fetchの古い順にソート
        return new Date(a.metadata.timestamp_fetch) - new Date(b.metadata.timestamp_fetch);
      });

      // 最大ファイル数に制限
      var result = enrichedFiles.slice(0, maxFiles).map(function (item) {
        return item.file;
      });

      Logger.log('未処理ファイル: 取得数=' + result.length + '件 (timestamp_fetch昇順でソート済み)');
      return result;
    } catch (error) {
      throw new Error('未処理ファイルの取得中にエラー: ' + error.toString());
    }
  }

  /**
   * ファイル配列にRecordingsシートの情報を関連付ける
   * @param {Array} files - ファイル配列
   * @return {Array} - メタデータ付きファイル配列
   */
  function getFilesWithRecordingsInfo(files) {
    try {
      // Recordingsシートを取得
      var spreadsheetId = EnvironmentConfig.get('RECORDINGS_SHEET_ID', '');
      if (!spreadsheetId) {
        Logger.log('RECORDINGS_SHEET_ID が設定されていません。ファイル名順でソートします。');
        return files.map(function (file) {
          return { file: file, metadata: null };
        });
      }

      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('Recordings');
      if (!sheet) {
        Logger.log('Recordingsシートが見つかりません。ファイル名順でソートします。');
        return files.map(function (file) {
          return { file: file, metadata: null };
        });
      }

      // Recordingsシートのデータを取得
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();

      // ヘッダー行が少なくとも存在する必要がある
      if (values.length <= 1) {
        Logger.log('Recordingsシートにデータが存在しません。ファイル名順でソートします。');
        return files.map(function (file) {
          return { file: file, metadata: null };
        });
      }

      // ファイル名からrecord_idを抽出する正規表現パターン
      var recordIdPattern = /zoom_call_\d+_(\w+)\.mp3/;

      // 各ファイルにRecordingsシートの情報を関連付ける
      return files.map(function (file) {
        var fileName = file.getName();
        var metadata = null;

        // ファイル名からrecord_idを抽出
        var match = fileName.match(recordIdPattern);
        var recordId = match ? match[1] : null;

        if (recordId) {
          // Recordingsシートでrecord_idに一致する行を検索
          for (var i = 1; i < values.length; i++) {
            var row = values[i];
            if (row[0] === recordId) { // record_idは1列目
              metadata = {
                record_id: row[0],
                timestamp_recording: row[1],
                timestamp_fetch: row[8], // timestamp_fetchは9列目
                status_fetch: row[9],
                timestamp_transcription: row[10],
                status_transcription: row[11]
              };
              break;
            }
          }
        }

        return {
          file: file,
          metadata: metadata
        };
      });
    } catch (error) {
      Logger.log('Recordingsシートの読み取り中にエラー: ' + error.toString());
      // エラーが発生した場合、元のファイル配列をそのまま返す
      return files.map(function (file) {
        return { file: file, metadata: null };
      });
    }
  }

  /**
   * ファイルを指定フォルダに移動する（共有ドライブ対応版）
   * @param {File} file - 移動対象ファイル
   * @param {string} targetFolderId - 移動先フォルダID
   */
  function moveFileToFolder(file, targetFolderId) {
    if (!file) {
      throw new Error('移動対象ファイルが指定されていません');
    }

    if (!targetFolderId) {
      throw new Error('移動先フォルダIDが指定されていません');
    }

    try {
      var targetFolder = DriveApp.getFolderById(targetFolderId);

      // 共有ドライブ対応：moveTo()メソッドを使用
      file.moveTo(targetFolder);
    } catch (error) {
      throw new Error('ファイルの移動中にエラー: ' + error.toString());
    }
  }

  /**
   * ファイルを指定フォルダにコピーする（代替手段としてのコピー処理）
   * @param {File} file - コピー対象ファイル
   * @param {string} targetFolderId - コピー先フォルダID
   * @return {File} - コピーされたファイル
   */
  function copyFileToFolder(file, targetFolderId) {
    if (!file) {
      throw new Error('コピー対象ファイルが指定されていません');
    }

    if (!targetFolderId) {
      throw new Error('コピー先フォルダIDが指定されていません');
    }

    try {
      var targetFolder = DriveApp.getFolderById(targetFolderId);
      var fileName = file.getName();
      var fileBlob = file.getBlob();

      // ファイルをコピー先フォルダにコピー
      var copiedFile = targetFolder.createFile(fileBlob);
      copiedFile.setName(fileName);

      return copiedFile;
    } catch (error) {
      throw new Error('ファイルのコピー中にエラー: ' + error.toString());
    }
  }

  // 公開メソッド
  return {
    getUnprocessedFiles: getUnprocessedFiles,
    moveFileToFolder: moveFileToFolder,
    copyFileToFolder: copyFileToFolder
  };
})();