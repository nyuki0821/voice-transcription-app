/**
 * スプレッドシート操作モジュール
 */
var SpreadsheetManager = (function () {
  /**
   * ファイル名から日付と時間情報を抽出する（日本時間への変換対応版）
   * @param {string} fileName - 音声ファイル名（例: call_recording_d4ea9100-35ca-49fa-9b10-45cb53d0309b_20250327042723.mp3）
   * @return {Object} - 抽出された日付と時間情報（JST変換済み）と元のrecord_id用文字列
   */
  function extractDateTimeFromFileName(fileName) {
    try {
      // 最初に zoom_call_ パターンをチェック (例: zoom_call_20250502011713_xxxx.mp3)
      var zoomRegex = /zoom_call_(\d{8})(\d{6})_/;
      var zoomMatch = fileName.match(zoomRegex);

      if (zoomMatch && zoomMatch[1] && zoomMatch[2]) {
        var dateStr = zoomMatch[1]; // 例: 20250502
        var timeStr = zoomMatch[2]; // 例: 011713

        var dateTimeStr = dateStr + timeStr; // 例: 20250502011713
        var originalDateTimeStr = dateTimeStr;

        // 日付部分を抽出（年月日）
        var year = parseInt(dateStr.substring(0, 4), 10);
        var month = parseInt(dateStr.substring(4, 6), 10);
        var day = parseInt(dateStr.substring(6, 8), 10);

        // 時間部分を抽出（時分秒）
        var hour = parseInt(timeStr.substring(0, 2), 10);
        var minute = parseInt(timeStr.substring(2, 4), 10);
        var second = parseInt(timeStr.substring(4, 6), 10);

        // UTC時間からJST（UTC+9）に変換
        hour += 9;

        // 日付の繰り上げ処理
        if (hour >= 24) {
          hour -= 24;

          // 日付を1日進める
          var date = new Date(year, month - 1, day); // JavaScriptの月は0始まり
          date.setDate(date.getDate() + 1);

          year = date.getFullYear();
          month = date.getMonth() + 1; // 0始まりを1始まりに戻す
          day = date.getDate();
        }

        // 桁数調整（1桁の場合は先頭に0を付ける）
        var formattedMonth = ('0' + month).slice(-2);
        var formattedDay = ('0' + day).slice(-2);
        var formattedHour = ('0' + hour).slice(-2);
        var formattedMinute = ('0' + minute).slice(-2);
        var formattedSecond = ('0' + second).slice(-2);

        // フォーマットされた日付と時間を返す
        return {
          formattedDate: year + '-' + formattedMonth + '-' + formattedDay,
          formattedTime: formattedHour + ':' + formattedMinute + ':' + formattedSecond,
          originalDateTime: originalDateTimeStr // record_id用に元の時間情報を保持
        };
      }

      // 従来のパターン: ファイル名から日時部分を抽出するための正規表現
      // パターン1: _YYYYMMDDHHMMSS. (直後がファイル拡張子)
      // パターン2: _YYYYMMDDHHMMSS_ (直後に別の情報がある場合)
      var regex = /_(\d{14})([_\.])/;
      var match = fileName.match(regex);

      if (match && match[1]) {
        var dateTimeStr = match[1]; // 例: 20250327042723

        // 元の日時文字列をrecord_id用に保存
        var originalDateTimeStr = dateTimeStr;

        // 日付部分を抽出（年月日）
        var year = parseInt(dateTimeStr.substring(0, 4), 10);
        var month = parseInt(dateTimeStr.substring(4, 6), 10);
        var day = parseInt(dateTimeStr.substring(6, 8), 10);

        // 時間部分を抽出（時分秒）
        var hour = parseInt(dateTimeStr.substring(8, 10), 10);
        var minute = parseInt(dateTimeStr.substring(10, 12), 10);
        var second = parseInt(dateTimeStr.substring(12, 14), 10);

        // UTC時間からJST（UTC+9）に変換
        hour += 9;

        // 日付の繰り上げ処理
        if (hour >= 24) {
          hour -= 24;

          // 日付を1日進める
          var date = new Date(year, month - 1, day); // JavaScriptの月は0始まり
          date.setDate(date.getDate() + 1);

          year = date.getFullYear();
          month = date.getMonth() + 1; // 0始まりを1始まりに戻す
          day = date.getDate();
        }

        // 桁数調整（1桁の場合は先頭に0を付ける）
        var formattedMonth = ('0' + month).slice(-2);
        var formattedDay = ('0' + day).slice(-2);
        var formattedHour = ('0' + hour).slice(-2);
        var formattedMinute = ('0' + minute).slice(-2);
        var formattedSecond = ('0' + second).slice(-2);

        // フォーマットされた日付と時間を返す
        return {
          formattedDate: year + '-' + formattedMonth + '-' + formattedDay,
          formattedTime: formattedHour + ':' + formattedMinute + ':' + formattedSecond,
          originalDateTime: originalDateTimeStr // record_id用に元の時間情報を保持
        };
      } else {
        // ファイル名から日時を抽出できない場合は現在時刻を使用
        var now = new Date();
        // 日付を手動でフォーマット
        var year = now.getFullYear();
        var month = ('0' + (now.getMonth() + 1)).slice(-2);
        var day = ('0' + now.getDate()).slice(-2);
        var hour = ('0' + now.getHours()).slice(-2);
        var minute = ('0' + now.getMinutes()).slice(-2);
        var second = ('0' + now.getSeconds()).slice(-2);

        // 現在時刻からrecord_id用の文字列を生成
        var originalDateTime = year + month + day + hour + minute + second;

        Logger.log('日時抽出失敗: ファイル名=' + fileName + ', デフォルト値を使用=' + originalDateTime);

        return {
          formattedDate: year + '-' + month + '-' + day,
          formattedTime: hour + ':' + minute + ':' + second,
          originalDateTime: originalDateTime // record_id用に元の時間情報を保持
        };
      }
    } catch (error) {
      // エラーが発生した場合でも現在時刻を返す
      var now = new Date();
      var year = now.getFullYear();
      var month = ('0' + (now.getMonth() + 1)).slice(-2);
      var day = ('0' + now.getDate()).slice(-2);
      var hour = ('0' + now.getHours()).slice(-2);
      var minute = ('0' + now.getMinutes()).slice(-2);
      var second = ('0' + now.getSeconds()).slice(-2);

      // 現在時刻からrecord_id用の文字列を生成
      var originalDateTime = year + month + day + hour + minute + second;

      return {
        formattedDate: year + '-' + month + '-' + day,
        formattedTime: hour + ':' + minute + ':' + second,
        originalDateTime: originalDateTime // record_id用に元の時間情報を保持
      };
    }
  }

  /**
   * 処理対象ファイルの一覧を取得する
   * @returns {Array} 未処理ファイルのリスト
   */
  function getUnprocessedFiles() {
    try {
      var folder = getSourceFolder();
      if (!folder) return [];

      var files = folder.getFiles();
      var unprocessed = [];

      while (files.hasNext()) {
        var file = files.next();
        // MP3ファイルのみを対象とする
        if (file.getName().toLowerCase().indexOf('.mp3') !== -1) {
          unprocessed.push(file);
        }
      }

      return unprocessed;
    } catch (e) {
      Logger.log('未処理ファイル取得エラー: ' + e.toString());
      return [];
    }
  }

  /**
   * 音声ファイルを処理中フォルダに移動
   * @param {DriveFile} file ドライブファイル
   * @returns {boolean} 成功したらtrue
   */
  function moveToProcessing(file) {
    try {
      var processingFolder = getProcessingFolder();
      if (!processingFolder) return false;

      // 元のフォルダから削除
      var parents = file.getParents();
      while (parents.hasNext()) {
        var parent = parents.next();
        parent.removeFile(file);
      }

      // 処理中フォルダに追加
      processingFolder.addFile(file);
      return true;
    } catch (e) {
      Logger.log('処理中フォルダ移動エラー: ' + e.toString());
      return false;
    }
  }

  /**
   * 音声ファイルを完了フォルダに移動
   * @param {DriveFile} file ドライブファイル
   * @returns {boolean} 成功したらtrue
   */
  function moveToCompleted(file) {
    try {
      var completedFolder = getCompletedFolder();
      if (!completedFolder) return false;

      // 元のフォルダから削除
      var parents = file.getParents();
      while (parents.hasNext()) {
        var parent = parents.next();
        parent.removeFile(file);
      }

      // 完了フォルダに追加
      completedFolder.addFile(file);
      return true;
    } catch (e) {
      Logger.log('完了フォルダ移動エラー: ' + e.toString());
      return false;
    }
  }

  /**
   * 音声ファイルをエラーフォルダに移動
   * @param {DriveFile} file ドライブファイル
   * @returns {boolean} 成功したらtrue
   */
  function moveToError(file) {
    try {
      var errorFolder = getErrorFolder();
      if (!errorFolder) return false;

      // 元のフォルダから削除
      var parents = file.getParents();
      while (parents.hasNext()) {
        var parent = parents.next();
        parent.removeFile(file);
      }

      // エラーフォルダに追加
      errorFolder.addFile(file);
      return true;
    } catch (e) {
      Logger.log('エラーフォルダ移動エラー: ' + e.toString());
      return false;
    }
  }

  // 各種フォルダの取得関数
  function getSourceFolder() {
    var folderId = EnvironmentConfig.get('SOURCE_FOLDER_ID', '');
    return folderId ? DriveApp.getFolderById(folderId) : null;
  }

  function getProcessingFolder() {
    var folderId = EnvironmentConfig.get('PROCESSING_FOLDER_ID', '');
    return folderId ? DriveApp.getFolderById(folderId) : null;
  }

  function getCompletedFolder() {
    var folderId = EnvironmentConfig.get('COMPLETED_FOLDER_ID', '');
    return folderId ? DriveApp.getFolderById(folderId) : null;
  }

  function getErrorFolder() {
    var folderId = EnvironmentConfig.get('ERROR_FOLDER_ID', '');
    return folderId ? DriveApp.getFolderById(folderId) : null;
  }

  // 公開API
  return {
    extractDateTimeFromFileName: extractDateTimeFromFileName,
    getUnprocessedFiles: getUnprocessedFiles,
    moveToProcessing: moveToProcessing,
    moveToCompleted: moveToCompleted,
    moveToError: moveToError,
    getSourceFolder: getSourceFolder,
    getProcessingFolder: getProcessingFolder,
    getCompletedFolder: getCompletedFolder,
    getErrorFolder: getErrorFolder
  };
})();