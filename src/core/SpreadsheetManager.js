/**
 * スプレッドシート操作モジュール
 */
var SpreadsheetManager = (function() {
  /**
   * ファイル名から日付と時間情報を抽出する（日本時間への変換対応版）
   * @param {string} fileName - 音声ファイル名（例: call_recording_d4ea9100-35ca-49fa-9b10-45cb53d0309b_20250327042723.mp3）
   * @return {Object} - 抽出された日付と時間情報（JST変換済み）と元のrecord_id用文字列
   */
  function extractDateTimeFromFileName(fileName) {
    try {
      // ファイル名から日時部分を抽出するための正規表現
      var regex = /_(\d{14})\./;
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
   * 処理ログをスプレッドシートに記録する
   * @param {string} fileName - 処理対象ファイル名
   * @param {string} status - 処理ステータス
   * @param {string} processStart - 処理開始時間（オプション）
   * @param {string} processEnd - 処理終了時間（オプション）
   */
  function logProcessing(fileName, status, processStart, processEnd) {
    try {
      // スプレッドシートIDを取得
      var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (!spreadsheetId) {
        return;
      }
      
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('処理ログ');
      
      if (!sheet) {
        // シートが存在しない場合は作成
        sheet = spreadsheet.insertSheet('処理ログ');
        // ヘッダー行を設定
        sheet.getRange(1, 1, 1, 5).setValues([
          ['日時', 'ファイル名', 'ステータス', 'process_start', 'process_end']
        ]);
      }
      
      // 最終行の次の行に追加
      var lastRow = sheet.getLastRow();
      var newRow = lastRow + 1;
      
      // 現在時刻を取得
      var now = new Date();
      // 日付を手動でフォーマット
      var year = now.getFullYear();
      var month = ('0' + (now.getMonth() + 1)).slice(-2);
      var day = ('0' + now.getDate()).slice(-2);
      var hour = ('0' + now.getHours()).slice(-2);
      var minute = ('0' + now.getMinutes()).slice(-2);
      var second = ('0' + now.getSeconds()).slice(-2);
      var formattedDateTime = year + '-' + month + '-' + day + ' ' + hour + ':' + minute + ':' + second;
      
      // 処理時間が指定されていない場合は現在時刻を使用
      var startTime = processStart || formattedDateTime;
      var endTime = processEnd || formattedDateTime;
      
      // データを配列として準備
      var rowData = [
        formattedDateTime,
        fileName,
        status,
        startTime,
        endTime
      ];
      
      // 行を追加
      sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    } catch (error) {
      // ログ記録のエラーは無視する
    }
  }
  
  // 公開メソッド
  return {
    logProcessing: logProcessing,
    extractDateTimeFromFileName: extractDateTimeFromFileName
  };
})();