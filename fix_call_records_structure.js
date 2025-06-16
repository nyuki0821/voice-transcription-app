/**
 * call_recordsシートの構造を修正する関数
 * record_id列を先頭に追加します
 */
function fixCallRecordsStructure() {
  try {
    // 設定を取得
    var config = ConfigManager.getConfig();
    var spreadsheet = SpreadsheetApp.openById(config.PROCESSED_SHEET_ID);
    var sheet = spreadsheet.getSheetByName('call_records');
    
    if (!sheet) {
      return 'エラー: call_recordsシートが見つかりません';
    }
    
    // 現在のヘッダーを取得
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('現在のヘッダー: ' + currentHeaders.join(', '));
    
    // record_idが既に存在するかチェック
    if (currentHeaders[0] === 'record_id') {
      return 'record_id列は既に存在します';
    }
    
    // A列の前に新しい列を挿入
    sheet.insertColumnBefore(1);
    Logger.log('新しい列を先頭に挿入しました');
    
    // 新しいヘッダーを設定
    sheet.getRange(1, 1).setValue('record_id');
    Logger.log('record_idヘッダーを設定しました');
    
    // 既存のデータがある場合、record_idを補完する必要があるかチェック
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      Logger.log('既存データが' + (lastRow - 1) + '行あります');
      
      // 既存データのrecord_idを空文字で初期化
      var emptyIds = [];
      for (var i = 2; i <= lastRow; i++) {
        emptyIds.push(['']); // 空のrecord_id
      }
      
      if (emptyIds.length > 0) {
        sheet.getRange(2, 1, emptyIds.length, 1).setValues(emptyIds);
        Logger.log('既存データのrecord_id列を空文字で初期化しました');
      }
    }
    
    // 最終確認
    var newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('修正後のヘッダー: ' + newHeaders.join(', '));
    Logger.log('修正後の列数: ' + sheet.getLastColumn());
    
    return '修正完了: call_recordsシートが正しい構造（15列）になりました';
    
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    return 'エラーが発生しました: ' + error.toString();
  }
}

/**
 * 現在のcall_recordsシート構造を確認する関数
 */
function checkCallRecordsStructure() {
  try {
    var config = ConfigManager.getConfig();
    var spreadsheet = SpreadsheetApp.openById(config.PROCESSED_SHEET_ID);
    var sheet = spreadsheet.getSheetByName('call_records');
    
    if (!sheet) {
      return 'call_recordsシートが見つかりません';
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var expectedHeaders = [
      'record_id', 'call_date', 'call_time', 'sales_phone_number', 
      'sales_company', 'customer_phone_number', 'customer_name', 
      'call_status1', 'call_status2', 'reason_for_refusal', 
      'reason_for_refusal_category', 'reason_for_appointment', 
      'reason_for_appointment_category', 'summary', 'full_transcript'
    ];
    
    Logger.log('=== call_recordsシート構造チェック ===');
    Logger.log('現在の列数: ' + headers.length);
    Logger.log('期待される列数: ' + expectedHeaders.length);
    Logger.log('');
    
    Logger.log('現在のヘッダー:');
    headers.forEach((header, index) => {
      Logger.log('  列' + (index + 1) + ': ' + header);
    });
    
    Logger.log('');
    Logger.log('期待されるヘッダー:');
    expectedHeaders.forEach((header, index) => {
      var currentHeader = headers[index] || '(欠落)';
      var match = currentHeader === header ? '✓' : '✗';
      Logger.log('  列' + (index + 1) + ': ' + header + ' → 現在: ' + currentHeader + ' ' + match);
    });
    
    // 問題の特定
    if (headers[0] !== 'record_id') {
      Logger.log('\n⚠️ 問題: record_id列が先頭にありません');
      Logger.log('修正方法: fixCallRecordsStructure() を実行してください');
    }
    
    return '構造チェック完了（詳細はログを確認）';
    
  } catch (error) {
    return 'エラー: ' + error.toString();
  }
}