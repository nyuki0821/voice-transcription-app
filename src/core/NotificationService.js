/**
 * 通知サービスモジュール
 */
var NotificationService = (function () {
  /**
   * 処理結果サマリーをメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} results - 処理結果オブジェクト
   * @param {number} processingTime - 処理時間（秒）
   */
  function sendProcessingSummary(email, results, processingTime) {
    if (!email) {
      return;
    }

    try {
      var subject = '顧客会話自動文字起こしシステム - 処理結果サマリー';

      var body = '処理結果サマリー\n\n' +
        '処理日時: ' + new Date().toLocaleString() + '\n' +
        '処理時間: ' + processingTime + '秒\n\n' +
        '成功: ' + results.success + '件\n' +
        'エラー: ' + results.error + '件\n\n';

      if (results.details && results.details.length > 0) {
        body += '詳細:\n';
        for (var i = 0; i < results.details.length; i++) {
          var detail = results.details[i];
          body += '- ' + detail.fileName + ': ' + detail.status + ' (' + detail.message + ')\n';
        }
      }

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
    }
  }

  /**
   * エラー通知をメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {string} fileName - ファイル名
   * @param {string} errorMessage - エラーメッセージ
   */
  function sendErrorNotification(email, fileName, errorMessage) {
    if (!email) {
      return;
    }

    try {
      var subject = '顧客会話自動文字起こしシステム - エラー通知';

      var body = 'ファイル処理中にエラーが発生しました\n\n' +
        'ファイル名: ' + fileName + '\n' +
        'エラー内容: ' + errorMessage + '\n' +
        '発生日時: ' + new Date().toLocaleString() + '\n';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
    }
  }

  /**
   * 日次処理結果サマリーをメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} results - 処理結果オブジェクト
   * @param {string} dateStr - 日付文字列
   */
  function sendDailyProcessingSummary(email, results, dateStr) {
    if (!email) {
      return;
    }

    try {
      var subject = '顧客会話自動文字起こしシステム - ' + dateStr + ' 日次処理結果サマリー';

      var body = dateStr + ' の処理結果サマリー\n\n' +
        '成功: ' + results.success + '件\n' +
        'エラー: ' + results.error + '件\n\n';

      // ファイル名リストは含めない
      body += '本日の処理件数合計: ' + (results.success + results.error) + '件\n';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
    }
  }

  // 公開メソッド
  return {
    sendProcessingSummary: sendProcessingSummary,
    sendErrorNotification: sendErrorNotification,
    sendDailyProcessingSummary: sendDailyProcessingSummary
  };
})();