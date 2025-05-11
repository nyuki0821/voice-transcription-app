/**
 * 通知サービスモジュール
 */
var NotificationService = (function () {
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
    sendDailyProcessingSummary: sendDailyProcessingSummary
  };
})();