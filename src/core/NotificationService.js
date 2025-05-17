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

      var body = dateStr + ' の処理結果サマリー\n\n';

      // 基本的な成功/エラーサマリー
      body += '基本サマリー:\n';
      body += '成功: ' + results.success + '件\n';
      body += 'エラー: ' + results.error + '件\n';
      body += '本日の処理件数合計: ' + (results.success + results.error) + '件\n\n';

      // status_fetchごとの件数
      if (results.fetchStatusCounts && Object.keys(results.fetchStatusCounts).length > 0) {
        body += 'ファイル取得ステータス別集計:\n';
        for (var status in results.fetchStatusCounts) {
          body += status + ': ' + results.fetchStatusCounts[status] + '件\n';
        }
        body += '\n';
      }

      // status_transcriptionごとの件数
      if (results.transcriptionStatusCounts && Object.keys(results.transcriptionStatusCounts).length > 0) {
        body += '文字起こしステータス別集計:\n';
        for (var status in results.transcriptionStatusCounts) {
          body += status + ': ' + results.transcriptionStatusCounts[status] + '件\n';
        }
        body += '\n';
      }

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
      Logger.log('メール送信中にエラー: ' + error.toString());
    }
  }

  // 公開メソッド
  return {
    sendDailyProcessingSummary: sendDailyProcessingSummary
  };
})();