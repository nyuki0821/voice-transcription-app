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

  /**
   * 部分的失敗検知結果をメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} results - 検知結果オブジェクト
   */
  function sendPartialFailureDetectionSummary(email, results) {
    if (!email || !results) {
      return;
    }

    try {
      var subject = '顧客会話自動文字起こしシステム - 部分的失敗検知結果';

      var body = '部分的失敗検知・復旧処理の結果\n\n';

      // 基本的なサマリー
      body += '検知サマリー:\n';
      body += '対象レコード数: ' + results.total + '件\n';
      body += '復旧成功: ' + results.recovered + '件\n';
      body += '復旧失敗: ' + results.failed + '件\n\n';

      // 詳細情報
      if (results.details && results.details.length > 0) {
        body += '詳細情報:\n';
        for (var i = 0; i < results.details.length; i++) {
          var detail = results.details[i];
          body += '- Record ID: ' + detail.recordId + '\n';
          body += '  ステータス: ' + detail.status + '\n';
          if (detail.issue) {
            body += '  問題: ' + detail.issue + '\n';
          }
          body += '  メッセージ: ' + detail.message + '\n';
          if (detail.fileFound !== undefined) {
            body += '  ファイル移動: ' + (detail.fileFound ? '成功' : '失敗') + '\n';
          }
          body += '\n';
        }
      }

      body += '\n注意: 部分的失敗が検知されたレコードは、エラーステータスに更新され、再処理の対象となります。\n';
      body += 'OpenAI APIクォータ制限が原因の場合は、クォータ復旧後に自動的に再処理されます。';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
      Logger.log('部分的失敗検知結果メール送信中にエラー: ' + error.toString());
    }
  }

  // 公開メソッド
  return {
    sendDailyProcessingSummary: sendDailyProcessingSummary,
    sendPartialFailureDetectionSummary: sendPartialFailureDetectionSummary
  };
})();