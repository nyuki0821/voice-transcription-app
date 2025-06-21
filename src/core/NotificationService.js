/**
 * 通知サービスモジュール
 */
var NotificationService = (function () {
  /**
   * エラーの重要度を判断する
   * @param {string} errorMessage - エラーメッセージ
   * @return {boolean} - 重要なエラーの場合はtrue、それ以外はfalse
   */
  function isHighPriorityError(errorMessage) {
    if (!errorMessage) return true; // エラーメッセージがない場合は重要とみなす

    const errorStr = errorMessage.toString().toLowerCase();

    // クリティカルなエラーパターン
    const criticalPatterns = [
      'api key', 'authentication', 'quota exceeded', 'rate limit',
      'storage', 'permission denied', 'access denied',
      'critical', 'fatal', 'unauthorized', '401', '403',
      'insufficient', 'not found', 'timeout exceeded',
      'database', 'connection', 'network error'
    ];

    // 一時的なエラーパターン（通知不要）
    const temporaryPatterns = [
      'syntax error', 'json', 'unexpected end', 'parse error',
      'timeout', 'temporary', 'retry', 'network',
      'too many requests', '429', 'server error', '5',
      'information extraction', '情報抽出'
    ];

    // クリティカルパターンにマッチする場合は高優先度
    for (const pattern of criticalPatterns) {
      if (errorStr.includes(pattern)) return true;
    }

    // 一時的なエラーパターンにマッチする場合は低優先度
    for (const pattern of temporaryPatterns) {
      if (errorStr.includes(pattern)) return false;
    }

    // デフォルトは通知する（安全側に倒す）
    return true;
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
   * リアルタイム処理エラーをメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} results - 処理結果オブジェクト
   */
  function sendRealtimeErrorNotification(email, results) {
    if (!email || !results || results.error === 0) {
      return;
    }

    try {
      // 重要なエラーがあるか確認
      let hasHighPriorityError = false;
      let highPriorityErrorDetails = [];

      if (results.details && results.details.length > 0) {
        for (var i = 0; i < results.details.length; i++) {
          var detail = results.details[i];
          if (detail.status === 'error') {
            // エラーの重要度を判断
            if (isHighPriorityError(detail.message)) {
              hasHighPriorityError = true;
              highPriorityErrorDetails.push(detail);
            }
          }
        }
      }

      // 重要なエラーがない場合は通知しない
      if (!hasHighPriorityError) {
        Logger.log('重要度の低いエラーのみのため、通知をスキップします');
        return;
      }

      var subject = '【緊急】顧客会話自動文字起こしシステム - 処理エラー発生';

      var body = 'リアルタイム処理中に重要なエラーが発生しました\n\n';

      // エラーサマリー
      body += 'エラーサマリー:\n';
      body += '処理対象ファイル数: ' + results.total + '件\n';
      body += '成功: ' + results.success + '件\n';
      body += 'エラー: ' + results.error + '件\n';
      body += '重要なエラー: ' + highPriorityErrorDetails.length + '件\n';
      body += '発生時刻: ' + new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }) + '\n\n';

      // エラー詳細
      if (highPriorityErrorDetails.length > 0) {
        body += '重要なエラー詳細:\n';
        for (var i = 0; i < highPriorityErrorDetails.length; i++) {
          var detail = highPriorityErrorDetails[i];
          body += (i + 1) + '. ファイル: ' + detail.fileName + '\n';
          body += '   エラー内容: ' + detail.message + '\n\n';
        }
      }

      body += '対処方法:\n';
      body += '1. OpenAI APIクォータ制限の場合: https://platform.openai.com/usage で使用量を確認\n';
      body += '2. 一時的な問題の場合: 自動復旧機能により再処理されます\n';
      body += '3. 継続的な問題の場合: システム管理者にお問い合わせください\n\n';
      body += '注意: エラーファイルは自動的にエラーフォルダに移動され、復旧処理の対象となります。';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
      Logger.log('リアルタイムエラー通知メール送信中にエラー: ' + error.toString());
    }
  }

  /**
   * 見逃しエラー検知結果をメールで送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} results - 検知結果オブジェクト
   */
  function sendPartialFailureDetectionSummary(email, results) {
    if (!email || !results || results.recovered === 0) {
      return;
    }

    try {
      // 重要なエラーがあるか確認
      let hasHighPriorityError = false;
      let highPriorityErrorDetails = [];

      if (results.details && results.details.length > 0) {
        for (var i = 0; i < results.details.length; i++) {
          var detail = results.details[i];
          if (detail.status === 'recovered') {
            // エラーの重要度を判断
            if (isHighPriorityError(detail.message)) {
              hasHighPriorityError = true;
              highPriorityErrorDetails.push(detail);
            }
          }
        }
      }

      // 重要なエラーがない場合は通知しない
      if (!hasHighPriorityError) {
        Logger.log('重要度の低い見逃しエラーのみのため、通知をスキップします');
        return;
      }

      var subject = '顧客会話自動文字起こしシステム - 重要な見逃しエラー検知・復旧完了';

      var body = '重要な見逃しエラーの検知・復旧処理の結果をお知らせします\n\n';

      // 検知サマリー
      body += '検知サマリー:\n';
      body += '検査対象レコード数: ' + results.total + '件\n';
      body += '見逃しエラー検知: ' + results.recovered + '件\n';
      body += '重要な見逃しエラー: ' + highPriorityErrorDetails.length + '件\n';
      body += '復旧失敗: ' + results.failed + '件\n';
      body += '検知時刻: ' + new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }) + '\n\n';

      // 詳細情報
      if (highPriorityErrorDetails.length > 0) {
        body += '検知された重要な問題の詳細:\n';
        for (var i = 0; i < highPriorityErrorDetails.length; i++) {
          var detail = highPriorityErrorDetails[i];
          body += (i + 1) + '. Record ID: ' + detail.recordId + '\n';
          if (detail.issue) {
            body += '   問題の種類: ' + detail.issue + '\n';
          }
          body += '   復旧状況: ' + detail.message + '\n';
          if (detail.fileFound !== undefined) {
            body += '   ファイル移動: ' + (detail.fileFound ? '成功' : 'ファイル未発見') + '\n';
          }
          body += '\n';
        }
      }

      body += '復旧処理について:\n';
      body += '- 検知されたレコードのステータスは「ERROR_DETECTED」に更新されました\n';
      body += '- 対応するファイルはエラーフォルダに移動されました\n';
      body += '- これらのファイルは自動復旧機能により再処理されます\n\n';

      body += '注意事項:\n';
      body += '- OpenAI APIクォータ制限が原因の場合は、クォータ復旧後に自動的に再処理されます\n';
      body += '- 継続的に同じエラーが発生する場合は、システム設定の確認が必要です\n\n';

      body += '見逃しエラーとは:\n';
      body += 'ステータスが「SUCCESS」になっているが、実際の文字起こし内容にエラーメッセージが含まれているケースです。\n';
      body += 'これらは処理途中でエラーが発生したものの、ステータス更新が適切に行われなかった可能性があります。';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
      Logger.log('見逃しエラー検知結果メール送信中にエラー: ' + error.toString());
    }
  }

  // 公開メソッド
  return {
    sendDailyProcessingSummary: sendDailyProcessingSummary,
    sendRealtimeErrorNotification: sendRealtimeErrorNotification,
    sendPartialFailureDetectionSummary: sendPartialFailureDetectionSummary,
    isHighPriorityError: isHighPriorityError // テスト用に公開
  };
})();