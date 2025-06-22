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
   * 詳細なエラーログを記録する
   * @param {Object} errorInfo - エラー情報
   * @param {string} processType - 処理タイプ
   */
  function logDetailedError(errorInfo, processType) {
    try {
      // 基本的なエラーログ
      var logMessage = '[ERROR] ' + processType + ': ' + errorInfo.message;
      Logger.log(logMessage);

      // 詳細なコンテキスト情報をログに記録
      var contextLog = {
        timestamp: new Date().toISOString(),
        processType: processType,
        errorCode: errorInfo.errorCode || 'UNKNOWN',
        message: errorInfo.message,
        context: errorInfo
      };

      // システム状態情報を追加
      contextLog.systemState = {
        memoryUsage: 'N/A (GAS environment)',
        executionTime: new Date().getTime(),
        triggerSource: 'Scheduled or Manual',
        environment: 'Google Apps Script'
      };

      // 処理タイプ別の追加情報
      if (processType === 'Zoom録音取得') {
        contextLog.apiEndpoints = [
          '/phone/recordings',
          '/phone/recording/download'
        ];
        contextLog.dependencies = [
          'Zoom Phone API',
          'Google Drive API'
        ];
      } else if (processType === '文字起こし') {
        contextLog.apiEndpoints = [
          'https://api.openai.com/v1/audio/transcriptions',
          'https://api.openai.com/v1/chat/completions'
        ];
        contextLog.dependencies = [
          'OpenAI Whisper API',
          'OpenAI GPT API'
        ];
      }

      // 構造化ログとして記録
      Logger.log('詳細エラーログ: ' + JSON.stringify(contextLog, null, 2));

      // エラーパターンの分析結果も記録
      var errorAnalysis = analyzeErrorPattern(errorInfo.message);
      Logger.log('エラー分析結果: ' + JSON.stringify(errorAnalysis, null, 2));

    } catch (logError) {
      Logger.log('詳細エラーログの記録中にエラー: ' + logError.toString());
    }
  }

  /**
   * エラーパターンを分析する
   * @param {string} errorMessage - エラーメッセージ
   * @return {Object} 分析結果
   */
  function analyzeErrorPattern(errorMessage) {
    var analysis = {
      category: 'UNKNOWN',
      severity: 'MEDIUM',
      possibleCauses: [],
      suggestedActions: [],
      isRetryable: false,
      estimatedRecoveryTime: 'Unknown'
    };

    if (!errorMessage) return analysis;

    var errorStr = errorMessage.toString().toLowerCase();

    // API関連エラー
    if (errorStr.includes('api') || errorStr.includes('quota') || errorStr.includes('rate limit')) {
      analysis.category = 'API_ERROR';
      analysis.severity = 'HIGH';
      analysis.possibleCauses = [
        'APIクォータ制限',
        'APIキーの無効化',
        'レート制限超過',
        'API仕様変更'
      ];
      analysis.suggestedActions = [
        'API使用量を確認する',
        'APIキーを更新する',
        '処理間隔を調整する',
        'API仕様を確認する'
      ];
      analysis.isRetryable = true;
      analysis.estimatedRecoveryTime = '1-24時間';
    }
    // ネットワーク関連エラー
    else if (errorStr.includes('network') || errorStr.includes('timeout') || errorStr.includes('connection')) {
      analysis.category = 'NETWORK_ERROR';
      analysis.severity = 'MEDIUM';
      analysis.possibleCauses = [
        'ネットワーク接続不安定',
        'タイムアウト設定が短い',
        'プロキシ設定問題',
        'DNS解決問題'
      ];
      analysis.suggestedActions = [
        'ネットワーク接続を確認する',
        'タイムアウト設定を調整する',
        'リトライ処理を実行する'
      ];
      analysis.isRetryable = true;
      analysis.estimatedRecoveryTime = '数分-1時間';
    }
    // ファイル関連エラー
    else if (errorStr.includes('file') || errorStr.includes('download') || errorStr.includes('upload')) {
      analysis.category = 'FILE_ERROR';
      analysis.severity = 'MEDIUM';
      analysis.possibleCauses = [
        'ファイルサイズ制限超過',
        'ファイル形式不正',
        'ストレージ容量不足',
        'アクセス権限不足'
      ];
      analysis.suggestedActions = [
        'ファイルサイズを確認する',
        'ファイル形式を確認する',
        'ストレージ容量を確認する',
        'アクセス権限を確認する'
      ];
      analysis.isRetryable = false;
      analysis.estimatedRecoveryTime = '手動対応が必要';
    }
    // 認証関連エラー
    else if (errorStr.includes('auth') || errorStr.includes('unauthorized') || errorStr.includes('forbidden')) {
      analysis.category = 'AUTH_ERROR';
      analysis.severity = 'HIGH';
      analysis.possibleCauses = [
        'トークンの期限切れ',
        'APIキーの無効化',
        'アクセス権限不足',
        '認証設定の変更'
      ];
      analysis.suggestedActions = [
        'トークンを更新する',
        'APIキーを確認する',
        'アクセス権限を確認する',
        '認証設定を見直す'
      ];
      analysis.isRetryable = false;
      analysis.estimatedRecoveryTime = '手動対応が必要';
    }
    // データ処理関連エラー
    else if (errorStr.includes('parse') || errorStr.includes('json') || errorStr.includes('format')) {
      analysis.category = 'DATA_ERROR';
      analysis.severity = 'LOW';
      analysis.possibleCauses = [
        'データ形式の不正',
        'API仕様変更',
        'データ破損',
        'エンコーディング問題'
      ];
      analysis.suggestedActions = [
        'データ形式を確認する',
        'API仕様を確認する',
        'データの再取得を試す'
      ];
      analysis.isRetryable = true;
      analysis.estimatedRecoveryTime = '数分-1時間';
    }

    return analysis;
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

      // 全体の処理件数を計算
      var totalFetchCount = 0;
      var totalTranscriptionCount = 0;

      // fetch_statusごとの件数
      if (results.fetchStatusCounts && Object.keys(results.fetchStatusCounts).length > 0) {
        body += 'ファイル取得ステータス別集計 (fetch_status):\n';
        for (var status in results.fetchStatusCounts) {
          var count = results.fetchStatusCounts[status];
          body += '  ' + status + ': ' + count + '件\n';
          totalFetchCount += count;
        }
        body += '  合計: ' + totalFetchCount + '件\n\n';
      } else {
        body += 'ファイル取得ステータス別集計 (fetch_status):\n';
        body += '  本日の処理対象レコードはありませんでした\n\n';
      }

      // transcription_statusごとの件数
      if (results.transcriptionStatusCounts && Object.keys(results.transcriptionStatusCounts).length > 0) {
        body += '文字起こしステータス別集計 (transcription_status):\n';
        for (var status in results.transcriptionStatusCounts) {
          var count = results.transcriptionStatusCounts[status];
          body += '  ' + status + ': ' + count + '件\n';
          totalTranscriptionCount += count;
        }
        body += '  合計: ' + totalTranscriptionCount + '件\n\n';
      } else {
        body += '文字起こしステータス別集計 (transcription_status):\n';
        body += '  本日の処理対象レコードはありませんでした\n\n';
      }

      // 注意事項を追加
      body += '注意事項:\n';
      body += '- このサマリーは本日(' + dateStr + ')のRecordingsシートのレコードを集計したものです\n';
      body += '- fetch_statusは録音ファイルの取得状況を示します\n';
      body += '- transcription_statusは文字起こし処理の状況を示します\n';
      body += '- エラーが発生している場合は、別途エラー通知メールが送信されます';

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
   * 統一フォーマットのエラー通知メールを送信する
   * @param {string} email - 送信先メールアドレス
   * @param {Object} errorDetails - エラー詳細オブジェクト
   */
  function sendUnifiedErrorNotification(email, errorDetails) {
    if (!email || !errorDetails) {
      return;
    }

    try {
      // エラーの重要度を判断
      var hasHighPriorityError = false;
      var highPriorityErrors = [];

      if (errorDetails.errors && errorDetails.errors.length > 0) {
        for (var i = 0; i < errorDetails.errors.length; i++) {
          var error = errorDetails.errors[i];
          if (isHighPriorityError(error.message)) {
            hasHighPriorityError = true;
            highPriorityErrors.push(error);
          }
        }
      }

      // 重要なエラーがない場合は通知しない
      if (!hasHighPriorityError) {
        Logger.log('重要度の低いエラーのみのため、通知をスキップします: ' + errorDetails.processType);
        return;
      }

      var subject = '【緊急】顧客会話自動文字起こしシステム - ' + errorDetails.processType + 'エラー発生';

      var body = errorDetails.processType + '処理中に重要なエラーが発生しました\n\n';

      // エラーサマリー
      body += 'エラーサマリー:\n';
      body += '処理タイプ: ' + errorDetails.processType + '\n';
      body += '処理対象数: ' + (errorDetails.total || 0) + '件\n';
      body += '成功: ' + (errorDetails.success || 0) + '件\n';
      body += 'エラー: ' + (errorDetails.error || 0) + '件\n';
      body += '重要なエラー: ' + highPriorityErrors.length + '件\n';
      body += '発生時刻: ' + new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }) + '\n\n';

      // エラー詳細
      if (highPriorityErrors.length > 0) {
        body += '重要なエラー詳細:\n';
        for (var i = 0; i < highPriorityErrors.length; i++) {
          var error = highPriorityErrors[i];
          body += (i + 1) + '. ';

          if (error.recordId) {
            body += 'Record ID: ' + error.recordId + '\n';
          }
          if (error.fileName) {
            body += '   ファイル名: ' + error.fileName + '\n';
          }
          if (error.downloadUrl) {
            body += '   ダウンロードURL: ' + error.downloadUrl.substring(0, 50) + '...\n';
          }
          if (error.phoneNumber) {
            body += '   電話番号: ' + error.phoneNumber + '\n';
          }
          if (error.duration) {
            body += '   通話時間: ' + error.duration + '秒\n';
          }

          body += '   エラー内容: ' + error.message + '\n';

          if (error.errorCode) {
            body += '   エラーコード: ' + error.errorCode + '\n';
          }
          if (error.timestamp) {
            body += '   発生時刻: ' + error.timestamp + '\n';
          }

          body += '\n';
        }
      }

      // 処理タイプ別の対処方法
      body += '対処方法:\n';
      if (errorDetails.processType === 'Zoom録音取得') {
        body += '1. Zoom APIトークンの有効性を確認してください\n';
        body += '2. Zoom APIクォータ制限の場合: Zoom Marketplaceで使用量を確認\n';
        body += '3. ダウンロードURLの有効期限切れの場合: 新しいURLを取得してください\n';
        body += '4. ネットワーク接続を確認してください\n';
      } else if (errorDetails.processType === '文字起こし') {
        body += '1. OpenAI APIクォータ制限の場合: https://platform.openai.com/usage で使用量を確認\n';
        body += '2. ファイルサイズが大きい場合: 自動分割処理が実行されます\n';
        body += '3. 音声ファイルの形式を確認してください\n';
        body += '4. OpenAI APIキーの有効性を確認してください\n';
      } else {
        body += '1. システムログを確認してください\n';
        body += '2. 一時的な問題の場合: 自動復旧機能により再処理されます\n';
        body += '3. 継続的な問題の場合: システム管理者にお問い合わせください\n';
      }

      body += '\n自動復旧について:\n';
      body += '- エラーが発生したレコードは自動的にエラーステータスに更新されます\n';
      body += '- 該当ファイルはエラーフォルダに移動され、復旧処理の対象となります\n';
      body += '- 一時的なエラーの場合、自動復旧機能により再処理されます\n\n';

      // LLM用の構造化ログ情報を追加
      body += '=== LLM解析用構造化ログ ===\n';
      body += '以下の情報をLLMに渡して問題解決を支援できます:\n\n';

      var structuredLog = createStructuredErrorLog(errorDetails, highPriorityErrors);
      body += JSON.stringify(structuredLog, null, 2) + '\n\n';

      body += '=== システムコンテキスト ===\n';
      body += '- システム名: 顧客会話自動文字起こしシステム\n';
      body += '- 実行環境: Google Apps Script\n';
      body += '- 処理フロー: Zoom録音取得 → 文字起こし → 情報抽出 → スプレッドシート保存\n';
      body += '- 使用API: Zoom Phone API, OpenAI Whisper API, OpenAI GPT API\n';
      body += '- ストレージ: Google Drive, Google Sheets\n\n';

      body += '注意: このメールは重要なエラーのみを通知しています。\n';
      body += '軽微なエラーや一時的なエラーは自動的にフィルタリングされています。';

      // メール送信
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      // エラーは無視
      Logger.log('統一エラー通知メール送信中にエラー: ' + error.toString());
    }
  }

  /**
   * LLM解析用の構造化エラーログを作成
   * @param {Object} errorDetails - エラー詳細
   * @param {Array} highPriorityErrors - 重要なエラーリスト
   * @return {Object} 構造化ログ
   */
  function createStructuredErrorLog(errorDetails, highPriorityErrors) {
    var structuredLog = {
      metadata: {
        systemName: '顧客会話自動文字起こしシステム',
        timestamp: new Date().toISOString(),
        processType: errorDetails.processType,
        severity: 'HIGH',
        environment: 'Google Apps Script'
      },
      summary: {
        totalProcessed: errorDetails.total || 0,
        successCount: errorDetails.success || 0,
        errorCount: errorDetails.error || 0,
        criticalErrorCount: highPriorityErrors.length
      },
      errors: [],
      systemState: {
        availableAPIs: ['Zoom Phone API', 'OpenAI Whisper API', 'OpenAI GPT API'],
        storageServices: ['Google Drive', 'Google Sheets'],
        processingPipeline: [
          'Zoom録音取得',
          'Google Driveへの保存',
          'Whisper文字起こし',
          'GPT情報抽出',
          'スプレッドシート保存'
        ]
      },
      troubleshootingContext: {
        commonCauses: [],
        suggestedActions: [],
        relatedComponents: []
      }
    };

    // エラー詳細を構造化
    for (var i = 0; i < highPriorityErrors.length; i++) {
      var error = highPriorityErrors[i];
      var structuredError = {
        errorId: 'ERR_' + new Date().getTime() + '_' + i,
        errorCode: error.errorCode || 'UNKNOWN',
        message: error.message,
        timestamp: error.timestamp,
        context: {}
      };

      // コンテキスト情報を追加
      if (error.recordId) structuredError.context.recordId = error.recordId;
      if (error.fileName) structuredError.context.fileName = error.fileName;
      if (error.downloadUrl) structuredError.context.downloadUrl = error.downloadUrl;
      if (error.phoneNumber) structuredError.context.phoneNumber = error.phoneNumber;
      if (error.duration) structuredError.context.duration = error.duration;

      // エラータイプ別の追加情報
      if (errorDetails.processType === 'Zoom録音取得') {
        structuredError.component = 'ZoomPhoneProcessor';
        structuredError.troubleshootingContext = {
          possibleCauses: [
            'Zoom APIトークンの期限切れ',
            'APIクォータ制限',
            'ダウンロードURLの無効化',
            'ネットワーク接続問題',
            'Google Driveストレージ容量不足'
          ],
          checkpoints: [
            'Zoom APIトークンの有効性',
            'APIレスポンスコード',
            'ダウンロードURL形式',
            'Google Driveアクセス権限'
          ]
        };
      } else if (errorDetails.processType === '文字起こし') {
        structuredError.component = 'TranscriptionService';
        structuredError.troubleshootingContext = {
          possibleCauses: [
            'OpenAI APIクォータ制限',
            'ファイルサイズ制限超過',
            '音声ファイル形式不正',
            'APIキーの無効化',
            'ネットワークタイムアウト'
          ],
          checkpoints: [
            'OpenAI APIキーの有効性',
            'ファイルサイズとフォーマット',
            'APIレスポンス内容',
            '処理時間制限'
          ]
        };
      }

      structuredLog.errors.push(structuredError);
    }

    // 処理タイプ別のトラブルシューティング情報
    if (errorDetails.processType === 'Zoom録音取得') {
      structuredLog.troubleshootingContext.commonCauses = [
        'Zoom APIトークンの期限切れまたは無効化',
        'Zoom APIクォータ制限に達している',
        'ダウンロードURLの有効期限切れ',
        'Google Driveのストレージ容量不足'
      ];
      structuredLog.troubleshootingContext.suggestedActions = [
        'Zoom APIトークンを更新する',
        'API使用量を確認し、必要に応じてプランを変更する',
        '新しいダウンロードURLを取得する',
        'Google Driveの容量を確認し、不要ファイルを削除する'
      ];
      structuredLog.troubleshootingContext.relatedComponents = [
        'ZoomAPIManager',
        'ZoomphoneService',
        'ZoomphoneProcessor',
        'Google Drive API'
      ];
    } else if (errorDetails.processType === '文字起こし') {
      structuredLog.troubleshootingContext.commonCauses = [
        'OpenAI APIクォータ制限に達している',
        'ファイルサイズが10MBを超えている',
        '音声ファイルの形式が対応していない',
        'OpenAI APIキーが無効または期限切れ'
      ];
      structuredLog.troubleshootingContext.suggestedActions = [
        'OpenAI APIの使用量と制限を確認する',
        'ファイル分割処理を確認する',
        '音声ファイルの形式を確認する',
        'OpenAI APIキーを更新する'
      ];
      structuredLog.troubleshootingContext.relatedComponents = [
        'TranscriptionService',
        'OpenAI Whisper API',
        'OpenAI GPT API',
        'FileProcessor'
      ];
    }

    return structuredLog;
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
    sendUnifiedErrorNotification: sendUnifiedErrorNotification,
    logDetailedError: logDetailedError,
    isHighPriorityError: isHighPriorityError // テスト用に公開
  };
})();