/**
 * 統一エラーハンドリングシステム
 * プロジェクト全体で一貫したエラー処理を提供
 */
var ErrorHandler = (function () {

  // エラーレベル定数
  var ERROR_LEVELS = {
    INFO: 'INFO',
    WARNING: 'WARNING',
    ERROR: 'ERROR',
    CRITICAL: 'CRITICAL'
  };

  // エラーカテゴリ定数
  var ERROR_CATEGORIES = {
    SYSTEM: 'SYSTEM',
    API: 'API',
    FILE_OPERATION: 'FILE_OPERATION',
    DATA_VALIDATION: 'DATA_VALIDATION',
    NETWORK: 'NETWORK',
    PERMISSION: 'PERMISSION',
    CONFIGURATION: 'CONFIGURATION'
  };

  // エラー通知設定
  var NOTIFICATION_CONFIG = {
    CRITICAL_ERRORS_NOTIFY: true,
    ERROR_BATCH_SIZE: 10,
    NOTIFICATION_COOLDOWN: 300000 // 5分間のクールダウン
  };

  // 最後の通知時刻を記録
  var lastNotificationTime = {};

  /**
   * 統一エラーハンドリング関数
   * @param {Error|string} error - エラーオブジェクトまたはメッセージ
   * @param {string} context - エラーが発生したコンテキスト
   * @param {string} level - エラーレベル（INFO, WARNING, ERROR, CRITICAL）
   * @param {string} category - エラーカテゴリ
   * @param {Object} additionalInfo - 追加情報
   * @return {Object} 処理されたエラー情報
   */
  function handleError(error, context, level, category, additionalInfo) {
    level = level || ERROR_LEVELS.ERROR;
    category = category || ERROR_CATEGORIES.SYSTEM;
    additionalInfo = additionalInfo || {};

    var errorInfo = {
      timestamp: new Date(),
      level: level,
      category: category,
      context: context,
      message: getErrorMessage(error),
      stack: getErrorStack(error),
      additionalInfo: additionalInfo,
      errorId: generateErrorId()
    };

    // ログ出力
    logError(errorInfo);

    // 重要なエラーの場合は通知
    if (shouldNotify(errorInfo)) {
      notifyError(errorInfo);
    }

    // エラー統計を更新
    updateErrorStatistics(errorInfo);

    return errorInfo;
  }

  /**
   * try-catch ブロックのラッパー関数
   * @param {Function} operation - 実行する処理
   * @param {string} context - コンテキスト
   * @param {Object} options - オプション設定
   * @return {Object} 実行結果
   */
  function safeExecute(operation, context, options) {
    options = options || {};
    var retryCount = options.retryCount || 0;
    var maxRetries = options.maxRetries || 0;
    var retryDelay = options.retryDelay || 1000;

    for (var attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        var result = operation();

        // 成功した場合
        if (attempt > 0) {
          Logger.log(context + ': ' + attempt + '回目の試行で成功しました');
        }

        return {
          success: true,
          result: result,
          attempts: attempt + 1
        };
      } catch (error) {
        var errorInfo = handleError(
          error,
          context + ' (試行 ' + (attempt + 1) + '/' + (maxRetries + 1) + ')',
          attempt === maxRetries ? ERROR_LEVELS.ERROR : ERROR_LEVELS.WARNING,
          options.category || ERROR_CATEGORIES.SYSTEM,
          { attempt: attempt + 1, maxRetries: maxRetries + 1 }
        );

        // 最後の試行でない場合は待機してリトライ
        if (attempt < maxRetries) {
          Logger.log(context + ': ' + retryDelay + 'ms待機後にリトライします');
          Utilities.sleep(retryDelay);
        } else {
          // 最終的に失敗
          return {
            success: false,
            error: errorInfo,
            attempts: attempt + 1
          };
        }
      }
    }
  }

  /**
   * API呼び出し専用のエラーハンドリング
   * @param {Function} apiCall - API呼び出し関数
   * @param {string} apiName - API名
   * @param {Object} options - オプション
   * @return {Object} API呼び出し結果
   */
  function handleApiCall(apiCall, apiName, options) {
    options = options || {};
    options.category = ERROR_CATEGORIES.API;
    options.maxRetries = options.maxRetries || 2;
    options.retryDelay = options.retryDelay || 2000;

    var context = 'API呼び出し: ' + apiName;

    return safeExecute(function () {
      var startTime = new Date();
      var result = apiCall();
      var endTime = new Date();
      var duration = endTime - startTime;

      // API呼び出し成功ログ
      Logger.log(context + ' 成功 (処理時間: ' + duration + 'ms)');

      return {
        data: result,
        duration: duration,
        timestamp: endTime
      };
    }, context, options);
  }

  /**
   * ファイル操作専用のエラーハンドリング
   * @param {Function} fileOperation - ファイル操作関数
   * @param {string} operationName - 操作名
   * @param {string} fileName - ファイル名
   * @param {Object} options - オプション
   * @return {Object} ファイル操作結果
   */
  function handleFileOperation(fileOperation, operationName, fileName, options) {
    options = options || {};
    options.category = ERROR_CATEGORIES.FILE_OPERATION;
    options.maxRetries = options.maxRetries || 1;

    var context = 'ファイル操作: ' + operationName + ' (' + fileName + ')';

    return safeExecute(fileOperation, context, options);
  }

  /**
   * バッチ処理のエラーハンドリング
   * @param {Array} items - 処理対象アイテム
   * @param {Function} processor - 各アイテムの処理関数
   * @param {string} batchName - バッチ名
   * @param {Object} options - オプション
   * @return {Object} バッチ処理結果
   */
  function handleBatchProcessing(items, processor, batchName, options) {
    options = options || {};
    var continueOnError = options.continueOnError !== false;
    var maxErrors = options.maxErrors || items.length;

    var results = {
      total: items.length,
      success: 0,
      failed: 0,
      errors: [],
      results: []
    };

    var startTime = new Date();
    Logger.log(batchName + ' 開始: ' + items.length + '件の処理');

    for (var i = 0; i < items.length; i++) {
      var item = items[i];

      try {
        var result = processor(item, i);
        results.success++;
        results.results.push({
          index: i,
          item: item,
          result: result,
          success: true
        });
      } catch (error) {
        var errorInfo = handleError(
          error,
          batchName + ' - アイテム ' + (i + 1),
          ERROR_LEVELS.WARNING,
          ERROR_CATEGORIES.SYSTEM,
          { itemIndex: i, item: item }
        );

        results.failed++;
        results.errors.push(errorInfo);
        results.results.push({
          index: i,
          item: item,
          error: errorInfo,
          success: false
        });

        // エラー数が上限に達した場合は中断
        if (results.failed >= maxErrors) {
          Logger.log(batchName + ': エラー数が上限(' + maxErrors + ')に達したため処理を中断します');
          break;
        }

        // continueOnErrorがfalseの場合は即座に中断
        if (!continueOnError) {
          Logger.log(batchName + ': エラーが発生したため処理を中断します');
          break;
        }
      }
    }

    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;

    var summary = batchName + ' 完了: 成功=' + results.success + '件, 失敗=' + results.failed + '件, 処理時間=' + duration + '秒';
    Logger.log(summary);

    results.duration = duration;
    results.summary = summary;

    return results;
  }

  /**
   * エラーメッセージを取得
   * @param {Error|string} error - エラー
   * @return {string} エラーメッセージ
   */
  function getErrorMessage(error) {
    if (typeof error === 'string') {
      return error;
    }

    if (error && error.message) {
      return error.message;
    }

    if (error && error.toString) {
      return error.toString();
    }

    return 'Unknown error';
  }

  /**
   * エラースタックを取得
   * @param {Error} error - エラーオブジェクト
   * @return {string} スタックトレース
   */
  function getErrorStack(error) {
    if (error && error.stack) {
      return error.stack;
    }
    return null;
  }

  /**
   * エラーIDを生成
   * @return {string} ユニークなエラーID
   */
  function generateErrorId() {
    var timestamp = new Date().getTime();
    var random = Math.floor(Math.random() * 1000);
    return 'ERR_' + timestamp + '_' + random;
  }

  /**
   * エラーをログ出力
   * @param {Object} errorInfo - エラー情報
   */
  function logError(errorInfo) {
    var logMessage = '[' + errorInfo.level + '] ' +
      errorInfo.context + ': ' +
      errorInfo.message +
      ' (ID: ' + errorInfo.errorId + ')';

    Logger.log(logMessage);

    // 追加情報がある場合は出力
    if (Object.keys(errorInfo.additionalInfo).length > 0) {
      Logger.log('追加情報: ' + JSON.stringify(errorInfo.additionalInfo));
    }

    // スタックトレースがある場合は出力
    if (errorInfo.stack) {
      Logger.log('スタックトレース: ' + errorInfo.stack);
    }
  }

  /**
   * エラー通知が必要かチェック
   * @param {Object} errorInfo - エラー情報
   * @return {boolean} 通知が必要な場合true
   */
  function shouldNotify(errorInfo) {
    // CRITICALレベルは常に通知
    if (errorInfo.level === ERROR_LEVELS.CRITICAL) {
      return true;
    }

    // 通知設定が無効の場合は通知しない
    if (!NOTIFICATION_CONFIG.CRITICAL_ERRORS_NOTIFY) {
      return false;
    }

    // クールダウン期間中は通知しない
    var lastNotification = lastNotificationTime[errorInfo.category] || 0;
    var now = new Date().getTime();

    if (now - lastNotification < NOTIFICATION_CONFIG.NOTIFICATION_COOLDOWN) {
      return false;
    }

    return errorInfo.level === ERROR_LEVELS.ERROR;
  }

  /**
   * エラー通知を送信
   * @param {Object} errorInfo - エラー情報
   */
  function notifyError(errorInfo) {
    try {
      var config = ConfigManager.getConfig();
      var adminEmails = config.ADMIN_EMAILS || [];

      if (adminEmails.length === 0) {
        Logger.log('管理者メールアドレスが設定されていないため、エラー通知をスキップします');
        return;
      }

      var subject = '[' + errorInfo.level + '] システムエラー通知: ' + errorInfo.context;
      var body = createErrorNotificationBody(errorInfo);

      for (var i = 0; i < adminEmails.length; i++) {
        MailApp.sendEmail(adminEmails[i], subject, body);
      }

      // 最後の通知時刻を更新
      lastNotificationTime[errorInfo.category] = new Date().getTime();

      Logger.log('エラー通知メールを送信しました: ' + errorInfo.errorId);
    } catch (notificationError) {
      Logger.log('エラー通知の送信に失敗: ' + notificationError.toString());
    }
  }

  /**
   * エラー通知メールの本文を作成
   * @param {Object} errorInfo - エラー情報
   * @return {string} メール本文
   */
  function createErrorNotificationBody(errorInfo) {
    var body = 'システムエラーが発生しました。\n\n';
    body += '【エラー詳細】\n';
    body += 'エラーID: ' + errorInfo.errorId + '\n';
    body += '発生時刻: ' + Utilities.formatDate(errorInfo.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') + '\n';
    body += 'レベル: ' + errorInfo.level + '\n';
    body += 'カテゴリ: ' + errorInfo.category + '\n';
    body += 'コンテキスト: ' + errorInfo.context + '\n';
    body += 'メッセージ: ' + errorInfo.message + '\n';

    if (Object.keys(errorInfo.additionalInfo).length > 0) {
      body += '\n【追加情報】\n';
      body += JSON.stringify(errorInfo.additionalInfo, null, 2) + '\n';
    }

    if (errorInfo.stack) {
      body += '\n【スタックトレース】\n';
      body += errorInfo.stack + '\n';
    }

    body += '\n※このメールは自動送信されています。';

    return body;
  }

  /**
   * エラー統計を更新
   * @param {Object} errorInfo - エラー情報
   */
  function updateErrorStatistics(errorInfo) {
    // 簡易的な統計更新（実装は省略）
    // 実際の実装では、PropertiesServiceやスプレッドシートに統計を保存
    Logger.log('エラー統計を更新: ' + errorInfo.category + ' - ' + errorInfo.level);
  }

  /**
   * 設定検証エラーのヘルパー
   * @param {string} configKey - 設定キー
   * @param {string} context - コンテキスト
   * @return {Object} エラー情報
   */
  function handleConfigurationError(configKey, context) {
    return handleError(
      '必要な設定が見つかりません: ' + configKey,
      context,
      ERROR_LEVELS.CRITICAL,
      ERROR_CATEGORIES.CONFIGURATION,
      { configKey: configKey }
    );
  }

  /**
   * 権限エラーのヘルパー
   * @param {string} resource - リソース名
   * @param {string} operation - 操作名
   * @param {string} context - コンテキスト
   * @return {Object} エラー情報
   */
  function handlePermissionError(resource, operation, context) {
    return handleError(
      resource + 'への' + operation + '権限がありません',
      context,
      ERROR_LEVELS.ERROR,
      ERROR_CATEGORIES.PERMISSION,
      { resource: resource, operation: operation }
    );
  }

  // 公開メソッド
  return {
    // 基本エラーハンドリング
    handleError: handleError,
    safeExecute: safeExecute,

    // 専用エラーハンドリング
    handleApiCall: handleApiCall,
    handleFileOperation: handleFileOperation,
    handleBatchProcessing: handleBatchProcessing,

    // ヘルパー関数
    handleConfigurationError: handleConfigurationError,
    handlePermissionError: handlePermissionError,

    // 定数
    ERROR_LEVELS: ERROR_LEVELS,
    ERROR_CATEGORIES: ERROR_CATEGORIES
  };
})(); 