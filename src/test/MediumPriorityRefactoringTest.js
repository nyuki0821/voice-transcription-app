/**
 * 優先度中リファクタリングテスト
 * Main.js分割、エラーハンドリング統一、処理フロー最適化のテスト
 */

/**
 * RecoveryServiceのテスト
 */
function testRecoveryService() {
  try {
    Logger.log("=== RecoveryService テスト開始 ===");

    // 1. 基本機能テスト
    Logger.log("1. RecoveryService基本機能テスト:");

    // RecoveryServiceが正しく定義されているかチェック
    if (typeof RecoveryService === 'undefined') {
      throw new Error('RecoveryServiceが定義されていません');
    }

    var requiredMethods = [
      'recoverInterruptedFiles',
      'recoverErrorFiles',
      'resetPendingTranscriptions',
      'forceRecoverAllErrorFiles',
      'runFullRecoveryProcess',
      'findAndMoveFileToSource'
    ];

    for (var i = 0; i < requiredMethods.length; i++) {
      var method = requiredMethods[i];
      if (typeof RecoveryService[method] !== 'function') {
        throw new Error('RecoveryService.' + method + 'が関数として定義されていません');
      }
      Logger.log("  ✓ " + method + " メソッドが定義されています");
    }

    // 2. 下位互換性テスト
    Logger.log("2. Main.js下位互換性テスト:");

    var mainCompatibilityMethods = [
      'recoverInterruptedFiles',
      'recoverErrorFiles',
      'resetPendingTranscriptions',
      'forceRecoverAllErrorFiles',
      'runFullRecoveryProcess'
    ];

    for (var i = 0; i < mainCompatibilityMethods.length; i++) {
      var method = mainCompatibilityMethods[i];
      // Google Apps Scriptではwindowオブジェクトは存在しないため、グローバルスコープを直接チェック
      try {
        if (typeof eval(method) === 'function') {
          Logger.log("  ✓ " + method + " グローバル関数が利用可能です");
        } else {
          Logger.log("  ⚠ " + method + " グローバル関数が見つかりません（下位互換性の問題の可能性）");
        }
      } catch (e) {
        Logger.log("  ⚠ " + method + " グローバル関数が見つかりません（下位互換性の問題の可能性）");
      }
    }

    Logger.log("RecoveryService テスト完了");
    return true;

  } catch (error) {
    Logger.log("RecoveryService テストエラー: " + error.toString());
    return false;
  }
}

/**
 * PartialFailureDetectorのテスト
 */
function testPartialFailureDetector() {
  try {
    Logger.log("=== PartialFailureDetector テスト開始 ===");

    // 1. 基本機能テスト
    Logger.log("1. PartialFailureDetector基本機能テスト:");

    // PartialFailureDetectorが正しく定義されているかチェック
    if (typeof PartialFailureDetector === 'undefined') {
      throw new Error('PartialFailureDetectorが定義されていません');
    }

    var requiredMethods = [
      'detectAndRecoverPartialFailures',
      'checkForErrorInTranscription',
      'moveFileToErrorFolder',
      'getDetectionStatistics'
    ];

    for (var i = 0; i < requiredMethods.length; i++) {
      var method = requiredMethods[i];
      if (typeof PartialFailureDetector[method] !== 'function') {
        throw new Error('PartialFailureDetector.' + method + 'が関数として定義されていません');
      }
      Logger.log("  ✓ " + method + " メソッドが定義されています");
    }

    // 2. エラーパターン検知テスト（モック）
    Logger.log("2. エラーパターン検知テスト:");

    var testPatterns = [
      'insufficient_quota',
      'GPT-4o-mini API呼び出しエラー',
      '【文字起こし失敗:',
      'OpenAI APIからのレスポンスエラー'
    ];

    for (var i = 0; i < testPatterns.length; i++) {
      var pattern = testPatterns[i];
      var found = false;

      // Constants.ERROR_PATTERNSでパターンが定義されているかチェック
      for (var j = 0; j < Constants.ERROR_PATTERNS.length; j++) {
        if (Constants.ERROR_PATTERNS[j].pattern === pattern) {
          found = true;
          Logger.log("  ✓ エラーパターン '" + pattern + "' が定義されています");
          break;
        }
      }

      if (!found) {
        Logger.log("  ⚠ エラーパターン '" + pattern + "' が見つかりません");
      }
    }

    // 3. 下位互換性テスト
    Logger.log("3. Main.js下位互換性テスト:");

    var mainCompatibilityMethods = [
      'detectAndRecoverPartialFailures',
      'checkForErrorInTranscription',
      'moveFileToErrorFolder'
    ];

    for (var i = 0; i < mainCompatibilityMethods.length; i++) {
      var method = mainCompatibilityMethods[i];
      // Google Apps Scriptではwindowオブジェクトは存在しないため、グローバルスコープを直接チェック
      try {
        if (typeof eval(method) === 'function') {
          Logger.log("  ✓ " + method + " グローバル関数が利用可能です");
        } else {
          Logger.log("  ⚠ " + method + " グローバル関数が見つかりません（下位互換性の問題の可能性）");
        }
      } catch (e) {
        Logger.log("  ⚠ " + method + " グローバル関数が見つかりません（下位互換性の問題の可能性）");
      }
    }

    Logger.log("PartialFailureDetector テスト完了");
    return true;

  } catch (error) {
    Logger.log("PartialFailureDetector テストエラー: " + error.toString());
    return false;
  }
}

/**
 * ErrorHandlerのテスト
 */
function testErrorHandler() {
  try {
    Logger.log("=== ErrorHandler テスト開始 ===");

    // 1. 基本機能テスト
    Logger.log("1. ErrorHandler基本機能テスト:");

    // ErrorHandlerが正しく定義されているかチェック
    if (typeof ErrorHandler === 'undefined') {
      throw new Error('ErrorHandlerが定義されていません');
    }

    var requiredMethods = [
      'handleError',
      'safeExecute',
      'handleApiCall',
      'handleFileOperation',
      'handleBatchProcessing',
      'handleConfigurationError',
      'handlePermissionError'
    ];

    for (var i = 0; i < requiredMethods.length; i++) {
      var method = requiredMethods[i];
      if (typeof ErrorHandler[method] !== 'function') {
        throw new Error('ErrorHandler.' + method + 'が関数として定義されていません');
      }
      Logger.log("  ✓ " + method + " メソッドが定義されています");
    }

    // 2. 定数テスト
    Logger.log("2. ErrorHandler定数テスト:");

    var requiredConstants = ['ERROR_LEVELS', 'ERROR_CATEGORIES'];

    for (var i = 0; i < requiredConstants.length; i++) {
      var constant = requiredConstants[i];
      if (typeof ErrorHandler[constant] === 'undefined') {
        throw new Error('ErrorHandler.' + constant + 'が定義されていません');
      }
      Logger.log("  ✓ " + constant + " 定数が定義されています");
    }

    // エラーレベルの確認
    var expectedLevels = ['INFO', 'WARNING', 'ERROR', 'CRITICAL'];
    for (var i = 0; i < expectedLevels.length; i++) {
      var level = expectedLevels[i];
      if (!ErrorHandler.ERROR_LEVELS[level]) {
        Logger.log("  ⚠ エラーレベル '" + level + "' が見つかりません");
      } else {
        Logger.log("  ✓ エラーレベル '" + level + "' が定義されています");
      }
    }

    // 3. safeExecute機能テスト
    Logger.log("3. safeExecute機能テスト:");

    // 成功ケース
    var successResult = ErrorHandler.safeExecute(function () {
      return "テスト成功";
    }, "成功テスト");

    if (successResult.success && successResult.result === "テスト成功") {
      Logger.log("  ✓ 成功ケースが正常に動作します");
    } else {
      Logger.log("  ⚠ 成功ケースで予期しない結果: " + JSON.stringify(successResult));
    }

    // エラーケース
    var errorResult = ErrorHandler.safeExecute(function () {
      throw new Error("テストエラー");
    }, "エラーテスト");

    if (!errorResult.success && errorResult.error) {
      Logger.log("  ✓ エラーケースが正常に処理されます");
    } else {
      Logger.log("  ⚠ エラーケースで予期しない結果: " + JSON.stringify(errorResult));
    }

    Logger.log("ErrorHandler テスト完了");
    return true;

  } catch (error) {
    Logger.log("ErrorHandler テストエラー: " + error.toString());
    return false;
  }
}

/**
 * Main.js統合テスト
 */
function testMainJsIntegration() {
  try {
    Logger.log("=== Main.js統合テスト開始 ===");

    // 1. main関数のエラーハンドリング統合テスト
    Logger.log("1. main関数エラーハンドリング統合テスト:");

    // main関数が存在するかチェック
    if (typeof main !== 'function') {
      throw new Error('main関数が定義されていません');
    }

    Logger.log("  ✓ main関数が定義されています");

    // 2. 依存関係チェック
    Logger.log("2. 依存関係チェック:");

    var requiredDependencies = [
      'ConfigManager',
      'Constants',
      'FileMovementService',
      'ErrorHandler',
      'RecoveryService',
      'PartialFailureDetector'
    ];

    for (var i = 0; i < requiredDependencies.length; i++) {
      var dependency = requiredDependencies[i];
      // Google Apps Scriptではwindowオブジェクトは存在しないため、グローバルスコープを直接チェック
      try {
        if (typeof eval(dependency) !== 'undefined') {
          Logger.log("  ✓ 依存関係 '" + dependency + "' が利用可能です");
        } else {
          Logger.log("  ⚠ 依存関係 '" + dependency + "' が見つかりません");
        }
      } catch (e) {
        Logger.log("  ⚠ 依存関係 '" + dependency + "' が見つかりません");
      }
    }

    // 3. 下位互換性関数チェック
    Logger.log("3. 下位互換性関数チェック:");

    var compatibilityFunctions = [
      'getSystemSettings',
      'getDefaultSettings',
      'loadSettings'
    ];

    for (var i = 0; i < compatibilityFunctions.length; i++) {
      var func = compatibilityFunctions[i];
      // Google Apps Scriptではwindowオブジェクトは存在しないため、グローバルスコープを直接チェック
      try {
        if (typeof eval(func) === 'function') {
          Logger.log("  ✓ 下位互換性関数 '" + func + "' が利用可能です");
        } else {
          Logger.log("  ⚠ 下位互換性関数 '" + func + "' が見つかりません");
        }
      } catch (e) {
        Logger.log("  ⚠ 下位互換性関数 '" + func + "' が見つかりません");
      }
    }

    Logger.log("Main.js統合テスト完了");
    return true;

  } catch (error) {
    Logger.log("Main.js統合テストエラー: " + error.toString());
    return false;
  }
}

/**
 * パフォーマンステスト
 */
function testPerformanceImprovements() {
  try {
    Logger.log("=== パフォーマンステスト開始 ===");

    // 1. エラーハンドリングのオーバーヘッドテスト
    Logger.log("1. エラーハンドリングオーバーヘッドテスト:");

    var iterations = 100;
    var startTime = new Date();

    for (var i = 0; i < iterations; i++) {
      ErrorHandler.safeExecute(function () {
        return "テスト" + i;
      }, "パフォーマンステスト");
    }

    var endTime = new Date();
    var duration = endTime - startTime;
    var avgTime = duration / iterations;

    Logger.log("  " + iterations + "回の実行時間: " + duration + "ms");
    Logger.log("  平均実行時間: " + avgTime.toFixed(2) + "ms");

    if (avgTime < 10) {
      Logger.log("  ✓ パフォーマンスは良好です");
    } else {
      Logger.log("  ⚠ パフォーマンスに改善の余地があります");
    }

    // 2. メモリ使用量の概算チェック
    Logger.log("2. メモリ使用量チェック:");

    // 大きなオブジェクトを作成してメモリ使用をテスト
    var largeArray = [];
    for (var i = 0; i < 1000; i++) {
      largeArray.push({
        index: i,
        data: "テストデータ" + i,
        timestamp: new Date()
      });
    }

    var batchResult = ErrorHandler.handleBatchProcessing(
      largeArray.slice(0, 10), // 最初の10件のみテスト
      function (item) {
        return item.data + "_processed";
      },
      "メモリテストバッチ"
    );

    if (batchResult.success === 10 && batchResult.failed === 0) {
      Logger.log("  ✓ バッチ処理が正常に動作します");
    } else {
      Logger.log("  ⚠ バッチ処理で問題が発生: " + JSON.stringify(batchResult));
    }

    Logger.log("パフォーマンステスト完了");
    return true;

  } catch (error) {
    Logger.log("パフォーマンステストエラー: " + error.toString());
    return false;
  }
}

/**
 * 優先度中リファクタリング全体テスト実行
 */
function runMediumPriorityRefactoringTests() {
  var startTime = new Date();
  Logger.log("=== 優先度中リファクタリングテスト開始 ===");
  Logger.log("開始時刻: " + startTime);

  var results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };

  var tests = [
    { name: "RecoveryService", func: testRecoveryService },
    { name: "PartialFailureDetector", func: testPartialFailureDetector },
    { name: "ErrorHandler", func: testErrorHandler },
    { name: "Main.js統合", func: testMainJsIntegration },
    { name: "パフォーマンス", func: testPerformanceImprovements }
  ];

  for (var i = 0; i < tests.length; i++) {
    var test = tests[i];
    results.total++;

    try {
      Logger.log("\n--- " + test.name + "テスト実行中 ---");
      var success = test.func();

      if (success) {
        results.passed++;
        results.details.push({ name: test.name, status: "PASS" });
        Logger.log(test.name + "テスト: PASS");
      } else {
        results.failed++;
        results.details.push({ name: test.name, status: "FAIL", reason: "テスト関数がfalseを返しました" });
        Logger.log(test.name + "テスト: FAIL");
      }
    } catch (error) {
      results.failed++;
      results.details.push({ name: test.name, status: "ERROR", reason: error.toString() });
      Logger.log(test.name + "テスト: ERROR - " + error.toString());
    }
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("\n=== 優先度中リファクタリングテスト結果 ===");
  Logger.log("総テスト数: " + results.total);
  Logger.log("成功: " + results.passed);
  Logger.log("失敗: " + results.failed);
  Logger.log("実行時間: " + duration + "秒");
  Logger.log("成功率: " + ((results.passed / results.total) * 100).toFixed(1) + "%");

  // 詳細結果
  Logger.log("\n=== 詳細結果 ===");
  for (var i = 0; i < results.details.length; i++) {
    var detail = results.details[i];
    var status = detail.status === "PASS" ? "✓" : "✗";
    var message = status + " " + detail.name + ": " + detail.status;
    if (detail.reason) {
      message += " (" + detail.reason + ")";
    }
    Logger.log(message);
  }

  return results;
} 