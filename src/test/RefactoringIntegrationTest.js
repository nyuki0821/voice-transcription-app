/**
 * リファクタリング統合テスト
 * 新しいサービス間の連携とMain.jsの変更部分をテストする
 */

/**
 * 新しいサービス間の連携テスト
 */
function testServiceIntegration() {
  try {
    Logger.log("=== サービス間連携テスト開始 ===");

    // Constants + ConfigManager の連携テスト
    var sheetName = Constants.SHEET_NAMES.RECORDINGS;
    var config = ConfigManager.getConfig();
    Logger.log("Constants経由でシート名取得: " + sheetName);
    Logger.log("ConfigManager経由で設定取得: " + (config ? "成功" : "失敗"));

    // Constants + FileMovementService の連携テスト
    var results = FileMovementService.createResultObject();
    FileMovementService.addSuccessResult(results, "test.mp3", "abc123", "テスト成功");

    var statusSuccess = Constants.STATUS.SUCCESS;
    Logger.log("Constants経由でステータス取得: " + statusSuccess);
    Logger.log("FileMovementService結果作成: " + JSON.stringify(results));

    // エラーパターンとメッセージフォーマットの連携
    var errorPattern = Constants.ERROR_PATTERNS[0];
    var errorMessage = Constants.formatMessage(
      Constants.ERROR_MESSAGES.SHEET_NOT_FOUND,
      { sheetName: "test_sheet" }
    );
    Logger.log("エラーパターン例: " + errorPattern.pattern + " -> " + errorPattern.issue);
    Logger.log("フォーマット済みエラーメッセージ: " + errorMessage);

    Logger.log("=== サービス間連携テスト完了 ===");
    return "サービス間連携テストに成功しました";
  } catch (error) {
    Logger.log("サービス間連携テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * Main.jsのリファクタリング部分のテスト（モック）
 */
function testMainJsRefactoringMock() {
  try {
    Logger.log("=== Main.jsリファクタリング部分テスト開始 ===");

    // recoverErrorFiles関数で使用される新機能のテスト
    Logger.log("1. FileMovementService機能テスト:");
    var results = FileMovementService.createResultObject();
    results.total = 5;
    FileMovementService.addSuccessResult(results, "file1.mp3", "id1", "復旧成功");
    FileMovementService.addFailureResult(results, "file2.mp3", "id2", "復旧失敗");

    var startTime = new Date();
    Utilities.sleep(50); // 少し時間を経過させる
    var summary = FileMovementService.logProcessingResult(startTime, results, "エラーファイル復旧");
    Logger.log("  生成されたサマリー: " + summary);

    // checkForErrorInTranscription関数で使用される新機能のテスト
    Logger.log("2. Constants機能テスト:");
    var testTranscription = "GPT-4o-mini API呼び出しエラーが発生しました";
    var foundError = false;

    for (var i = 0; i < Constants.ERROR_PATTERNS.length; i++) {
      var pattern = Constants.ERROR_PATTERNS[i];
      if (testTranscription.indexOf(pattern.pattern) !== -1) {
        foundError = true;
        Logger.log("  エラーパターン検知: " + pattern.issue);
        break;
      }
    }

    if (!foundError) {
      Logger.log("  エラーパターンは検知されませんでした");
    }

    // detectAndRecoverPartialFailures関数で使用される新機能のテスト
    Logger.log("3. ConfigManager機能テスト:");
    var recordingsSheetId = ConfigManager.get('RECORDINGS_SHEET_ID');
    var processedSheetId = ConfigManager.get('PROCESSED_SHEET_ID');
    Logger.log("  RECORDINGS_SHEET_ID: " + (recordingsSheetId ? "設定済み" : "未設定"));
    Logger.log("  PROCESSED_SHEET_ID: " + (processedSheetId ? "設定済み" : "未設定"));

    Logger.log("=== Main.jsリファクタリング部分テスト完了 ===");
    return "Main.jsリファクタリング部分のテストに成功しました";
  } catch (error) {
    Logger.log("Main.jsリファクタリング部分テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 下位互換性テスト
 */
function testBackwardCompatibility() {
  try {
    Logger.log("=== 下位互換性テスト開始 ===");

    // 既存のグローバル変数が利用可能かテスト
    Logger.log("既存グローバル変数の確認:");

    // SPREADSHEET_IDの確認（Main.jsで定義されているはず）
    try {
      if (typeof SPREADSHEET_ID !== 'undefined') {
        Logger.log("  SPREADSHEET_ID: 利用可能");
      } else {
        Logger.log("  SPREADSHEET_ID: 未定義");
      }
    } catch (e) {
      Logger.log("  SPREADSHEET_ID: アクセスエラー - " + e.toString());
    }

    // 既存の関数が利用可能かテスト
    Logger.log("既存関数の確認:");
    var existingFunctions = [
      'getSystemSettings',
      'loadSettings',
      'extractMetadataFromFile'
    ];

    for (var i = 0; i < existingFunctions.length; i++) {
      var funcName = existingFunctions[i];
      try {
        if (typeof eval(funcName) === 'function') {
          Logger.log("  " + funcName + ": 利用可能");
        } else {
          Logger.log("  " + funcName + ": 関数ではない");
        }
      } catch (e) {
        Logger.log("  " + funcName + ": アクセスエラー - " + e.toString());
      }
    }

    Logger.log("=== 下位互換性テスト完了 ===");
    return "下位互換性テストに成功しました";
  } catch (error) {
    Logger.log("下位互換性テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * パフォーマンステスト
 */
function testPerformanceImprovements() {
  try {
    Logger.log("=== パフォーマンステスト開始 ===");

    // 設定取得のパフォーマンステスト
    Logger.log("設定取得パフォーマンス:");

    // ConfigManagerのキャッシュ効果テスト
    ConfigManager.clearCache();

    var times = [];
    for (var i = 0; i < 3; i++) {
      var start = new Date().getTime();
      var config = ConfigManager.getConfig();
      var end = new Date().getTime();
      times.push(end - start);
      Logger.log("  " + (i + 1) + "回目: " + times[i] + "ms");
    }

    // 平均時間の計算
    var average = times.reduce(function (sum, time) { return sum + time; }, 0) / times.length;
    Logger.log("  平均取得時間: " + average.toFixed(2) + "ms");

    // Constants使用のパフォーマンステスト
    Logger.log("Constants使用パフォーマンス:");
    var constantsStart = new Date().getTime();

    for (var j = 0; j < 100; j++) {
      var status = Constants.STATUS.SUCCESS;
      var sheetName = Constants.SHEET_NAMES.RECORDINGS;
      var isAudio = Constants.isAudioFile(null); // 軽量なテスト
    }

    var constantsEnd = new Date().getTime();
    Logger.log("  100回アクセス時間: " + (constantsEnd - constantsStart) + "ms");

    Logger.log("=== パフォーマンステスト完了 ===");
    return "パフォーマンステストに成功しました";
  } catch (error) {
    Logger.log("パフォーマンステストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * エラーハンドリングテスト
 */
function testErrorHandling() {
  try {
    Logger.log("=== エラーハンドリングテスト開始 ===");

    // FileMovementServiceのエラーハンドリング
    Logger.log("FileMovementServiceエラーハンドリング:");
    try {
      FileMovementService.moveFileWithFallback(null, "test");
      Logger.log("  エラー: null引数が受け入れられました");
    } catch (e) {
      Logger.log("  正常: null引数が適切に拒否されました");
    }

    // Constantsのエラーハンドリング
    Logger.log("Constantsエラーハンドリング:");
    var nullFileResult = Constants.isAudioFile(null);
    Logger.log("  nullファイル判定: " + nullFileResult + " (falseが期待値)");

    var nullFileNameResult = Constants.extractRecordIdFromFileName(null);
    Logger.log("  nullファイル名からID抽出: " + nullFileNameResult + " (nullが期待値)");

    // ConfigManagerのエラーハンドリング
    Logger.log("ConfigManagerエラーハンドリング:");
    var nonExistentConfig = ConfigManager.get('NON_EXISTENT_KEY', 'default');
    Logger.log("  存在しない設定取得: " + nonExistentConfig + " (defaultが期待値)");

    Logger.log("=== エラーハンドリングテスト完了 ===");
    return "エラーハンドリングテストに成功しました";
  } catch (error) {
    Logger.log("エラーハンドリングテストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * リファクタリング統合テストの全テストを実行
 */
function runAllRefactoringIntegrationTests() {
  var results = [];

  try {
    Logger.log("====== リファクタリング統合テスト 全テスト開始 ======");

    // 各テストを実行
    results.push({ name: "サービス間連携", result: testServiceIntegration() });
    results.push({ name: "Main.jsリファクタリング部分", result: testMainJsRefactoringMock() });
    results.push({ name: "下位互換性", result: testBackwardCompatibility() });
    results.push({ name: "パフォーマンス", result: testPerformanceImprovements() });
    results.push({ name: "エラーハンドリング", result: testErrorHandling() });

    // 結果をログに出力
    Logger.log("====== テスト結果サマリー ======");
    for (var i = 0; i < results.length; i++) {
      Logger.log((i + 1) + ". " + results[i].name + ": " +
        (results[i].result.indexOf("エラー") === -1 ? "成功" : "失敗"));
    }

    Logger.log("====== リファクタリング統合テスト 全テスト完了 ======");
    return "全テストが完了しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("テスト実行中にエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
} 