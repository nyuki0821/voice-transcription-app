/**
 * FileMovementServiceのテスト用スクリプト
 * ファイル移動サービスの各種機能をテストする
 */

/**
 * 処理結果オブジェクトの作成・操作テスト
 */
function testResultObjectOperations() {
  try {
    Logger.log("=== 処理結果オブジェクト操作テスト開始 ===");

    // 結果オブジェクトの作成
    var results = FileMovementService.createResultObject();
    Logger.log("初期結果オブジェクト: " + JSON.stringify(results));

    // 成功ケースの追加
    FileMovementService.addSuccessResult(results, "test_file_1.mp3", "abc123", "移動成功");
    FileMovementService.addSuccessResult(results, "test_file_2.mp3", "def456", "復旧完了");

    // 失敗ケースの追加
    FileMovementService.addFailureResult(results, "test_file_3.mp3", "ghi789", "移動先フォルダが見つかりません");

    // 結果の確認
    Logger.log("最終結果オブジェクト: " + JSON.stringify(results, null, 2));
    Logger.log("合計: " + results.total + ", 成功: " + results.recovered + ", 失敗: " + results.failed);

    // 詳細の確認
    for (var i = 0; i < results.details.length; i++) {
      var detail = results.details[i];
      Logger.log("詳細 " + (i + 1) + ": " + detail.fileName + " (" + detail.status + ") - " + detail.message);
    }

    Logger.log("=== 処理結果オブジェクト操作テスト完了 ===");
    return "処理結果オブジェクトの操作テストに成功しました";
  } catch (error) {
    Logger.log("処理結果オブジェクト操作テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * ログ出力機能のテスト
 */
function testLogProcessingResult() {
  try {
    Logger.log("=== ログ出力機能テスト開始 ===");

    var startTime = new Date();

    // テスト用の結果オブジェクトを作成
    var results = {
      total: 10,
      recovered: 7,
      failed: 3,
      recordIdFound: 8,
      recordIdNotFound: 2
    };

    // 少し時間を経過させる
    Utilities.sleep(100);

    // ログ出力のテスト
    var summary = FileMovementService.logProcessingResult(startTime, results, "見逃しエラー検知");
    Logger.log("生成されたサマリー: " + summary);

    // 別パターンのテスト
    var simpleResults = {
      total: 5,
      success: 4,
      error: 1
    };

    var simpleSummary = FileMovementService.logProcessingResult(startTime, simpleResults, "ファイル復旧処理");
    Logger.log("シンプルサマリー: " + simpleSummary);

    Logger.log("=== ログ出力機能テスト完了 ===");
    return "ログ出力機能のテストに成功しました";
  } catch (error) {
    Logger.log("ログ出力機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * ファイル移動機能のモックテスト
 * 実際のファイル移動は行わず、パラメータチェックのみ
 */
function testMoveFileWithFallbackValidation() {
  try {
    Logger.log("=== ファイル移動機能バリデーションテスト開始 ===");

    // パラメータエラーのテスト
    try {
      FileMovementService.moveFileWithFallback(null, "test_folder_id");
      Logger.log("エラー: nullファイルが受け入れられました");
      return "バリデーションテストに失敗しました";
    } catch (error) {
      Logger.log("正常: nullファイルが拒否されました - " + error.toString());
    }

    try {
      // モックファイルオブジェクト
      var mockFile = {
        getName: function () { return "test.mp3"; },
        getDescription: function () { return "test description"; },
        setDescription: function (desc) { this.description = desc; }
      };

      FileMovementService.moveFileWithFallback(mockFile, null);
      Logger.log("エラー: nullフォルダIDが受け入れられました");
      return "バリデーションテストに失敗しました";
    } catch (error) {
      Logger.log("正常: nullフォルダIDが拒否されました - " + error.toString());
    }

    Logger.log("=== ファイル移動機能バリデーションテスト完了 ===");
    return "ファイル移動機能のバリデーションテストに成功しました";
  } catch (error) {
    Logger.log("ファイル移動機能バリデーションテストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * FileMovementServiceの全テストを実行
 */
function runAllFileMovementServiceTests() {
  var startTime = new Date();
  Logger.log("====== FileMovementService 全テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  var results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };

  var tests = [
    { name: "処理結果オブジェクト操作", func: testResultObjectOperations },
    { name: "ログ出力機能", func: testLogProcessingResult },
    { name: "ファイル移動バリデーション", func: testMoveFileWithFallbackValidation }
  ];

  for (var i = 0; i < tests.length; i++) {
    var test = tests[i];
    results.total++;

    try {
      Logger.log("\n--- " + test.name + "テスト実行中 ---");
      var result = test.func();

      if (result.indexOf("エラー") === -1) {
        results.passed++;
        results.details.push({ name: test.name, status: "PASS", result: result });
        Logger.log(test.name + "テスト: PASS");
      } else {
        results.failed++;
        results.details.push({ name: test.name, status: "FAIL", reason: result });
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

  Logger.log("\n====== FileMovementService テスト結果 ======");
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

  Logger.log("====== FileMovementService 全テスト完了 ======");
  return results;
} 