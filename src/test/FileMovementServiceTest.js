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
  var results = [];

  try {
    Logger.log("====== FileMovementService 全テスト開始 ======");

    // 各テストを実行
    results.push({ name: "処理結果オブジェクト操作", result: testResultObjectOperations() });
    results.push({ name: "ログ出力機能", result: testLogProcessingResult() });
    results.push({ name: "ファイル移動バリデーション", result: testMoveFileWithFallbackValidation() });

    // 結果をログに出力
    Logger.log("====== テスト結果サマリー ======");
    for (var i = 0; i < results.length; i++) {
      Logger.log((i + 1) + ". " + results[i].name + ": " +
        (results[i].result.indexOf("エラー") === -1 ? "成功" : "失敗"));
    }

    Logger.log("====== FileMovementService 全テスト完了 ======");
    return "全テストが完了しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("テスト実行中にエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
} 