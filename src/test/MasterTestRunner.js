/**
 * マスターテストランナー
 * 全てのテストスイートを一括実行し、総合的な結果を提供する
 * 
 * 追加されたテストスイート:
 * - WhisperServiceTest: Whisperベース機能の単体テスト
 * - WhisperIntegrationTest: Whisperベースシステムの統合テスト
 */

/**
 * 全テストスイートを実行
 */
function runAllTests() {
  var overallResults = [];
  var startTime = new Date();

  try {
    Logger.log("========================================");
    Logger.log("リファクタリング後 全テスト実行開始");
    Logger.log("実行開始時刻: " + startTime.toLocaleString());
    Logger.log("========================================");

    // 1. 既存の環境設定テスト
    Logger.log("\n【1/5】環境設定テスト実行中...");
    try {
      var configResult = runAllConfigTests();
      overallResults.push({ suite: "環境設定", result: configResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "環境設定", result: error.toString(), status: "失敗" });
    }

    // 2. FileMovementServiceテスト
    Logger.log("\n【2/5】FileMovementServiceテスト実行中...");
    try {
      var fileMovementResult = runAllFileMovementServiceTests();
      overallResults.push({ suite: "FileMovementService", result: fileMovementResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "FileMovementService", result: error.toString(), status: "失敗" });
    }

    // 3. Constantsテスト
    Logger.log("\n【3/5】Constantsテスト実行中...");
    try {
      var constantsResult = runAllConstantsTests();
      overallResults.push({ suite: "Constants", result: constantsResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "Constants", result: error.toString(), status: "失敗" });
    }

    // 4. ConfigManagerテスト
    Logger.log("\n【4/5】ConfigManagerテスト実行中...");
    try {
      var configManagerResult = runAllConfigManagerTests();
      overallResults.push({ suite: "ConfigManager", result: configManagerResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "ConfigManager", result: error.toString(), status: "失敗" });
    }

    // 5. 統合テスト
    Logger.log("\n【5/7】統合テスト実行中...");
    try {
      var integrationResult = runAllRefactoringIntegrationTests();
      overallResults.push({ suite: "統合テスト", result: integrationResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "統合テスト", result: error.toString(), status: "失敗" });
    }

    // 6. Whisperサービス単体テスト
    Logger.log("\n【6/7】Whisperサービステスト実行中...");
    try {
      var whisperResult = runWhisperServiceTests();
      overallResults.push({ suite: "Whisperサービス", result: whisperResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "Whisperサービス", result: error.toString(), status: "失敗" });
    }

    // 7. Zoom単体テスト
    Logger.log("\n【7/9】Zoom単体テスト実行中...");
    try {
      var zoomUnitResult = runZoomUnitTests();
      overallResults.push({ suite: "Zoom単体", result: zoomUnitResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "Zoom単体", result: error.toString(), status: "失敗" });
    }

    // 8. Zoom統合テスト
    Logger.log("\n【8/9】Zoom統合テスト実行中...");
    try {
      var zoomIntegrationResult = runZoomIntegrationTests();
      overallResults.push({ suite: "Zoom統合", result: zoomIntegrationResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "Zoom統合", result: error.toString(), status: "失敗" });
    }

    // 9. Whisper統合テスト
    Logger.log("\n【9/12】Whisper統合テスト実行中...");
    try {
      var whisperIntegrationResult = runWhisperIntegrationTests();
      overallResults.push({ suite: "Whisper統合", result: whisperIntegrationResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "Whisper統合", result: error.toString(), status: "失敗" });
    }

    // 10. ClientMasterDataLoader単体テスト
    Logger.log("\n【10/12】ClientMasterDataLoader単体テスト実行中...");
    try {
      var clientMasterUnitResult = runClientMasterDataLoaderUnitTests();
      overallResults.push({ suite: "ClientMasterDataLoader単体", result: clientMasterUnitResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "ClientMasterDataLoader単体", result: error.toString(), status: "失敗" });
    }

    // 11. ClientMasterDataLoader統合テスト
    Logger.log("\n【11/12】ClientMasterDataLoader統合テスト実行中...");
    try {
      var clientMasterIntegrationResult = runClientMasterDataIntegrationTests();
      overallResults.push({ suite: "ClientMasterDataLoader統合", result: clientMasterIntegrationResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "ClientMasterDataLoader統合", result: error.toString(), status: "失敗" });
    }

    // 12. ClientMasterDataLoader E2Eテスト
    Logger.log("\n【12/14】ClientMasterDataLoader E2Eテスト実行中...");
    try {
      var clientMasterE2EResult = runClientMasterDataE2ETests();
      overallResults.push({ suite: "ClientMasterDataLoader E2E", result: clientMasterE2EResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "ClientMasterDataLoader E2E", result: error.toString(), status: "失敗" });
    }

    // 13. SalesPersonMasterLoader単体テスト
    Logger.log("\n【13/14】SalesPersonMasterLoader単体テスト実行中...");
    try {
      var salesPersonUnitResult = runSalesPersonMasterLoaderUnitTests();
      overallResults.push({ suite: "SalesPersonMasterLoader単体", result: salesPersonUnitResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "SalesPersonMasterLoader単体", result: error.toString(), status: "失敗" });
    }

    // 14. SalesPersonMasterLoader統合テスト
    Logger.log("\n【14/14】SalesPersonMasterLoader統合テスト実行中...");
    try {
      var salesPersonIntegrationResult = runSalesPersonMasterIntegrationTests();
      overallResults.push({ suite: "SalesPersonMasterLoader統合", result: salesPersonIntegrationResult, status: "成功" });
    } catch (error) {
      overallResults.push({ suite: "SalesPersonMasterLoader統合", result: error.toString(), status: "失敗" });
    }

    // 総合結果の出力
    var endTime = new Date();
    var totalTime = (endTime - startTime) / 1000;

    Logger.log("\n========================================");
    Logger.log("全テスト実行完了");
    Logger.log("実行終了時刻: " + endTime.toLocaleString());
    Logger.log("総実行時間: " + totalTime + "秒");
    Logger.log("========================================");

    // 結果サマリー
    var successCount = 0;
    var failureCount = 0;

    Logger.log("\n【総合結果サマリー】");
    for (var i = 0; i < overallResults.length; i++) {
      var result = overallResults[i];
      var statusIcon = result.status === "成功" ? "✓" : "✗";
      Logger.log(statusIcon + " " + result.suite + ": " + result.status);

      if (result.status === "成功") {
        successCount++;
      } else {
        failureCount++;
      }
    }

    Logger.log("\n成功: " + successCount + "件, 失敗: " + failureCount + "件");

    if (failureCount === 0) {
      Logger.log("🎉 全てのテストが成功しました！");
      return "全テスト成功: " + successCount + "件のテストスイートが正常に完了しました";
    } else {
      Logger.log("⚠️  一部のテストで問題が発生しました");
      return "一部テスト失敗: " + successCount + "件成功, " + failureCount + "件失敗";
    }

  } catch (error) {
    Logger.log("マスターテストランナーでエラー: " + error);
    return "テスト実行中に予期しないエラーが発生しました: " + error;
  }
}

/**
 * 特定のテストスイートのみを実行
 * @param {string} suiteName - 実行するテストスイート名
 */
function runSpecificTestSuite(suiteName) {
  try {
    Logger.log("========================================");
    Logger.log("特定テストスイート実行: " + suiteName);
    Logger.log("========================================");

    var result;
    switch (suiteName.toLowerCase()) {
      case "config":
      case "環境設定":
        result = runAllConfigTests();
        break;
      case "filemovement":
      case "fileMovementService":
        result = runAllFileMovementServiceTests();
        break;
      case "constants":
        result = runAllConstantsTests();
        break;
      case "configmanager":
        result = runAllConfigManagerTests();
        break;
      case "integration":
      case "統合":
        result = runAllRefactoringIntegrationTests();
        break;
      case "whisper":
      case "whisperservice":
        result = runWhisperServiceTests();
        break;
      case "whisperintegration":
        result = runWhisperIntegrationTests();
        break;
      case "clientmasterunit":
      case "clientmasterdataunit":
        result = runClientMasterDataLoaderUnitTests();
        break;
      case "clientmasterintegration":
      case "clientmasterdataintegration":
        result = runClientMasterDataIntegrationTests();
        break;
      case "clientmastere2e":
      case "clientmasterdatae2e":
        result = runClientMasterDataE2ETests();
        break;
      case "salesperson":
      case "salespersonunit":
        result = runSalesPersonMasterLoaderUnitTests();
        break;
      case "salespersonintegration":
        result = runSalesPersonMasterIntegrationTests();
        break;
      default:
        throw new Error("不明なテストスイート名: " + suiteName);
    }

    Logger.log("========================================");
    Logger.log("特定テストスイート完了: " + suiteName);
    Logger.log("========================================");

    return result;
  } catch (error) {
    Logger.log("特定テストスイート実行でエラー: " + error);
    return "エラー: " + error.toString();
  }
}

/**
 * 軽量テスト（基本機能のみ）
 */
function runLightweightTests() {
  try {
    Logger.log("========================================");
    Logger.log("軽量テスト実行開始");
    Logger.log("========================================");

    var results = [];

    // 基本的な定数アクセステスト
    Logger.log("1. 基本定数アクセステスト");
    try {
      var status = Constants.STATUS.SUCCESS;
      var sheetName = Constants.SHEET_NAMES.RECORDINGS;
      Logger.log("  ✓ Constants基本アクセス成功");
      results.push("Constants基本アクセス: 成功");
    } catch (e) {
      Logger.log("  ✗ Constants基本アクセス失敗: " + e);
      results.push("Constants基本アクセス: 失敗");
    }

    // 基本的な設定取得テスト
    Logger.log("2. 基本設定取得テスト");
    try {
      var config = ConfigManager.getConfig();
      Logger.log("  ✓ ConfigManager基本取得成功");
      results.push("ConfigManager基本取得: 成功");
    } catch (e) {
      Logger.log("  ✗ ConfigManager基本取得失敗: " + e);
      results.push("ConfigManager基本取得: 失敗");
    }

    // 基本的な結果オブジェクト作成テスト
    Logger.log("3. 基本結果オブジェクト作成テスト");
    try {
      var results_obj = FileMovementService.createResultObject();
      Logger.log("  ✓ FileMovementService基本機能成功");
      results.push("FileMovementService基本機能: 成功");
    } catch (e) {
      Logger.log("  ✗ FileMovementService基本機能失敗: " + e);
      results.push("FileMovementService基本機能: 失敗");
    }

    Logger.log("========================================");
    Logger.log("軽量テスト完了");
    Logger.log("結果: " + results.join(", "));
    Logger.log("========================================");

    return "軽量テスト完了: " + results.length + "項目テスト済み";
  } catch (error) {
    Logger.log("軽量テスト実行でエラー: " + error);
    return "軽量テスト中にエラーが発生しました: " + error;
  }
}

/**
 * テスト環境の健全性チェック
 */
function checkTestEnvironment() {
  try {
    Logger.log("========================================");
    Logger.log("テスト環境健全性チェック開始");
    Logger.log("========================================");

    var checks = [];

    // 必要なモジュールの存在確認
    Logger.log("1. 必要なモジュールの存在確認");
    var requiredModules = [
      { name: "Constants", obj: Constants },
      { name: "ConfigManager", obj: ConfigManager },
      { name: "FileMovementService", obj: FileMovementService },
      { name: "EnvironmentConfig", obj: EnvironmentConfig }
    ];

    for (var i = 0; i < requiredModules.length; i++) {
      var module = requiredModules[i];
      if (typeof module.obj !== 'undefined') {
        Logger.log("  ✓ " + module.name + ": 利用可能");
        checks.push(module.name + ": OK");
      } else {
        Logger.log("  ✗ " + module.name + ": 利用不可");
        checks.push(module.name + ": NG");
      }
    }

    // Google Apps Script APIの利用可能性確認
    Logger.log("2. Google Apps Script API確認");
    try {
      var testDate = new Date();
      Utilities.sleep(1);
      Logger.log("  ✓ 基本API: 利用可能");
      checks.push("基本API: OK");
    } catch (e) {
      Logger.log("  ✗ 基本API: 利用不可 - " + e);
      checks.push("基本API: NG");
    }

    // ログ機能の確認
    Logger.log("3. ログ機能確認");
    try {
      Logger.log("  ✓ Logger: 正常動作");
      checks.push("Logger: OK");
    } catch (e) {
      checks.push("Logger: NG");
    }

    Logger.log("========================================");
    Logger.log("テスト環境健全性チェック完了");
    Logger.log("チェック結果: " + checks.join(", "));
    Logger.log("========================================");

    return "環境チェック完了: " + checks.length + "項目チェック済み";
  } catch (error) {
    Logger.log("テスト環境チェックでエラー: " + error);
    return "環境チェック中にエラーが発生しました: " + error;
  }
}

/**
 * 利用可能なテスト関数の一覧を表示
 */
function showAvailableTests() {
  Logger.log("========================================");
  Logger.log("利用可能なテスト関数一覧");
  Logger.log("========================================");

  Logger.log("【メイン実行関数】");
  Logger.log("• runAllTests() - 全テストスイートを実行");
  Logger.log("• runLightweightTests() - 軽量テスト（基本機能のみ）");
  Logger.log("• checkTestEnvironment() - テスト環境の健全性チェック");

  Logger.log("\n【個別テストスイート】");
  Logger.log("• runAllConfigTests() - 環境設定テスト");
  Logger.log("• runAllFileMovementServiceTests() - FileMovementServiceテスト");
  Logger.log("• runAllConstantsTests() - Constantsテスト");
  Logger.log("• runAllConfigManagerTests() - ConfigManagerテスト");
  Logger.log("• runAllRefactoringIntegrationTests() - 統合テスト");
  Logger.log("• runWhisperServiceTests() - Whisperサービステスト");
  Logger.log("• runWhisperIntegrationTests() - Whisper統合テスト");
  Logger.log("• runClientMasterDataLoaderUnitTests() - ClientMasterDataLoader単体テスト");
  Logger.log("• runClientMasterDataIntegrationTests() - ClientMasterDataLoader統合テスト");
  Logger.log("• runClientMasterDataE2ETests() - ClientMasterDataLoader E2Eテスト");
  Logger.log("• runSalesPersonMasterLoaderUnitTests() - SalesPersonMasterLoader単体テスト");
  Logger.log("• runSalesPersonMasterIntegrationTests() - SalesPersonMasterLoader統合テスト");

  Logger.log("\n【特定テストスイート実行】");
  Logger.log("• runSpecificTestSuite('config') - 環境設定テストのみ");
  Logger.log("• runSpecificTestSuite('constants') - Constantsテストのみ");
  Logger.log("• runSpecificTestSuite('configmanager') - ConfigManagerテストのみ");
  Logger.log("• runSpecificTestSuite('filemovement') - FileMovementServiceテストのみ");
  Logger.log("• runSpecificTestSuite('integration') - 統合テストのみ");
  Logger.log("• runSpecificTestSuite('clientmasterunit') - ClientMasterDataLoader単体テストのみ");
  Logger.log("• runSpecificTestSuite('clientmasterintegration') - ClientMasterDataLoader統合テストのみ");
  Logger.log("• runSpecificTestSuite('clientmastere2e') - ClientMasterDataLoader E2Eテストのみ");
  Logger.log("• runSpecificTestSuite('salesperson') - SalesPersonMasterLoader単体テストのみ");
  Logger.log("• runSpecificTestSuite('salespersonintegration') - SalesPersonMasterLoader統合テストのみ");

  Logger.log("\n【使用例】");
  Logger.log("1. 全テスト実行: runAllTests()");
  Logger.log("2. 軽量テスト: runLightweightTests()");
  Logger.log("3. 環境チェック: checkTestEnvironment()");

  Logger.log("========================================");

  return "テスト関数一覧を表示しました。詳細はログを確認してください。";
}

/**
 * 全テストスイートを実行（優先度中リファクタリングテスト追加版）
 */
function runAllTestsWithMediumPriority() {
  var startTime = new Date();
  Logger.log('=== 全テストスイート実行開始（優先度中リファクタリング含む） ===');
  Logger.log('開始時刻: ' + startTime);

  var allResults = {
    total: 0,
    passed: 0,
    failed: 0,
    suites: []
  };

  var testSuites = [
    { name: 'Environment', func: runEnvironmentTests },
    { name: 'FileMovementService', func: runAllFileMovementServiceTests },
    { name: 'Constants', func: runAllConstantsTests },
    { name: 'ConfigManager', func: runAllConfigManagerTests },
    { name: 'RefactoringIntegration', func: runAllRefactoringIntegrationTests },
    { name: 'MediumPriorityRefactoring', func: runMediumPriorityRefactoringTests }
  ];

  for (var i = 0; i < testSuites.length; i++) {
    var suite = testSuites[i];
    Logger.log('\n--- ' + suite.name + ' テストスイート実行中 ---');

    try {
      var result = suite.func();
      allResults.total += result.total;
      allResults.passed += result.passed;
      allResults.failed += result.failed;
      allResults.suites.push({
        name: suite.name,
        result: result,
        status: 'COMPLETED'
      });

      Logger.log(suite.name + ' テストスイート完了: ' + result.passed + '/' + result.total + ' 成功');
    } catch (error) {
      allResults.failed++;
      allResults.total++;
      allResults.suites.push({
        name: suite.name,
        status: 'ERROR',
        error: error.toString()
      });

      Logger.log(suite.name + ' テストスイートでエラー: ' + error.toString());
    }
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log('\n=== 全テストスイート実行結果（優先度中リファクタリング含む） ===');
  Logger.log('総テスト数: ' + allResults.total);
  Logger.log('成功: ' + allResults.passed);
  Logger.log('失敗: ' + allResults.failed);
  Logger.log('実行時間: ' + duration + '秒');
  Logger.log('成功率: ' + ((allResults.passed / allResults.total) * 100).toFixed(1) + '%');

  // スイート別結果
  Logger.log('\n=== スイート別結果 ===');
  for (var i = 0; i < allResults.suites.length; i++) {
    var suite = allResults.suites[i];
    if (suite.status === 'COMPLETED') {
      var successRate = ((suite.result.passed / suite.result.total) * 100).toFixed(1);
      Logger.log('✓ ' + suite.name + ': ' + suite.result.passed + '/' + suite.result.total + ' (' + successRate + '%)');
    } else {
      Logger.log('✗ ' + suite.name + ': ERROR - ' + suite.error);
    }
  }

  return allResults;
}

/**
 * 環境テストを実行（統一形式）
 */
function runEnvironmentTests() {
  var results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };

  try {
    Logger.log("====== 環境テスト開始 ======");

    // テスト1: 基本環境チェック
    results.total++;
    try {
      var envResult = checkTestEnvironment();
      results.passed++;
      results.details.push({ name: "基本環境チェック", status: "PASS", result: envResult });
      Logger.log("✓ 基本環境チェック: 成功");
    } catch (error) {
      results.failed++;
      results.details.push({ name: "基本環境チェック", status: "FAIL", reason: error.toString() });
      Logger.log("✗ 基本環境チェック: 失敗 - " + error);
    }

    // テスト2: 必要なモジュール存在確認
    results.total++;
    try {
      var requiredModules = [
        { name: "Constants", obj: Constants },
        { name: "ConfigManager", obj: ConfigManager },
        { name: "FileMovementService", obj: FileMovementService },
        { name: "EnvironmentConfig", obj: EnvironmentConfig }
      ];

      var moduleCheckResults = [];
      for (var i = 0; i < requiredModules.length; i++) {
        var module = requiredModules[i];
        if (typeof module.obj !== 'undefined') {
          moduleCheckResults.push(module.name + ": OK");
        } else {
          moduleCheckResults.push(module.name + ": NG");
        }
      }

      results.passed++;
      results.details.push({
        name: "モジュール存在確認",
        status: "PASS",
        result: "モジュールチェック完了: " + moduleCheckResults.join(", ")
      });
      Logger.log("✓ モジュール存在確認: 成功");
    } catch (error) {
      results.failed++;
      results.details.push({ name: "モジュール存在確認", status: "FAIL", reason: error.toString() });
      Logger.log("✗ モジュール存在確認: 失敗 - " + error);
    }

    // テスト3: Google Apps Script API確認
    results.total++;
    try {
      var testDate = new Date();
      Utilities.sleep(1);
      var apiTest = "基本API動作確認完了";

      results.passed++;
      results.details.push({ name: "Google Apps Script API確認", status: "PASS", result: apiTest });
      Logger.log("✓ Google Apps Script API確認: 成功");
    } catch (error) {
      results.failed++;
      results.details.push({ name: "Google Apps Script API確認", status: "FAIL", reason: error.toString() });
      Logger.log("✗ Google Apps Script API確認: 失敗 - " + error);
    }

    Logger.log("====== 環境テスト完了 ======");
    Logger.log("環境テスト結果: " + results.passed + "/" + results.total + " 成功");

    return results;
  } catch (error) {
    Logger.log("環境テスト実行中にエラー: " + error);
    results.failed++;
    results.total++;
    results.details.push({ name: "環境テスト実行", status: "ERROR", reason: error.toString() });
    return results;
  }
} 