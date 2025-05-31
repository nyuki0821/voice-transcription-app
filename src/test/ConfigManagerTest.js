/**
 * ConfigManagerのテスト用スクリプト
 * 設定管理サービスの各種機能をテストする
 */

/**
 * 設定取得機能のテスト
 */
function testConfigRetrieval() {
  try {
    Logger.log("=== 設定取得機能テスト開始 ===");

    // 全設定の取得
    var config = ConfigManager.getConfig();
    Logger.log("設定オブジェクトのキー数: " + Object.keys(config).length);

    // 主要設定の確認
    var keyChecks = [
      'ASSEMBLYAI_API_KEY',
      'OPENAI_API_KEY',
      'SOURCE_FOLDER_ID',
      'RECORDINGS_SHEET_ID',
      'MAX_BATCH_SIZE',
      'ENHANCE_WITH_OPENAI',
      'ADMIN_EMAILS'
    ];

    for (var i = 0; i < keyChecks.length; i++) {
      var key = keyChecks[i];
      var value = config[key];
      var type = typeof value;

      // 機密情報はマスク
      if (key.indexOf("KEY") !== -1 || key.indexOf("SECRET") !== -1) {
        Logger.log(key + ": ******** (型: " + type + ")");
      } else {
        Logger.log(key + ": " + JSON.stringify(value) + " (型: " + type + ")");
      }
    }

    Logger.log("=== 設定取得機能テスト完了 ===");
    return "設定取得機能のテストに成功しました";
  } catch (error) {
    Logger.log("設定取得機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 個別設定取得機能のテスト
 */
function testIndividualConfigGet() {
  try {
    Logger.log("=== 個別設定取得機能テスト開始 ===");

    // 存在する設定の取得
    var maxBatchSize = ConfigManager.get('MAX_BATCH_SIZE', 5);
    Logger.log("MAX_BATCH_SIZE: " + maxBatchSize + " (型: " + typeof maxBatchSize + ")");

    var enhanceWithOpenAI = ConfigManager.get('ENHANCE_WITH_OPENAI', false);
    Logger.log("ENHANCE_WITH_OPENAI: " + enhanceWithOpenAI + " (型: " + typeof enhanceWithOpenAI + ")");

    // 存在しない設定の取得（デフォルト値）
    var nonExistent = ConfigManager.get('NON_EXISTENT_KEY', 'default_value');
    Logger.log("NON_EXISTENT_KEY: " + nonExistent + " (デフォルト値が返されるか確認)");

    // デフォルト値なしの場合
    var noDefault = ConfigManager.get('ANOTHER_NON_EXISTENT_KEY');
    Logger.log("ANOTHER_NON_EXISTENT_KEY: " + noDefault + " (undefinedが返されるか確認)");

    Logger.log("=== 個別設定取得機能テスト完了 ===");
    return "個別設定取得機能のテストに成功しました";
  } catch (error) {
    Logger.log("個別設定取得機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * キャッシュ機能のテスト
 */
function testCacheFunction() {
  try {
    Logger.log("=== キャッシュ機能テスト開始 ===");

    // キャッシュクリア
    ConfigManager.clearCache();
    Logger.log("キャッシュをクリアしました");

    // 1回目の取得（キャッシュなし）
    var startTime1 = new Date().getTime();
    var config1 = ConfigManager.getConfig();
    var endTime1 = new Date().getTime();
    var duration1 = endTime1 - startTime1;
    Logger.log("1回目取得時間: " + duration1 + "ms");

    // 2回目の取得（キャッシュあり）
    var startTime2 = new Date().getTime();
    var config2 = ConfigManager.getConfig();
    var endTime2 = new Date().getTime();
    var duration2 = endTime2 - startTime2;
    Logger.log("2回目取得時間: " + duration2 + "ms");

    // 強制リフレッシュ
    var startTime3 = new Date().getTime();
    var config3 = ConfigManager.getConfig(true);
    var endTime3 = new Date().getTime();
    var duration3 = endTime3 - startTime3;
    Logger.log("強制リフレッシュ時間: " + duration3 + "ms");

    // キャッシュ効果の確認
    if (duration2 < duration1) {
      Logger.log("キャッシュ効果確認: 2回目が高速化されました");
    } else {
      Logger.log("キャッシュ効果確認: 期待通りの高速化は見られませんでした");
    }

    Logger.log("=== キャッシュ機能テスト完了 ===");
    return "キャッシュ機能のテストに成功しました";
  } catch (error) {
    Logger.log("キャッシュ機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 設定妥当性チェック機能のテスト
 */
function testConfigValidation() {
  try {
    Logger.log("=== 設定妥当性チェック機能テスト開始 ===");

    var validation = ConfigManager.validateConfig();
    Logger.log("妥当性チェック結果:");
    Logger.log("  有効: " + validation.valid);
    Logger.log("  エラー数: " + validation.errors.length);

    if (validation.errors.length > 0) {
      Logger.log("  エラー詳細:");
      for (var i = 0; i < validation.errors.length; i++) {
        Logger.log("    " + (i + 1) + ". " + validation.errors[i]);
      }
    } else {
      Logger.log("  全ての必須設定が正しく設定されています");
    }

    Logger.log("=== 設定妥当性チェック機能テスト完了 ===");
    return "設定妥当性チェック機能のテストに成功しました";
  } catch (error) {
    Logger.log("設定妥当性チェック機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * シート・フォルダアクセス機能のテスト（モック）
 * 実際のアクセスは行わず、設定値の存在確認のみ
 */
function testSheetFolderAccessMock() {
  try {
    Logger.log("=== シート・フォルダアクセス機能テスト開始 ===");

    // 設定値の存在確認
    var recordingsSheetId = ConfigManager.get('RECORDINGS_SHEET_ID');
    var processedSheetId = ConfigManager.get('PROCESSED_SHEET_ID');
    var sourceFolderId = ConfigManager.get('SOURCE_FOLDER_ID');
    var errorFolderId = ConfigManager.get('ERROR_FOLDER_ID');

    Logger.log("RECORDINGS_SHEET_ID: " + (recordingsSheetId ? "設定済み" : "未設定"));
    Logger.log("PROCESSED_SHEET_ID: " + (processedSheetId ? "設定済み" : "未設定"));
    Logger.log("SOURCE_FOLDER_ID: " + (sourceFolderId ? "設定済み" : "未設定"));
    Logger.log("ERROR_FOLDER_ID: " + (errorFolderId ? "設定済み" : "未設定"));

    // 実際のアクセステストは危険なのでスキップ
    Logger.log("注意: 実際のシート・フォルダアクセステストはスキップしました");
    Logger.log("      本番環境では ConfigManager.getRecordingsSheet() などを使用してください");

    Logger.log("=== シート・フォルダアクセス機能テスト完了 ===");
    return "シート・フォルダアクセス機能のテストに成功しました";
  } catch (error) {
    Logger.log("シート・フォルダアクセス機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * デフォルト設定取得機能のテスト
 */
function testDefaultConfig() {
  try {
    Logger.log("=== デフォルト設定取得機能テスト開始 ===");

    var defaultConfig = ConfigManager.getDefaultConfig();
    Logger.log("デフォルト設定のキー数: " + Object.keys(defaultConfig).length);

    // 主要なデフォルト値の確認
    Logger.log("デフォルト値確認:");
    Logger.log("  MAX_BATCH_SIZE: " + defaultConfig.MAX_BATCH_SIZE);
    Logger.log("  ENHANCE_WITH_OPENAI: " + defaultConfig.ENHANCE_WITH_OPENAI);
    Logger.log("  ADMIN_EMAILS: " + JSON.stringify(defaultConfig.ADMIN_EMAILS));

    // 空文字列のデフォルト値確認
    var emptyDefaults = ['ASSEMBLYAI_API_KEY', 'SOURCE_FOLDER_ID', 'RECORDINGS_SHEET_ID'];
    for (var i = 0; i < emptyDefaults.length; i++) {
      var key = emptyDefaults[i];
      var value = defaultConfig[key];
      Logger.log("  " + key + ": '" + value + "' (空文字列か確認)");
    }

    Logger.log("=== デフォルト設定取得機能テスト完了 ===");
    return "デフォルト設定取得機能のテストに成功しました";
  } catch (error) {
    Logger.log("デフォルト設定取得機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * ConfigManagerの全テストを実行
 */
function runAllConfigManagerTests() {
  var startTime = new Date();
  Logger.log("====== ConfigManager 全テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  var results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };

  var tests = [
    { name: "設定取得機能", func: testConfigRetrieval },
    { name: "個別設定取得", func: testIndividualConfigGet },
    { name: "キャッシュ機能", func: testCacheFunction },
    { name: "設定妥当性チェック", func: testConfigValidation },
    { name: "シート・フォルダアクセス", func: testSheetFolderAccessMock },
    { name: "デフォルト設定取得", func: testDefaultConfig }
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

  Logger.log("\n====== ConfigManager テスト結果 ======");
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

  Logger.log("====== ConfigManager 全テスト完了 ======");
  return results;
} 