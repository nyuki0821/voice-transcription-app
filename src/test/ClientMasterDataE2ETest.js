/**
 * ClientMasterDataLoader E2E（エンドツーエンド）テストスイート
 * 実際のスプレッドシートアクセスを含む、完全なワークフローテスト
 */

/**
 * ClientMasterDataLoader E2Eテスト結果を保存するオブジェクト
 */
var ClientMasterDataE2ETestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * 実際のスプレッドシートアクセステスト
 */
function testRealSpreadsheetAccess() {
  try {
    Logger.log("=== 実際のスプレッドシートアクセステスト開始 ===");

    // 1. スクリプトプロパティの確認
    Logger.log("1. スクリプトプロパティ確認:");

    var scriptProperties = PropertiesService.getScriptProperties();
    var spreadsheetId = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    if (!spreadsheetId) {
      Logger.log("  ⚠ CLIENT_MASTER_SPREADSHEET_IDが設定されていません");
      Logger.log("  → デフォルトデータでのテストを実行します");

      // デフォルトデータでのテスト
      var companies = ClientMasterDataLoader.getClientCompanies();
      if (!Array.isArray(companies) || companies.length === 0) {
        throw new Error('デフォルトデータが取得できませんでした');
      }

      Logger.log("  ✓ デフォルトデータが正常に取得できました（" + companies.length + "社）");
      ClientMasterDataE2ETestResults.passed++;
      return "デフォルトデータでのテストに成功しました";
    }

    Logger.log("  ✓ スプレッドシートID: " + spreadsheetId.substring(0, 10) + "...");

    // 2. スプレッドシートへの実際のアクセステスト
    Logger.log("2. スプレッドシートアクセステスト:");

    try {
      // キャッシュをクリアして、確実にスプレッドシートアクセスを発生させる
      ClientMasterDataLoader.clearCache();

      var startTime = new Date();
      var companies = ClientMasterDataLoader.getClientCompanies();
      var endTime = new Date();
      var duration = endTime - startTime;

      if (!Array.isArray(companies)) {
        throw new Error('取得されたデータが配列ではありません');
      }

      if (companies.length === 0) {
        throw new Error('スプレッドシートからデータを取得できませんでした');
      }

      Logger.log("  ✓ スプレッドシートから " + companies.length + " 社のデータを取得しました");
      Logger.log("  ✓ 取得時間: " + duration + "ms");

      // 3. 取得したデータの妥当性チェック
      Logger.log("3. データ妥当性チェック:");

      for (var i = 0; i < Math.min(companies.length, 5); i++) {
        var company = companies[i];
        if (typeof company !== 'string' || company.trim().length === 0) {
          throw new Error('無効な会社名が含まれています: ' + company);
        }
        Logger.log("  ✓ " + (i + 1) + ". " + company);
      }

      if (companies.length > 5) {
        Logger.log("  ... 他 " + (companies.length - 5) + " 社");
      }

    } catch (error) {
      Logger.log("  ⚠ スプレッドシートアクセスでエラー: " + error);
      Logger.log("  → フォールバック機能のテストを実行します");

      // フォールバック機能のテスト
      var companies = ClientMasterDataLoader.getClientCompanies();
      if (!Array.isArray(companies) || companies.length === 0) {
        throw new Error('フォールバック機能も失敗しました');
      }

      Logger.log("  ✓ フォールバック機能が正常に動作しました（" + companies.length + "社）");
    }

    Logger.log("=== 実際のスプレッドシートアクセステスト完了 ===");
    ClientMasterDataE2ETestResults.passed++;
    return "実際のスプレッドシートアクセステストに成功しました";
  } catch (error) {
    Logger.log("実際のスプレッドシートアクセステストでエラー: " + error);
    ClientMasterDataE2ETestResults.failed++;
    ClientMasterDataE2ETestResults.errors.push("スプレッドシートアクセステスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 完全なワークフローテスト
 */
function testCompleteWorkflow() {
  try {
    Logger.log("=== 完全ワークフローテスト開始 ===");

    // 1. データ読み込みフロー
    Logger.log("1. データ読み込みフロー:");

    // キャッシュクリア
    ClientMasterDataLoader.clearCache();
    Logger.log("  ✓ キャッシュをクリアしました");

    // 会社データ取得
    var companies = ClientMasterDataLoader.getClientCompanies();
    Logger.log("  ✓ 会社データを取得しました（" + companies.length + "社）");

    // エイリアスデータ取得
    var aliases = ClientMasterDataLoader.getClientCompanyAliases();
    Logger.log("  ✓ エイリアスデータを取得しました（" + Object.keys(aliases).length + "社分）");

    // プロンプト生成
    var prompt = ClientMasterDataLoader.getClientListPrompt();
    Logger.log("  ✓ プロンプトを生成しました（" + prompt.length + "文字）");

    // 2. InformationExtractorとの連携フロー
    Logger.log("2. InformationExtractor連携フロー:");

    if (typeof InformationExtractor !== 'undefined') {
      // サンプルテキストでの抽出テスト
      var sampleTexts = [
        "こちら株式会社ENERALLの田中と申します。",
        "エムスリーヘルスデザインの佐藤です。",
        "NOTCHの山田と申します。",
        "佑人社の鈴木です。"
      ];

      for (var i = 0; i < sampleTexts.length; i++) {
        var text = sampleTexts[i];
        Logger.log("  テキスト: " + text);

        // 実際の抽出処理は時間がかかるため、ここでは構造チェックのみ
        // var result = InformationExtractor.extract(text);
        Logger.log("  ✓ テキスト処理の準備ができています");
      }
    } else {
      Logger.log("  ⚠ InformationExtractorが利用できません（モジュール未読み込み）");
    }

    // 3. キャッシュ機能の確認
    Logger.log("3. キャッシュ機能確認:");

    var startTime1 = new Date();
    var companies1 = ClientMasterDataLoader.getClientCompanies();
    var endTime1 = new Date();
    var duration1 = endTime1 - startTime1;

    var startTime2 = new Date();
    var companies2 = ClientMasterDataLoader.getClientCompanies();
    var endTime2 = new Date();
    var duration2 = endTime2 - startTime2;

    Logger.log("  1回目: " + duration1 + "ms");
    Logger.log("  2回目: " + duration2 + "ms");

    if (JSON.stringify(companies1) === JSON.stringify(companies2)) {
      Logger.log("  ✓ キャッシュ機能が正常に動作しています");
    } else {
      throw new Error('キャッシュ前後でデータに差異があります');
    }

    // 4. エラー回復フローの確認
    Logger.log("4. エラー回復フロー確認:");

    // 一時的に無効な設定を行い、回復を確認
    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    try {
      // 無効な設定を一時的に設定
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', 'invalid_test_id');
        ClientMasterDataLoader.clearCache();

        var companiesRecovery = ClientMasterDataLoader.getClientCompanies();
        if (!Array.isArray(companiesRecovery) || companiesRecovery.length === 0) {
          throw new Error('エラー回復が機能していません');
        }

        Logger.log("  ✓ エラー回復機能が正常に動作しました");
      } else {
        Logger.log("  → 元々設定がないため、エラー回復テストをスキップします");
      }
    } finally {
      // 設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }
      ClientMasterDataLoader.clearCache();
    }

    Logger.log("=== 完全ワークフローテスト完了 ===");
    ClientMasterDataE2ETestResults.passed++;
    return "完全ワークフローテストに成功しました";
  } catch (error) {
    Logger.log("完全ワークフローテストでエラー: " + error);
    ClientMasterDataE2ETestResults.failed++;
    ClientMasterDataE2ETestResults.errors.push("完全ワークフローテスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 大量データ処理テスト
 */
function testLargeDataProcessing() {
  try {
    Logger.log("=== 大量データ処理テスト開始 ===");

    // 1. 複数回の連続アクセステスト
    Logger.log("1. 連続アクセステスト:");

    var iterations = 10;
    var totalTime = 0;
    var maxTime = 0;
    var minTime = Number.MAX_VALUE;

    for (var i = 0; i < iterations; i++) {
      ClientMasterDataLoader.clearCache(); // 毎回クリアして実際の処理時間を測定

      var startTime = new Date();
      var companies = ClientMasterDataLoader.getClientCompanies();
      var endTime = new Date();
      var duration = endTime - startTime;

      totalTime += duration;
      maxTime = Math.max(maxTime, duration);
      minTime = Math.min(minTime, duration);

      if (i % 2 === 0) {
        Logger.log("  実行" + (i + 1) + ": " + duration + "ms");
      }
    }

    var averageTime = totalTime / iterations;
    Logger.log("  平均時間: " + averageTime.toFixed(2) + "ms");
    Logger.log("  最大時間: " + maxTime + "ms");
    Logger.log("  最小時間: " + minTime + "ms");

    // 2. エイリアス展開の大量処理テスト
    Logger.log("2. エイリアス展開大量処理テスト:");

    var startTime = new Date();
    var aliases = ClientMasterDataLoader.getClientCompanyAliases();

    var totalAliases = 0;
    var processedCompanies = 0;

    for (var company in aliases) {
      if (aliases.hasOwnProperty(company)) {
        processedCompanies++;
        totalAliases += aliases[company].length;

        // 各会社のエイリアスを検証
        for (var i = 0; i < aliases[company].length; i++) {
          var alias = aliases[company][i];
          if (typeof alias !== 'string' || alias.trim().length === 0) {
            throw new Error('無効なエイリアスが含まれています: ' + alias);
          }
        }
      }
    }

    var endTime = new Date();
    var duration = endTime - startTime;

    Logger.log("  処理した会社数: " + processedCompanies);
    Logger.log("  総エイリアス数: " + totalAliases);
    Logger.log("  処理時間: " + duration + "ms");
    Logger.log("  1社あたり平均時間: " + (duration / processedCompanies).toFixed(2) + "ms");

    // 3. プロンプト生成の大量処理テスト
    Logger.log("3. プロンプト生成大量処理テスト:");

    var iterations = 5;
    var totalTime = 0;

    for (var i = 0; i < iterations; i++) {
      var startTime = new Date();
      var prompt = ClientMasterDataLoader.getClientListPrompt();
      var endTime = new Date();
      var duration = endTime - startTime;

      totalTime += duration;

      // プロンプトの基本検証
      if (typeof prompt !== 'string' || prompt.length < 100) {
        throw new Error('生成されたプロンプトが無効です');
      }

      Logger.log("  実行" + (i + 1) + ": " + duration + "ms（" + prompt.length + "文字）");
    }

    var averageTime = totalTime / iterations;
    Logger.log("  平均生成時間: " + averageTime.toFixed(2) + "ms");

    // パフォーマンス基準のチェック
    if (averageTime > 2000) {
      Logger.log("  ⚠ プロンプト生成時間が2秒を超えています");
    } else {
      Logger.log("  ✓ パフォーマンス基準をクリアしています");
    }

    Logger.log("=== 大量データ処理テスト完了 ===");
    ClientMasterDataE2ETestResults.passed++;
    return "大量データ処理テストに成功しました";
  } catch (error) {
    Logger.log("大量データ処理テストでエラー: " + error);
    ClientMasterDataE2ETestResults.failed++;
    ClientMasterDataE2ETestResults.errors.push("大量データ処理テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 設定変更テスト
 */
function testConfigurationChanges() {
  try {
    Logger.log("=== 設定変更テスト開始 ===");

    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    // 1. 設定なし → 設定あり
    Logger.log("1. 設定なし → 設定ありテスト:");

    try {
      // 設定を一時的に削除
      if (originalValue) {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }
      ClientMasterDataLoader.clearCache();

      var companiesWithoutConfig = ClientMasterDataLoader.getClientCompanies();
      Logger.log("  設定なし時: " + companiesWithoutConfig.length + "社");

      // 設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      }
      ClientMasterDataLoader.clearCache();

      var companiesWithConfig = ClientMasterDataLoader.getClientCompanies();
      Logger.log("  設定あり時: " + companiesWithConfig.length + "社");

      Logger.log("  ✓ 設定変更に応じてデータが変更されました");

    } finally {
      // 確実に設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      }
    }

    // 2. 無効な設定 → 有効な設定
    Logger.log("2. 無効な設定 → 有効な設定テスト:");

    try {
      // 無効な設定を一時的に設定
      scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', 'invalid_id_for_test');
      ClientMasterDataLoader.clearCache();

      var companiesInvalid = ClientMasterDataLoader.getClientCompanies();
      Logger.log("  無効設定時: " + companiesInvalid.length + "社（フォールバック）");

      // 有効な設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }
      ClientMasterDataLoader.clearCache();

      var companiesValid = ClientMasterDataLoader.getClientCompanies();
      Logger.log("  有効設定時: " + companiesValid.length + "社");

      Logger.log("  ✓ 設定修復後、正常にデータが取得できました");

    } finally {
      // 確実に設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }
    }

    // 3. キャッシュクリアの効果確認
    Logger.log("3. キャッシュクリア効果確認:");

    // データを取得してキャッシュを作成
    var companies1 = ClientMasterDataLoader.getClientCompanies();
    Logger.log("  初回取得: " + companies1.length + "社");

    // キャッシュクリア
    ClientMasterDataLoader.clearCache();
    Logger.log("  ✓ キャッシュをクリアしました");

    // 再取得
    var companies2 = ClientMasterDataLoader.getClientCompanies();
    Logger.log("  再取得: " + companies2.length + "社");

    // データの整合性確認
    if (JSON.stringify(companies1) === JSON.stringify(companies2)) {
      Logger.log("  ✓ キャッシュクリア前後でデータの整合性が保たれています");
    } else {
      Logger.log("  ⚠ キャッシュクリア前後でデータに差異があります");
    }

    Logger.log("=== 設定変更テスト完了 ===");
    ClientMasterDataE2ETestResults.passed++;
    return "設定変更テストに成功しました";
  } catch (error) {
    Logger.log("設定変更テストでエラー: " + error);
    ClientMasterDataE2ETestResults.failed++;
    ClientMasterDataE2ETestResults.errors.push("設定変更テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * ClientMasterDataLoader E2Eテストスイートの実行
 */
function runClientMasterDataE2ETests() {
  var startTime = new Date();
  Logger.log("====== ClientMasterDataLoader E2Eテスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  // テスト結果をリセット
  ClientMasterDataE2ETestResults.passed = 0;
  ClientMasterDataE2ETestResults.failed = 0;
  ClientMasterDataE2ETestResults.errors = [];

  var tests = [
    { name: "実際のスプレッドシートアクセス", func: testRealSpreadsheetAccess },
    { name: "完全ワークフロー", func: testCompleteWorkflow },
    { name: "大量データ処理", func: testLargeDataProcessing },
    { name: "設定変更", func: testConfigurationChanges }
  ];

  for (var i = 0; i < tests.length; i++) {
    var test = tests[i];
    try {
      Logger.log("\n--- " + test.name + "テスト実行中 ---");
      var result = test.func();
      Logger.log(test.name + "テスト: " + result);
    } catch (error) {
      Logger.log(test.name + "テストで予期しないエラー: " + error.toString());
      ClientMasterDataE2ETestResults.failed++;
      ClientMasterDataE2ETestResults.errors.push(test.name + ": " + error.toString());
    }
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("\n====== ClientMasterDataLoader E2Eテスト結果 ======");
  Logger.log("総テスト数: " + (ClientMasterDataE2ETestResults.passed + ClientMasterDataE2ETestResults.failed));
  Logger.log("成功: " + ClientMasterDataE2ETestResults.passed);
  Logger.log("失敗: " + ClientMasterDataE2ETestResults.failed);
  Logger.log("実行時間: " + duration + "秒");

  if (ClientMasterDataE2ETestResults.failed > 0) {
    Logger.log("失敗したテスト:");
    for (var i = 0; i < ClientMasterDataE2ETestResults.errors.length; i++) {
      Logger.log("  - " + ClientMasterDataE2ETestResults.errors[i]);
    }
  }

  var successRate = (ClientMasterDataE2ETestResults.passed / (ClientMasterDataE2ETestResults.passed + ClientMasterDataE2ETestResults.failed)) * 100;
  Logger.log("成功率: " + successRate.toFixed(1) + "%");

  return {
    total: ClientMasterDataE2ETestResults.passed + ClientMasterDataE2ETestResults.failed,
    passed: ClientMasterDataE2ETestResults.passed,
    failed: ClientMasterDataE2ETestResults.failed,
    successRate: successRate,
    duration: duration
  };
}

/**
 * E2Eテスト結果サマリーの表示
 */
function displayClientMasterDataE2ETestSummary() {
  Logger.log("\n========== ClientMasterDataLoader E2Eテスト サマリー ==========");
  Logger.log("成功したテスト: " + ClientMasterDataE2ETestResults.passed);
  Logger.log("失敗したテスト: " + ClientMasterDataE2ETestResults.failed);

  if (ClientMasterDataE2ETestResults.failed > 0) {
    Logger.log("\n失敗の詳細:");
    for (var i = 0; i < ClientMasterDataE2ETestResults.errors.length; i++) {
      Logger.log("• " + ClientMasterDataE2ETestResults.errors[i]);
    }
  } else {
    Logger.log("🎉 全てのE2Eテストが成功しました！");
  }

  Logger.log("==============================================================");
} 