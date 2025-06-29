/**
 * ClientMasterDataLoader統合テストスイート
 * ClientMasterDataLoaderと他のモジュール（InformationExtractor等）との連携テスト
 */

/**
 * ClientMasterDataLoader統合テスト結果を保存するオブジェクト
 */
var ClientMasterDataIntegrationTestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * InformationExtractorとの統合テスト
 */
function testInformationExtractorIntegration() {
  try {
    Logger.log("=== InformationExtractor統合テスト開始 ===");

    // 1. InformationExtractorが存在するかチェック
    if (typeof InformationExtractor === 'undefined') {
      throw new Error('InformationExtractorが定義されていません');
    }

    // 2. validateSalesCompany関数のテスト
    Logger.log("1. validateSalesCompany関数テスト:");

    // テストケース: 正確な会社名
    var testCases = [
      { input: "株式会社ENERALL", expected: "株式会社ENERALL" },
      { input: "ENERALL", expected: "株式会社ENERALL" },
      { input: "エネラル", expected: "株式会社ENERALL" },
      { input: "NOTCH", expected: "株式会社NOTCH" },
      { input: "ノッチ", expected: "株式会社NOTCH" },
      { input: "エムスリーヘルス", expected: "エムスリーヘルスデザイン株式会社" },
      { input: "存在しない会社", expected: "" }
    ];

    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];

      // InformationExtractorのvalidateSalesCompany関数は直接アクセスできないため、
      // extractメソッド経由でテストするか、モック的にテストする
      Logger.log("  テストケース: '" + testCase.input + "' -> 期待値: '" + testCase.expected + "'");

      // 実際のテストは、ClientMasterDataLoaderのデータが正しく使用されているかを確認
      var companies = ClientMasterDataLoader.getClientCompanies();
      var aliases = ClientMasterDataLoader.getClientCompanyAliases();

      var found = false;
      for (var company in aliases) {
        if (aliases.hasOwnProperty(company)) {
          var companyAliases = aliases[company];
          if (companyAliases.indexOf(testCase.input) !== -1) {
            found = true;
            if (company === testCase.expected) {
              Logger.log("  ✓ '" + testCase.input + "' が正しく '" + company + "' にマッピングされています");
            } else {
              Logger.log("  ⚠ '" + testCase.input + "' のマッピングが期待値と異なります");
            }
            break;
          }
        }
      }

      if (!found && testCase.expected === "") {
        Logger.log("  ✓ '" + testCase.input + "' は正しく認識されていません（期待通り）");
      } else if (!found && testCase.expected !== "") {
        Logger.log("  ⚠ '" + testCase.input + "' が見つかりませんでした");
      }
    }

    // 3. プロンプト生成の統合テスト
    Logger.log("2. プロンプト生成統合テスト:");
    var prompt = ClientMasterDataLoader.getClientListPrompt();

    if (prompt.length < 100) {
      throw new Error('生成されたプロンプトが短すぎます');
    }

    // プロンプトに必要な要素が含まれているかチェック
    var requiredElements = [
      "営業会社名は以下のリストから",
      "株式会社ENERALL",
      "エムスリーヘルスデザイン株式会社"
    ];

    for (var i = 0; i < requiredElements.length; i++) {
      var element = requiredElements[i];
      if (prompt.indexOf(element) === -1) {
        throw new Error('プロンプトに必要な要素が含まれていません: ' + element);
      }
    }

    Logger.log("  ✓ プロンプトが正しく生成されています（" + prompt.length + "文字）");

    Logger.log("=== InformationExtractor統合テスト完了 ===");
    ClientMasterDataIntegrationTestResults.passed++;
    return "InformationExtractor統合テストに成功しました";
  } catch (error) {
    Logger.log("InformationExtractor統合テストでエラー: " + error);
    ClientMasterDataIntegrationTestResults.failed++;
    ClientMasterDataIntegrationTestResults.errors.push("InformationExtractor統合テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * データ整合性テスト
 */
function testDataConsistency() {
  try {
    Logger.log("=== データ整合性テスト開始 ===");

    // 1. 会社リストとエイリアスの整合性チェック
    var companies = ClientMasterDataLoader.getClientCompanies();
    var aliases = ClientMasterDataLoader.getClientCompanyAliases();

    Logger.log("1. 会社リストとエイリアスの整合性:");

    for (var i = 0; i < companies.length; i++) {
      var company = companies[i];

      if (!aliases[company]) {
        throw new Error('会社 "' + company + '" のエイリアスが定義されていません');
      }

      if (!Array.isArray(aliases[company])) {
        throw new Error('会社 "' + company + '" のエイリアスが配列ではありません');
      }

      if (aliases[company].length === 0) {
        throw new Error('会社 "' + company + '" のエイリアスが空です');
      }

      // 会社名自体がエイリアスに含まれているかチェック
      if (aliases[company].indexOf(company) === -1) {
        throw new Error('会社 "' + company + '" の正式名称がエイリアスに含まれていません');
      }

      Logger.log("  ✓ " + company + ": " + aliases[company].length + "個のエイリアス");
    }

    // 2. 重複チェック
    Logger.log("2. 会社名重複チェック:");

    for (var i = 0; i < companies.length; i++) {
      for (var j = i + 1; j < companies.length; j++) {
        if (companies[i] === companies[j]) {
          throw new Error('重複した会社名が見つかりました: ' + companies[i]);
        }
      }
    }

    Logger.log("  ✓ 会社名に重複はありません");

    // 3. エイリアスの逆引きテスト
    Logger.log("3. エイリアス逆引きテスト:");

    var aliasToCompany = {};
    for (var company in aliases) {
      if (aliases.hasOwnProperty(company)) {
        var companyAliases = aliases[company];
        for (var i = 0; i < companyAliases.length; i++) {
          var alias = companyAliases[i];
          if (aliasToCompany[alias] && aliasToCompany[alias] !== company) {
            Logger.log("  ⚠ エイリアス '" + alias + "' が複数の会社で使用されています: " +
              aliasToCompany[alias] + " と " + company);
          } else {
            aliasToCompany[alias] = company;
          }
        }
      }
    }

    Logger.log("  ✓ エイリアス逆引きマップを作成しました（" + Object.keys(aliasToCompany).length + "個）");

    Logger.log("=== データ整合性テスト完了 ===");
    ClientMasterDataIntegrationTestResults.passed++;
    return "データ整合性テストに成功しました";
  } catch (error) {
    Logger.log("データ整合性テストでエラー: " + error);
    ClientMasterDataIntegrationTestResults.failed++;
    ClientMasterDataIntegrationTestResults.errors.push("データ整合性テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * パフォーマンステスト
 */
function testPerformance() {
  try {
    Logger.log("=== パフォーマンステスト開始 ===");

    // 1. 複数回呼び出しのパフォーマンステスト
    Logger.log("1. 複数回呼び出しテスト:");

    var iterations = 5;
    var totalTime = 0;
    var results = [];

    for (var i = 0; i < iterations; i++) {
      var startTime = new Date();
      var companies = ClientMasterDataLoader.getClientCompanies();
      var endTime = new Date();
      var duration = endTime - startTime;

      totalTime += duration;
      results.push(duration);

      Logger.log("  実行" + (i + 1) + ": " + duration + "ms（" + companies.length + "社）");
    }

    var averageTime = totalTime / iterations;
    Logger.log("  平均実行時間: " + averageTime.toFixed(2) + "ms");

    // 2. 大量データ処理テスト（エイリアス展開）
    Logger.log("2. エイリアス展開パフォーマンステスト:");

    var startTime = new Date();
    var aliases = ClientMasterDataLoader.getClientCompanyAliases();
    var totalAliases = 0;

    for (var company in aliases) {
      if (aliases.hasOwnProperty(company)) {
        totalAliases += aliases[company].length;
      }
    }

    var endTime = new Date();
    var duration = endTime - startTime;

    Logger.log("  総エイリアス数: " + totalAliases);
    Logger.log("  処理時間: " + duration + "ms");

    // 3. プロンプト生成パフォーマンステスト
    Logger.log("3. プロンプト生成パフォーマンステスト:");

    var startTime = new Date();
    var prompt = ClientMasterDataLoader.getClientListPrompt();
    var endTime = new Date();
    var duration = endTime - startTime;

    Logger.log("  プロンプト長: " + prompt.length + "文字");
    Logger.log("  生成時間: " + duration + "ms");

    // パフォーマンス基準チェック（1秒以内）
    if (averageTime > 1000) {
      Logger.log("  ⚠ 平均実行時間が1秒を超えています");
    } else {
      Logger.log("  ✓ パフォーマンス基準をクリアしています");
    }

    Logger.log("=== パフォーマンステスト完了 ===");
    ClientMasterDataIntegrationTestResults.passed++;
    return "パフォーマンステストに成功しました";
  } catch (error) {
    Logger.log("パフォーマンステストでエラー: " + error);
    ClientMasterDataIntegrationTestResults.failed++;
    ClientMasterDataIntegrationTestResults.errors.push("パフォーマンステスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 下位互換性テスト
 */
function testBackwardCompatibility() {
  try {
    Logger.log("=== 下位互換性テスト開始 ===");

    // 1. 既存の会社名が全て含まれているかテスト
    Logger.log("1. 既存会社名包含テスト:");

    var expectedCompanies = [
      "株式会社ENERALL",
      "エムスリーヘルスデザイン株式会社",
      "株式会社TOKIUM",
      "株式会社グッドワークス",
      "テコム看護",
      "ハローワールド株式会社",
      "株式会社ワーサル",
      "株式会社NOTCH",
      "株式会社ジースタイラス",
      "株式会社佑人社",
      "株式会社リディラバ",
      "株式会社インフィニットマインド"
    ];

    var companies = ClientMasterDataLoader.getClientCompanies();

    for (var i = 0; i < expectedCompanies.length; i++) {
      var expected = expectedCompanies[i];
      if (companies.indexOf(expected) === -1) {
        throw new Error('期待される既存会社名が見つかりません: ' + expected);
      }
      Logger.log("  ✓ " + expected + " が含まれています");
    }

    // 2. 既存のエイリアスが機能するかテスト
    Logger.log("2. 既存エイリアステスト:");

    var aliasTests = [
      { alias: "ENERALL", expectedCompany: "株式会社ENERALL" },
      { alias: "エネラル", expectedCompany: "株式会社ENERALL" },
      { alias: "NOTCH", expectedCompany: "株式会社NOTCH" },
      { alias: "ノッチ", expectedCompany: "株式会社NOTCH" },
      { alias: "エムスリーヘルス", expectedCompany: "エムスリーヘルスデザイン株式会社" },
      { alias: "佑人社", expectedCompany: "株式会社佑人社" },
      { alias: "ゆうじんしゃ", expectedCompany: "株式会社佑人社" }
    ];

    var aliases = ClientMasterDataLoader.getClientCompanyAliases();

    for (var i = 0; i < aliasTests.length; i++) {
      var test = aliasTests[i];
      var found = false;

      if (aliases[test.expectedCompany] &&
        aliases[test.expectedCompany].indexOf(test.alias) !== -1) {
        found = true;
        Logger.log("  ✓ エイリアス '" + test.alias + "' -> '" + test.expectedCompany + "'");
      }

      if (!found) {
        Logger.log("  ⚠ エイリアス '" + test.alias + "' が見つかりません");
      }
    }

    // 3. プロンプトフォーマット互換性テスト
    Logger.log("3. プロンプトフォーマット互換性テスト:");

    var prompt = ClientMasterDataLoader.getClientListPrompt();

    // 既存のプロンプトで期待される要素
    var expectedElements = [
      "営業会社名は以下のリストから最も適切なものを選んでください",
      "会話は基本的にこのリストの会社のいずれかから行われています",
      "営業担当者が自社名を名乗っている部分を特定し",
      "該当する会社を選択してください"
    ];

    for (var i = 0; i < expectedElements.length; i++) {
      var element = expectedElements[i];
      if (prompt.indexOf(element) === -1) {
        throw new Error('プロンプトに期待される要素が含まれていません: ' + element);
      }
      Logger.log("  ✓ 期待される要素が含まれています");
    }

    Logger.log("=== 下位互換性テスト完了 ===");
    ClientMasterDataIntegrationTestResults.passed++;
    return "下位互換性テストに成功しました";
  } catch (error) {
    Logger.log("下位互換性テストでエラー: " + error);
    ClientMasterDataIntegrationTestResults.failed++;
    ClientMasterDataIntegrationTestResults.errors.push("下位互換性テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * エラー回復テスト
 */
function testErrorRecovery() {
  try {
    Logger.log("=== エラー回復テスト開始 ===");

    // 1. スプレッドシートアクセスエラー時のフォールバック
    Logger.log("1. フォールバック機能テスト:");

    // キャッシュをクリアして、確実にloadClientDataが呼ばれるようにする
    ClientMasterDataLoader.clearCache();

    // スクリプトプロパティを一時的に無効な値に設定
    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    try {
      // 無効なスプレッドシートIDを設定
      scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', 'invalid_id_12345');

      // この状態でデータ取得を試行（フォールバックが動作するはず）
      var companies = ClientMasterDataLoader.getClientCompanies();

      if (!Array.isArray(companies) || companies.length === 0) {
        throw new Error('フォールバック時にデフォルトデータが取得できませんでした');
      }

      Logger.log("  ✓ 無効なスプレッドシートID時もデフォルトデータが取得できました（" + companies.length + "社）");

    } finally {
      // 設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }

      // キャッシュをクリアして正常状態に戻す
      ClientMasterDataLoader.clearCache();
    }

    // 2. 部分的なデータ破損への対応
    Logger.log("2. データ破損対応テスト:");

    // 正常なデータが取得できることを確認
    var companies = ClientMasterDataLoader.getClientCompanies();
    var aliases = ClientMasterDataLoader.getClientCompanyAliases();

    if (companies.length === 0) {
      throw new Error('会社データが空です');
    }

    if (Object.keys(aliases).length === 0) {
      throw new Error('エイリアスデータが空です');
    }

    Logger.log("  ✓ データが正常に復旧されています");

    Logger.log("=== エラー回復テスト完了 ===");
    ClientMasterDataIntegrationTestResults.passed++;
    return "エラー回復テストに成功しました";
  } catch (error) {
    Logger.log("エラー回復テストでエラー: " + error);
    ClientMasterDataIntegrationTestResults.failed++;
    ClientMasterDataIntegrationTestResults.errors.push("エラー回復テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * ClientMasterDataLoader統合テストスイートの実行
 */
function runClientMasterDataIntegrationTests() {
  var startTime = new Date();
  Logger.log("====== ClientMasterDataLoader統合テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  // テスト結果をリセット
  ClientMasterDataIntegrationTestResults.passed = 0;
  ClientMasterDataIntegrationTestResults.failed = 0;
  ClientMasterDataIntegrationTestResults.errors = [];

  var tests = [
    { name: "InformationExtractor統合", func: testInformationExtractorIntegration },
    { name: "データ整合性", func: testDataConsistency },
    { name: "パフォーマンス", func: testPerformance },
    { name: "下位互換性", func: testBackwardCompatibility },
    { name: "エラー回復", func: testErrorRecovery }
  ];

  for (var i = 0; i < tests.length; i++) {
    var test = tests[i];
    try {
      Logger.log("\n--- " + test.name + "テスト実行中 ---");
      var result = test.func();
      Logger.log(test.name + "テスト: " + result);
    } catch (error) {
      Logger.log(test.name + "テストで予期しないエラー: " + error.toString());
      ClientMasterDataIntegrationTestResults.failed++;
      ClientMasterDataIntegrationTestResults.errors.push(test.name + ": " + error.toString());
    }
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("\n====== ClientMasterDataLoader統合テスト結果 ======");
  Logger.log("総テスト数: " + (ClientMasterDataIntegrationTestResults.passed + ClientMasterDataIntegrationTestResults.failed));
  Logger.log("成功: " + ClientMasterDataIntegrationTestResults.passed);
  Logger.log("失敗: " + ClientMasterDataIntegrationTestResults.failed);
  Logger.log("実行時間: " + duration + "秒");

  if (ClientMasterDataIntegrationTestResults.failed > 0) {
    Logger.log("失敗したテスト:");
    for (var i = 0; i < ClientMasterDataIntegrationTestResults.errors.length; i++) {
      Logger.log("  - " + ClientMasterDataIntegrationTestResults.errors[i]);
    }
  }

  var successRate = (ClientMasterDataIntegrationTestResults.passed / (ClientMasterDataIntegrationTestResults.passed + ClientMasterDataIntegrationTestResults.failed)) * 100;
  Logger.log("成功率: " + successRate.toFixed(1) + "%");

  return {
    total: ClientMasterDataIntegrationTestResults.passed + ClientMasterDataIntegrationTestResults.failed,
    passed: ClientMasterDataIntegrationTestResults.passed,
    failed: ClientMasterDataIntegrationTestResults.failed,
    successRate: successRate,
    duration: duration
  };
}

/**
 * 統合テスト結果サマリーの表示
 */
function displayClientMasterDataIntegrationTestSummary() {
  Logger.log("\n========== ClientMasterDataLoader統合テスト サマリー ==========");
  Logger.log("成功したテスト: " + ClientMasterDataIntegrationTestResults.passed);
  Logger.log("失敗したテスト: " + ClientMasterDataIntegrationTestResults.failed);

  if (ClientMasterDataIntegrationTestResults.failed > 0) {
    Logger.log("\n失敗の詳細:");
    for (var i = 0; i < ClientMasterDataIntegrationTestResults.errors.length; i++) {
      Logger.log("• " + ClientMasterDataIntegrationTestResults.errors[i]);
    }
  } else {
    Logger.log("🎉 全ての統合テストが成功しました！");
  }

  Logger.log("================================================================");
} 