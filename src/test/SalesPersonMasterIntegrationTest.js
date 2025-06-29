/**
 * SalesPersonMasterLoader統合テスト
 * TranscriptionServiceとの連携や実際の使用シナリオをテストする
 */

// テスト結果を格納するオブジェクト
var SalesPersonMasterIntegrationTestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * SalesPersonMasterLoader統合テストスイート
 */
function runSalesPersonMasterIntegrationTests() {
  var startTime = new Date();
  Logger.log("====== SalesPersonMasterLoader統合テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  // テスト結果をリセット
  SalesPersonMasterIntegrationTestResults.passed = 0;
  SalesPersonMasterIntegrationTestResults.failed = 0;
  SalesPersonMasterIntegrationTestResults.errors = [];

  var tests = [
    { name: "TranscriptionService統合", func: testTranscriptionServiceIntegration },
    { name: "名前正規化統合", func: testNameNormalizationIntegration },
    { name: "データ整合性", func: testSalesPersonDataConsistency },
    { name: "パフォーマンス", func: testSalesPersonPerformance },
    { name: "エラー回復", func: testSalesPersonErrorRecovery }
  ];

  for (var i = 0; i < tests.length; i++) {
    var test = tests[i];
    try {
      Logger.log("\n--- " + test.name + "テスト実行中 ---");
      var result = test.func();
      Logger.log(test.name + "テスト: " + result);
    } catch (error) {
      Logger.log(test.name + "テストで予期しないエラー: " + error.toString());
      SalesPersonMasterIntegrationTestResults.failed++;
      SalesPersonMasterIntegrationTestResults.errors.push(test.name + ": " + error.toString());
    }
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("\n====== SalesPersonMasterLoader統合テスト結果 ======");
  Logger.log("総テスト数: " + (SalesPersonMasterIntegrationTestResults.passed + SalesPersonMasterIntegrationTestResults.failed));
  Logger.log("成功: " + SalesPersonMasterIntegrationTestResults.passed);
  Logger.log("失敗: " + SalesPersonMasterIntegrationTestResults.failed);
  Logger.log("実行時間: " + duration + "秒");

  if (SalesPersonMasterIntegrationTestResults.failed > 0) {
    Logger.log("失敗したテスト:");
    for (var i = 0; i < SalesPersonMasterIntegrationTestResults.errors.length; i++) {
      Logger.log("  - " + SalesPersonMasterIntegrationTestResults.errors[i]);
    }
  }

  var successRate = (SalesPersonMasterIntegrationTestResults.passed / (SalesPersonMasterIntegrationTestResults.passed + SalesPersonMasterIntegrationTestResults.failed)) * 100;
  Logger.log("成功率: " + successRate.toFixed(1) + "%");

  return {
    total: SalesPersonMasterIntegrationTestResults.passed + SalesPersonMasterIntegrationTestResults.failed,
    passed: SalesPersonMasterIntegrationTestResults.passed,
    failed: SalesPersonMasterIntegrationTestResults.failed,
    successRate: successRate,
    duration: duration
  };
}

/**
 * TranscriptionServiceとの統合テスト
 */
function testTranscriptionServiceIntegration() {
  try {
    Logger.log("=== TranscriptionService統合テスト開始 ===");

    // 1. normalizeSalesPersonNamesInText関数のテスト
    Logger.log("1. normalizeSalesPersonNamesInText関数テスト:");

    var testTexts = [
      {
        input: "【株式会社ENERALL 高野】 お世話になっております。",
        expected: "【株式会社ENERALL 高野 仁】"
      },
      {
        input: "【株式会社NOTCH たかの】 ご連絡ありがとうございます。",
        expected: "【株式会社NOTCH 高野 仁】"
      },
      {
        input: "【エムスリーヘルスデザイン株式会社 枡田】 確認いたします。",
        expected: "【エムスリーヘルスデザイン株式会社 桝田 明良】"
      },
      {
        input: "【株式会社佑人社 さいとう】 承知いたしました。",
        expected: "【株式会社佑人社 齋藤 茜】"
      }
    ];

    var passedCount = 0;
    for (var i = 0; i < testTexts.length; i++) {
      var test = testTexts[i];
      var result = TranscriptionService.normalizeSalesPersonNamesInText(test.input);

      if (result.indexOf(test.expected) !== -1) {
        Logger.log("  ✓ 正規化成功: " + test.expected);
        passedCount++;
      } else {
        Logger.log("  ✗ 正規化失敗: 期待値 '" + test.expected + "' が結果に含まれていません");
        Logger.log("    結果: " + result);
      }
    }

    if (passedCount !== testTexts.length) {
      throw new Error(passedCount + "/" + testTexts.length + " のテストケースのみ成功");
    }

    // 2. 複数行テキストの処理
    Logger.log("2. 複数行テキストの処理テスト:");

    var multilineText =
      "【株式会社ENERALL 高野】 お世話になっております。\n" +
      "本日はお時間をいただきありがとうございます。\n\n" +
      "【顧客会社 担当者】 こちらこそよろしくお願いします。\n\n" +
      "【株式会社ENERALL たかの】 それでは早速ですが...";

    var normalizedMultiline = TranscriptionService.normalizeSalesPersonNamesInText(multilineText);

    if (normalizedMultiline.indexOf("【株式会社ENERALL 高野 仁】") !== -1) {
      Logger.log("  ✓ 複数行テキストでの正規化が正常に動作しています");
    } else {
      throw new Error("複数行テキストでの正規化が失敗しました");
    }

    Logger.log("=== TranscriptionService統合テスト完了 ===");
    SalesPersonMasterIntegrationTestResults.passed++;
    return "TranscriptionService統合テストに成功しました";
  } catch (error) {
    Logger.log("TranscriptionService統合テストでエラー: " + error);
    SalesPersonMasterIntegrationTestResults.failed++;
    SalesPersonMasterIntegrationTestResults.errors.push("TranscriptionService統合テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 名前正規化統合テスト
 */
function testNameNormalizationIntegration() {
  try {
    Logger.log("=== 名前正規化統合テスト開始 ===");

    // 1. 様々なパターンでの正規化テスト
    Logger.log("1. 複雑なパターンでの正規化:");

    var complexPatterns = [
      // 姓の異体字
      { input: "斎藤", expected: "齋藤 茜" },
      { input: "斉藤", expected: "齋藤 茜" },
      { input: "齋藤", expected: "齋藤 茜" },

      // 大文字小文字混在
      { input: "Takano", expected: "高野 仁" },
      { input: "MASUDA", expected: "桝田 明良" },

      // 部分一致
      { input: "高野さん", expected: "高野 仁" },
      { input: "本間君", expected: "本間君" }, // 「君」は登録されていないので変換されない

      // スペースのバリエーション
      { input: "高野　仁", expected: "高野 仁" },
      { input: "高野  仁", expected: "高野 仁" }
    ];

    var passedCount = 0;
    for (var i = 0; i < complexPatterns.length; i++) {
      var pattern = complexPatterns[i];
      var result = SalesPersonMasterLoader.normalizeSalesPersonName(pattern.input);

      if (result === pattern.expected) {
        passedCount++;
        Logger.log("  ✓ '" + pattern.input + "' → '" + result + "'");
      } else {
        Logger.log("  ✗ '" + pattern.input + "' → '" + result + "' (期待値: '" + pattern.expected + "')");
      }
    }

    if (passedCount < complexPatterns.length * 0.8) {
      throw new Error("複雑なパターンでの正規化成功率が低すぎます: " + passedCount + "/" + complexPatterns.length);
    }

    // 2. 連続正規化のパフォーマンステスト
    Logger.log("2. 連続正規化パフォーマンステスト:");

    var iterations = 100;
    var testNames = ["高野", "たかの", "枡田", "ますだ", "さいとう"];

    var startTime = new Date();
    for (var i = 0; i < iterations; i++) {
      var name = testNames[i % testNames.length];
      SalesPersonMasterLoader.normalizeSalesPersonName(name);
    }
    var endTime = new Date();
    var totalTime = endTime - startTime;
    var avgTime = totalTime / iterations;

    Logger.log("  " + iterations + "回の正規化: " + totalTime + "ms (平均: " + avgTime.toFixed(2) + "ms/回)");

    if (avgTime > 10) {
      Logger.log("  ⚠ 正規化処理が遅い可能性があります");
    } else {
      Logger.log("  ✓ 正規化処理は十分高速です");
    }

    Logger.log("=== 名前正規化統合テスト完了 ===");
    SalesPersonMasterIntegrationTestResults.passed++;
    return "名前正規化統合テストに成功しました";
  } catch (error) {
    Logger.log("名前正規化統合テストでエラー: " + error);
    SalesPersonMasterIntegrationTestResults.failed++;
    SalesPersonMasterIntegrationTestResults.errors.push("名前正規化統合テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * データ整合性テスト
 */
function testSalesPersonDataConsistency() {
  try {
    Logger.log("=== データ整合性テスト開始 ===");

    // 1. 担当者リストとエイリアスの整合性
    Logger.log("1. 担当者リストとエイリアスの整合性:");

    var persons = SalesPersonMasterLoader.getSalesPersons();
    var aliases = SalesPersonMasterLoader.getSalesPersonAliases();

    // すべての担当者にエイリアスが存在するか確認
    for (var i = 0; i < persons.length; i++) {
      var person = persons[i];

      if (!aliases[person]) {
        throw new Error('担当者 "' + person + '" のエイリアスが存在しません');
      }

      // 本人の名前がエイリアスに含まれているか確認
      if (aliases[person].indexOf(person) === -1) {
        throw new Error('担当者 "' + person + '" の正式名称がエイリアスに含まれていません');
      }

      Logger.log("  ✓ " + person + ": " + aliases[person].length + "個のエイリアス");
    }

    // 2. 重複チェック
    Logger.log("2. 担当者名重複チェック:");

    for (var i = 0; i < persons.length; i++) {
      for (var j = i + 1; j < persons.length; j++) {
        if (persons[i] === persons[j]) {
          throw new Error('重複した担当者名が見つかりました: ' + persons[i]);
        }
      }
    }

    Logger.log("  ✓ 担当者名に重複はありません");

    // 3. エイリアスの逆引きテスト
    Logger.log("3. エイリアス逆引きテスト:");

    var aliasToPersonMap = {};
    var conflictCount = 0;

    for (var person in aliases) {
      if (aliases.hasOwnProperty(person)) {
        var personAliases = aliases[person];
        for (var i = 0; i < personAliases.length; i++) {
          var alias = personAliases[i];

          // 姓のみのエイリアスは複数の人で重複する可能性があるので除外
          var parts = person.split(/[\s　]+/);
          var lastName = parts[0];

          if (alias === lastName && alias !== person) {
            continue; // 姓のみのエイリアスはスキップ
          }

          if (aliasToPersonMap[alias] && aliasToPersonMap[alias] !== person) {
            Logger.log("  ⚠ エイリアス '" + alias + "' が複数の担当者で使用されています: " +
              aliasToPersonMap[alias] + " と " + person);
            conflictCount++;
          } else {
            aliasToPersonMap[alias] = person;
          }
        }
      }
    }

    if (conflictCount > 5) {
      throw new Error('エイリアスの重複が多すぎます: ' + conflictCount + '件');
    }

    Logger.log("  ✓ エイリアス逆引きマップを作成しました（" + Object.keys(aliasToPersonMap).length + "個）");

    Logger.log("=== データ整合性テスト完了 ===");
    SalesPersonMasterIntegrationTestResults.passed++;
    return "データ整合性テストに成功しました";
  } catch (error) {
    Logger.log("データ整合性テストでエラー: " + error);
    SalesPersonMasterIntegrationTestResults.failed++;
    SalesPersonMasterIntegrationTestResults.errors.push("データ整合性テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * パフォーマンステスト
 */
function testSalesPersonPerformance() {
  try {
    Logger.log("=== パフォーマンステスト開始 ===");

    // 1. 複数回呼び出しのパフォーマンステスト
    Logger.log("1. 複数回呼び出しテスト:");

    var iterations = 5;
    var totalTime = 0;
    var results = [];

    for (var i = 0; i < iterations; i++) {
      var startTime = new Date();
      var persons = SalesPersonMasterLoader.getSalesPersons();
      var endTime = new Date();
      var duration = endTime - startTime;

      totalTime += duration;
      results.push(duration);

      Logger.log("  実行" + (i + 1) + ": " + duration + "ms（" + persons.length + "名）");
    }

    var averageTime = totalTime / iterations;
    Logger.log("  平均実行時間: " + averageTime.toFixed(2) + "ms");

    // 2. エイリアス展開パフォーマンステスト
    Logger.log("2. エイリアス展開パフォーマンステスト:");

    var startTime = new Date();
    var aliases = SalesPersonMasterLoader.getSalesPersonAliases();
    var totalAliases = 0;

    for (var person in aliases) {
      if (aliases.hasOwnProperty(person)) {
        totalAliases += aliases[person].length;
      }
    }

    var endTime = new Date();
    var duration = endTime - startTime;

    Logger.log("  総エイリアス数: " + totalAliases);
    Logger.log("  処理時間: " + duration + "ms");

    // 3. プロンプト生成パフォーマンステスト
    Logger.log("3. プロンプト生成パフォーマンステスト:");

    var startTime = new Date();
    var prompt = SalesPersonMasterLoader.getSalesPersonListPrompt();
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
    SalesPersonMasterIntegrationTestResults.passed++;
    return "パフォーマンステストに成功しました";
  } catch (error) {
    Logger.log("パフォーマンステストでエラー: " + error);
    SalesPersonMasterIntegrationTestResults.failed++;
    SalesPersonMasterIntegrationTestResults.errors.push("パフォーマンステスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * エラー回復テスト
 */
function testSalesPersonErrorRecovery() {
  try {
    Logger.log("=== エラー回復テスト開始 ===");

    // 1. スプレッドシートアクセスエラー時のフォールバック
    Logger.log("1. フォールバック機能テスト:");

    // キャッシュをクリアして、確実にloadSalesPersonDataが呼ばれるようにする
    SalesPersonMasterLoader.clearCache();

    // スクリプトプロパティを一時的に無効な値に設定
    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    try {
      // 無効なスプレッドシートIDを設定
      scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', 'invalid_id_12345');

      // この状態でデータ取得を試行（フォールバックが動作するはず）
      var persons = SalesPersonMasterLoader.getSalesPersons();

      if (!Array.isArray(persons) || persons.length === 0) {
        throw new Error('フォールバック時にデフォルトデータが取得できませんでした');
      }

      Logger.log("  ✓ 無効なスプレッドシートID時もデフォルトデータが取得できました（" + persons.length + "名）");

    } finally {
      // 設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }

      // キャッシュをクリアして正常状態に戻す
      SalesPersonMasterLoader.clearCache();
    }

    // 2. 不正な入力への対応
    Logger.log("2. 不正な入力への対応テスト:");

    var invalidInputs = [
      null,
      undefined,
      "",
      "   ",
      123,
      {},
      []
    ];

    var errorCount = 0;
    for (var i = 0; i < invalidInputs.length; i++) {
      try {
        var result = SalesPersonMasterLoader.normalizeSalesPersonName(invalidInputs[i]);
        // エラーが発生しなければOK
      } catch (e) {
        errorCount++;
        Logger.log("  ✗ 入力 " + JSON.stringify(invalidInputs[i]) + " でエラー: " + e);
      }
    }

    if (errorCount > 0) {
      throw new Error('不正な入力でエラーが発生しました: ' + errorCount + '件');
    }

    Logger.log("  ✓ すべての不正な入力が適切に処理されました");

    Logger.log("=== エラー回復テスト完了 ===");
    SalesPersonMasterIntegrationTestResults.passed++;
    return "エラー回復テストに成功しました";
  } catch (error) {
    Logger.log("エラー回復テストでエラー: " + error);
    SalesPersonMasterIntegrationTestResults.failed++;
    SalesPersonMasterIntegrationTestResults.errors.push("エラー回復テスト: " + error.toString());
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 統合テスト結果サマリーの表示
 */
function displaySalesPersonMasterIntegrationTestSummary() {
  Logger.log("\n========== SalesPersonMasterLoader統合テスト サマリー ==========");
  Logger.log("成功したテスト: " + SalesPersonMasterIntegrationTestResults.passed);
  Logger.log("失敗したテスト: " + SalesPersonMasterIntegrationTestResults.failed);

  if (SalesPersonMasterIntegrationTestResults.failed > 0) {
    Logger.log("\n失敗の詳細:");
    for (var i = 0; i < SalesPersonMasterIntegrationTestResults.errors.length; i++) {
      Logger.log("• " + SalesPersonMasterIntegrationTestResults.errors[i]);
    }
  } else {
    Logger.log("🎉 全ての統合テストが成功しました！");
  }

  Logger.log("================================================================");
} 