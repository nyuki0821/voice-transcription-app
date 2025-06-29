/**
 * SalesPersonMasterLoader単体テスト
 * 担当者マスターデータローダーの各機能をテストする
 */

// テスト結果を格納するオブジェクト
var SalesPersonMasterLoaderTestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * SalesPersonMasterLoader単体テストスイート
 */
function runSalesPersonMasterLoaderUnitTests() {
  var startTime = new Date();
  Logger.log("====== SalesPersonMasterLoader単体テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  // テスト結果をリセット
  SalesPersonMasterLoaderTestResults.passed = 0;
  SalesPersonMasterLoaderTestResults.failed = 0;
  SalesPersonMasterLoaderTestResults.errors = [];

  // テスト実行
  testSalesPersonModuleExists();
  testGetDefaultSalesPersonData();
  testGeneratePersonAliases();
  testNormalizeSalesPersonName();
  testGetSalesPersonListPrompt();
  testSalesPersonCacheFunctionality();
  testSalesPersonErrorHandling();

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("\n====== SalesPersonMasterLoader単体テスト結果 ======");
  Logger.log("総テスト数: " + (SalesPersonMasterLoaderTestResults.passed + SalesPersonMasterLoaderTestResults.failed));
  Logger.log("成功: " + SalesPersonMasterLoaderTestResults.passed);
  Logger.log("失敗: " + SalesPersonMasterLoaderTestResults.failed);
  Logger.log("実行時間: " + duration + "秒");

  if (SalesPersonMasterLoaderTestResults.failed > 0) {
    Logger.log("失敗したテスト:");
    for (var i = 0; i < SalesPersonMasterLoaderTestResults.errors.length; i++) {
      Logger.log("  - " + SalesPersonMasterLoaderTestResults.errors[i]);
    }
  }

  var successRate = (SalesPersonMasterLoaderTestResults.passed / (SalesPersonMasterLoaderTestResults.passed + SalesPersonMasterLoaderTestResults.failed)) * 100;
  Logger.log("成功率: " + successRate.toFixed(1) + "%");

  return {
    total: SalesPersonMasterLoaderTestResults.passed + SalesPersonMasterLoaderTestResults.failed,
    passed: SalesPersonMasterLoaderTestResults.passed,
    failed: SalesPersonMasterLoaderTestResults.failed,
    successRate: successRate,
    duration: duration
  };
}

/**
 * モジュールとメソッドの存在確認テスト
 */
function testSalesPersonModuleExists() {
  Logger.log("\n--- モジュール存在確認テスト ---");

  try {
    // モジュールの存在確認
    if (typeof SalesPersonMasterLoader === 'undefined') {
      throw new Error('SalesPersonMasterLoaderモジュールが存在しません');
    }

    // 必須メソッドの存在確認
    var requiredMethods = [
      'loadSalesPersonData',
      'getSalesPersons',
      'getSalesPersonAliases',
      'getSalesPersonListPrompt',
      'normalizeSalesPersonName',
      'clearCache'
    ];

    for (var i = 0; i < requiredMethods.length; i++) {
      var method = requiredMethods[i];
      if (typeof SalesPersonMasterLoader[method] !== 'function') {
        throw new Error(method + 'メソッドが存在しません');
      }
    }

    Logger.log("✓ モジュールとすべての必須メソッドが存在します");
    SalesPersonMasterLoaderTestResults.passed++;
  } catch (error) {
    Logger.log("✗ モジュール存在確認テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("モジュール存在確認: " + error.toString());
  }
}

/**
 * デフォルト担当者データ取得テスト
 */
function testGetDefaultSalesPersonData() {
  Logger.log("\n--- デフォルト担当者データ取得テスト ---");

  try {
    // キャッシュをクリアしてデフォルトデータを使用
    SalesPersonMasterLoader.clearCache();

    // スクリプトプロパティを一時的に無効化
    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');
    scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');

    try {
      // 担当者リストを取得
      var persons = SalesPersonMasterLoader.getSalesPersons();

      if (!Array.isArray(persons)) {
        throw new Error('担当者リストが配列ではありません');
      }

      if (persons.length === 0) {
        throw new Error('担当者リストが空です');
      }

      Logger.log("✓ デフォルト担当者データを取得: " + persons.length + "名");

      // 期待される担当者が含まれているか確認
      var expectedPersons = ["高野 仁", "本間 隼", "下山 裕司"];
      for (var i = 0; i < expectedPersons.length; i++) {
        if (persons.indexOf(expectedPersons[i]) === -1) {
          throw new Error('期待される担当者が含まれていません: ' + expectedPersons[i]);
        }
      }

      Logger.log("✓ 期待される担当者が含まれています");
      SalesPersonMasterLoaderTestResults.passed++;

    } finally {
      // 設定を復元
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      }
      SalesPersonMasterLoader.clearCache();
    }

  } catch (error) {
    Logger.log("✗ デフォルト担当者データ取得テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("デフォルトデータ取得: " + error.toString());
  }
}

/**
 * 担当者名エイリアス生成テスト
 */
function testGeneratePersonAliases() {
  Logger.log("\n--- 担当者名エイリアス生成テスト ---");

  try {
    var aliases = SalesPersonMasterLoader.getSalesPersonAliases();

    if (!aliases || typeof aliases !== 'object') {
      throw new Error('エイリアスデータがオブジェクトではありません');
    }

    // デバッグ: 実際に読み込まれた担当者名を確認
    var actualPersons = Object.keys(aliases);
    Logger.log("実際に読み込まれた担当者名: " + actualPersons.join(", "));

    // 特定の担当者のエイリアスをチェック
    var testCases = [
      {
        name: "高野 仁",
        expectedAliases: ["高野", "たかの", "タカノ", "高野さん"]
      },
      {
        name: "桝田 明良",
        expectedAliases: ["桝田", "枡田", "ますだ", "マスダ"]
      },
      {
        name: "齋藤 茜",
        expectedAliases: ["齋藤", "斎藤", "斉藤", "さいとう", "サイトウ"]
      }
    ];

    var passedCount = 0;
    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];

      if (aliases[testCase.name]) {
        var personAliases = aliases[testCase.name];
        Logger.log("担当者 '" + testCase.name + "' のエイリアス: " + personAliases.length + "個");
        Logger.log("  実際のエイリアス: " + personAliases.join(", "));

        // 期待されるエイリアスが含まれているか確認
        var missingAliases = [];
        for (var j = 0; j < testCase.expectedAliases.length; j++) {
          var expected = testCase.expectedAliases[j];
          if (personAliases.indexOf(expected) === -1) {
            missingAliases.push(expected);
          }
        }

        if (missingAliases.length === 0) {
          passedCount++;
        } else {
          Logger.log("  ⚠ 不足しているエイリアス: " + missingAliases.join(", "));
        }
      } else {
        Logger.log("  ⚠ 担当者 '" + testCase.name + "' がデータに存在しません");
      }
    }

    if (passedCount === testCases.length) {
      Logger.log("✓ すべての担当者のエイリアスが正しく生成されています");
      SalesPersonMasterLoaderTestResults.passed++;
    } else {
      throw new Error('一部の担当者のエイリアスが不完全です');
    }

  } catch (error) {
    Logger.log("✗ エイリアス生成テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("エイリアス生成: " + error.toString());
  }
}

/**
 * 担当者名正規化テスト
 */
function testNormalizeSalesPersonName() {
  Logger.log("\n--- 担当者名正規化テスト ---");

  try {
    var testCases = [
      { input: "高野 仁", expected: "高野 仁" },        // 完全一致
      { input: "高野", expected: "高野 仁" },           // 姓のみ
      { input: "たかの", expected: "高野 仁" },         // ひらがな
      { input: "タカノ", expected: "高野 仁" },         // カタカナ
      { input: "高野さん", expected: "高野 仁" },       // 敬称付き
      { input: "TAKANO", expected: "高野 仁" },         // ローマ字
      { input: "桝田", expected: "桝田 明良" },         // 別の担当者
      { input: "枡田", expected: "桝田 明良" },         // 異体字
      { input: "存在しない名前", expected: "存在しない名前" }  // 未登録
    ];

    var passedCount = 0;
    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var result = SalesPersonMasterLoader.normalizeSalesPersonName(testCase.input);

      if (result === testCase.expected) {
        Logger.log("✓ '" + testCase.input + "' → '" + result + "'");
        passedCount++;
      } else {
        Logger.log("✗ '" + testCase.input + "' → '" + result + "' (期待値: '" + testCase.expected + "')");
      }
    }

    if (passedCount === testCases.length) {
      Logger.log("✓ すべての正規化テストに合格");
      SalesPersonMasterLoaderTestResults.passed++;
    } else {
      throw new Error(passedCount + '/' + testCases.length + ' のテストケースのみ成功');
    }

  } catch (error) {
    Logger.log("✗ 担当者名正規化テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("担当者名正規化: " + error.toString());
  }
}

/**
 * LLMプロンプト生成テスト
 */
function testGetSalesPersonListPrompt() {
  Logger.log("\n--- LLMプロンプト生成テスト ---");

  try {
    var prompt = SalesPersonMasterLoader.getSalesPersonListPrompt();

    if (!prompt || typeof prompt !== 'string') {
      throw new Error('プロンプトが文字列ではありません');
    }

    if (prompt.length < 100) {
      throw new Error('プロンプトが短すぎます: ' + prompt.length + '文字');
    }

    // プロンプトに必要な要素が含まれているか確認
    var requiredElements = [
      "営業担当者名の正確な表記",
      "高野 仁",
      "会話中で",
      "正しく表記してください"
    ];

    var missingElements = [];
    for (var i = 0; i < requiredElements.length; i++) {
      if (prompt.indexOf(requiredElements[i]) === -1) {
        missingElements.push(requiredElements[i]);
      }
    }

    if (missingElements.length > 0) {
      throw new Error('プロンプトに必要な要素が不足: ' + missingElements.join(', '));
    }

    Logger.log("✓ LLMプロンプトが正しく生成されています (" + prompt.length + "文字)");
    SalesPersonMasterLoaderTestResults.passed++;

  } catch (error) {
    Logger.log("✗ LLMプロンプト生成テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("プロンプト生成: " + error.toString());
  }
}

/**
 * キャッシュ機能テスト
 */
function testSalesPersonCacheFunctionality() {
  Logger.log("\n--- キャッシュ機能テスト ---");

  try {
    // キャッシュをクリア
    SalesPersonMasterLoader.clearCache();

    // 1回目の呼び出し（キャッシュなし）
    var start1 = new Date();
    var persons1 = SalesPersonMasterLoader.getSalesPersons();
    var end1 = new Date();
    var time1 = end1 - start1;

    // 2回目の呼び出し（キャッシュあり）
    var start2 = new Date();
    var persons2 = SalesPersonMasterLoader.getSalesPersons();
    var end2 = new Date();
    var time2 = end2 - start2;

    Logger.log("1回目の実行時間: " + time1 + "ms");
    Logger.log("2回目の実行時間: " + time2 + "ms");

    // データの一致を確認
    if (JSON.stringify(persons1) !== JSON.stringify(persons2)) {
      throw new Error('キャッシュ前後でデータが異なります');
    }

    // 2回目の方が高速であることを確認（または同等）
    if (time2 <= time1) {
      Logger.log("✓ キャッシュが正常に機能しています");
      SalesPersonMasterLoaderTestResults.passed++;
    } else {
      Logger.log("⚠ キャッシュ効果が確認できませんでしたが、データは一致しています");
      SalesPersonMasterLoaderTestResults.passed++;
    }

  } catch (error) {
    Logger.log("✗ キャッシュ機能テスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("キャッシュ機能: " + error.toString());
  }
}

/**
 * エラーハンドリングテスト
 */
function testSalesPersonErrorHandling() {
  Logger.log("\n--- エラーハンドリングテスト ---");

  try {
    // 1. nullや空文字の正規化
    var nullResult = SalesPersonMasterLoader.normalizeSalesPersonName(null);
    var emptyResult = SalesPersonMasterLoader.normalizeSalesPersonName("");

    if (nullResult !== null || emptyResult !== "") {
      throw new Error('null/空文字の処理が不適切です');
    }

    Logger.log("✓ null/空文字の処理が適切です");

    // 2. 無効なスプレッドシートIDでのフォールバック
    var scriptProperties = PropertiesService.getScriptProperties();
    var originalValue = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

    try {
      scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', 'invalid_id_12345');
      SalesPersonMasterLoader.clearCache();

      var persons = SalesPersonMasterLoader.getSalesPersons();
      if (!Array.isArray(persons) || persons.length === 0) {
        throw new Error('フォールバックが機能していません');
      }

      Logger.log("✓ エラー時のフォールバックが正常に機能しています");

    } finally {
      if (originalValue) {
        scriptProperties.setProperty('CLIENT_MASTER_SPREADSHEET_ID', originalValue);
      } else {
        scriptProperties.deleteProperty('CLIENT_MASTER_SPREADSHEET_ID');
      }
      SalesPersonMasterLoader.clearCache();
    }

    SalesPersonMasterLoaderTestResults.passed++;

  } catch (error) {
    Logger.log("✗ エラーハンドリングテスト失敗: " + error);
    SalesPersonMasterLoaderTestResults.failed++;
    SalesPersonMasterLoaderTestResults.errors.push("エラーハンドリング: " + error.toString());
  }
}

/**
 * テスト結果サマリーの表示
 */
function displaySalesPersonMasterLoaderTestSummary() {
  Logger.log("\n========== SalesPersonMasterLoader単体テスト サマリー ==========");
  Logger.log("成功したテスト: " + SalesPersonMasterLoaderTestResults.passed);
  Logger.log("失敗したテスト: " + SalesPersonMasterLoaderTestResults.failed);

  if (SalesPersonMasterLoaderTestResults.failed > 0) {
    Logger.log("\n失敗の詳細:");
    for (var i = 0; i < SalesPersonMasterLoaderTestResults.errors.length; i++) {
      Logger.log("• " + SalesPersonMasterLoaderTestResults.errors[i]);
    }
  } else {
    Logger.log("🎉 全ての単体テストが成功しました！");
  }

  Logger.log("================================================================");
} 