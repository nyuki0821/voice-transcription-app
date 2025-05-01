/**
 * 環境設定モジュールのテスト用スクリプト
 * 各種テスト関数を実行して動作確認を行うためのファイル
 */

/**
 * 設定ファイルから全ての設定を読み込んでログに出力するテスト
 */
function testConfigLoading() {
  try {
    Logger.log("=== 環境設定読み込みテスト開始 ===");

    // 設定を取得
    var config = EnvironmentConfig.getConfig();

    // 結果を整形してログに出力
    var configKeys = Object.keys(config).sort();
    Logger.log("読み込んだ設定キー数: " + configKeys.length);

    for (var i = 0; i < configKeys.length; i++) {
      var key = configKeys[i];
      var value = config[key];

      // 機密情報はマスク
      if (key.indexOf("KEY") !== -1 || key.indexOf("SECRET") !== -1) {
        value = "********";
      }

      Logger.log(key + " = " + value);
    }

    Logger.log("=== 環境設定読み込みテスト完了 ===");
    return "設定の読み込みに成功しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("環境設定読み込みテストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 特定の設定値を取得するテスト
 */
function testGetSpecificConfig() {
  try {
    Logger.log("=== 特定の環境設定取得テスト開始 ===");

    // テスト対象の設定キー
    var testKeys = [
      "SPREADSHEET_ID",
      "MAX_BATCH_SIZE",
      "ADMIN_EMAILS",
      "ENHANCE_WITH_OPENAI",
      "非存在キー" // 存在しないキーのテスト
    ];

    for (var i = 0; i < testKeys.length; i++) {
      var key = testKeys[i];
      var value = EnvironmentConfig.get(key, "デフォルト値");

      // 機密情報はマスク
      if (key.indexOf("KEY") !== -1 || key.indexOf("SECRET") !== -1) {
        Logger.log(key + " = ********");
      } else {
        Logger.log(key + " = " + value + " (型: " + typeof value + ")");
      }

      // 配列の場合は内容も表示
      if (Array.isArray(value)) {
        Logger.log(" → 配列要素数: " + value.length);
        Logger.log(" → 配列内容: " + JSON.stringify(value));
      }
    }

    Logger.log("=== 特定の環境設定取得テスト完了 ===");
    return "特定の設定値の取得に成功しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("特定の環境設定取得テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 設定キャッシュのテスト
 */
function testConfigCache() {
  try {
    Logger.log("=== 環境設定キャッシュテスト開始 ===");

    // 1回目の取得（キャッシュなし）
    var startTime1 = new Date().getTime();
    var config1 = EnvironmentConfig.getConfig();
    var endTime1 = new Date().getTime();
    var duration1 = endTime1 - startTime1;

    Logger.log("1回目取得時間: " + duration1 + "ms");

    // 2回目の取得（キャッシュあり）
    var startTime2 = new Date().getTime();
    var config2 = EnvironmentConfig.getConfig();
    var endTime2 = new Date().getTime();
    var duration2 = endTime2 - startTime2;

    Logger.log("2回目取得時間: " + duration2 + "ms");

    // 強制リフレッシュ
    var startTime3 = new Date().getTime();
    var config3 = EnvironmentConfig.getConfig(true);
    var endTime3 = new Date().getTime();
    var duration3 = endTime3 - startTime3;

    Logger.log("強制リフレッシュ時間: " + duration3 + "ms");

    // キャッシュクリア
    EnvironmentConfig.clearCache();

    // キャッシュクリア後の取得
    var startTime4 = new Date().getTime();
    var config4 = EnvironmentConfig.getConfig();
    var endTime4 = new Date().getTime();
    var duration4 = endTime4 - startTime4;

    Logger.log("キャッシュクリア後取得時間: " + duration4 + "ms");

    Logger.log("=== 環境設定キャッシュテスト完了 ===");
    return "キャッシュ機能のテストに成功しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("環境設定キャッシュテストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 環境変数の型変換テスト
 */
function testConfigTypeConversion() {
  try {
    Logger.log("=== 環境設定型変換テスト開始 ===");

    var config = EnvironmentConfig.getConfig();

    // 数値型の確認
    var maxBatchSize = config.MAX_BATCH_SIZE;
    Logger.log("MAX_BATCH_SIZE = " + maxBatchSize + " (型: " + typeof maxBatchSize + ")");

    // 真偽値の確認
    var enhanceWithOpenAI = config.ENHANCE_WITH_OPENAI;
    Logger.log("ENHANCE_WITH_OPENAI = " + enhanceWithOpenAI + " (型: " + typeof enhanceWithOpenAI + ")");

    // 配列の確認
    var adminEmails = config.ADMIN_EMAILS;
    Logger.log("ADMIN_EMAILS = " + JSON.stringify(adminEmails) + " (型: " + typeof adminEmails + ", 配列？: " + Array.isArray(adminEmails) + ")");

    Logger.log("=== 環境設定型変換テスト完了 ===");
    return "型変換テストに成功しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("環境設定型変換テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 全てのテストを実行
 */
function runAllConfigTests() {
  var results = [];

  try {
    Logger.log("====== 環境設定モジュール 全テスト開始 ======");

    // 各テストを実行して結果を保存
    results.push({ name: "設定読み込み", result: testConfigLoading() });
    results.push({ name: "特定設定取得", result: testGetSpecificConfig() });
    results.push({ name: "キャッシュ機能", result: testConfigCache() });
    results.push({ name: "型変換", result: testConfigTypeConversion() });

    // 結果をログに出力
    Logger.log("====== テスト結果サマリー ======");
    for (var i = 0; i < results.length; i++) {
      Logger.log((i + 1) + ". " + results[i].name + ": " +
        (results[i].result.indexOf("エラー") === -1 ? "成功" : "失敗"));
    }

    Logger.log("====== 環境設定モジュール 全テスト完了 ======");
    return "全テストが完了しました。詳細はログを確認してください。";
  } catch (error) {
    Logger.log("テスト実行中にエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
} 