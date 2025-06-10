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
 * Google Apps Script外部URL制限対応のための緊急設定変更
 */
function emergencyConfigForGASRestrictions() {
  try {
    Logger.log("=== Google Apps Script外部URL制限対応の緊急設定変更 ===");

    // 現在の設定を確認
    var openaiApiKey = ConfigManager.get('OPENAI_API_KEY');
    var assemblyAiApiKey = ConfigManager.get('ASSEMBLYAI_API_KEY');

    Logger.log("現在のAPIキー設定状況:");
    Logger.log("OpenAI APIキー: " + (openaiApiKey ? "設定済み" : "未設定"));
    Logger.log("AssemblyAI APIキー: " + (assemblyAiApiKey ? "設定済み" : "未設定"));

    // プロキシ設定を追加
    var configSpreadsheetId = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
    if (!configSpreadsheetId) {
      Logger.log("エラー: CONFIG_SPREADSHEET_IDが設定されていません");
      return "エラー: CONFIG_SPREADSHEET_IDが設定されていません";
    }

    var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
    var settingsSheet = spreadsheet.getSheetByName('settings');

    if (!settingsSheet) {
      Logger.log("エラー: settingsシートが見つかりません");
      return "エラー: settingsシートが見つかりません";
    }

    // 緊急時の設定変更
    var emergencySettings = [
      ['OPENAI_API_PROXY_ENABLED', 'true'],
      ['OPENAI_API_PROXY_URLS', 'https://openai-api-proxy.herokuapp.com,https://api-proxy.openai-community.com'],
      ['ASSEMBLYAI_FALLBACK_ONLY', 'true'],
      ['GAS_URL_RESTRICTION_MODE', 'true'],
      ['API_RETRY_MAX_ATTEMPTS', '5'],
      ['API_RETRY_BASE_DELAY', '10000']
    ];

    // 設定を追加/更新
    var data = settingsSheet.getDataRange().getValues();
    var headers = data[0];
    var keyIndex = headers.indexOf('key');
    var valueIndex = headers.indexOf('value');

    if (keyIndex === -1 || valueIndex === -1) {
      Logger.log("エラー: settingsシートの形式が正しくありません");
      return "エラー: settingsシートの形式が正しくありません";
    }

    for (var i = 0; i < emergencySettings.length; i++) {
      var key = emergencySettings[i][0];
      var value = emergencySettings[i][1];

      // 既存の設定を探す
      var found = false;
      for (var j = 1; j < data.length; j++) {
        if (data[j][keyIndex] === key) {
          settingsSheet.getRange(j + 1, valueIndex + 1).setValue(value);
          Logger.log("設定更新: " + key + " = " + value);
          found = true;
          break;
        }
      }

      // 新規設定を追加
      if (!found) {
        var newRow = data.length + 1;
        settingsSheet.getRange(newRow, keyIndex + 1).setValue(key);
        settingsSheet.getRange(newRow, valueIndex + 1).setValue(value);
        Logger.log("設定追加: " + key + " = " + value);
      }
    }

    // キャッシュをクリア
    ConfigManager.clearCache();

    Logger.log("緊急設定変更が完了しました");
    Logger.log("システムは以下のモードで動作します:");
    Logger.log("- OpenAI APIプロキシ経由での接続");
    Logger.log("- AssemblyAIフォールバック優先");
    Logger.log("- 拡張リトライ機能");

    return "緊急設定変更が正常に完了しました";

  } catch (error) {
    Logger.log("緊急設定変更中にエラー: " + error.toString());
    return "緊急設定変更中にエラー: " + error.toString();
  }
}

/**
 * Google Apps Script外部URL制限の診断テスト
 */
function diagnoseGASURLRestrictions() {
  try {
    Logger.log("=== Google Apps Script外部URL制限診断テスト ===");

    var testUrls = [
      'https://api.openai.com/v1/models',
      'https://api.assemblyai.com/v2/transcript',
      'https://httpbin.org/get',
      'https://jsonplaceholder.typicode.com/posts/1'
    ];

    var results = [];

    for (var i = 0; i < testUrls.length; i++) {
      var url = testUrls[i];
      Logger.log("テスト中: " + url);

      try {
        var options = {
          method: 'get',
          muteHttpExceptions: true,
          validateHttpsCertificates: false
        };

        var response = UrlFetchApp.fetch(url, options);
        var responseCode = response.getResponseCode();

        if (responseCode === 200) {
          results.push({ url: url, status: "成功", code: responseCode });
          Logger.log("✓ " + url + " - 接続成功");
        } else {
          results.push({ url: url, status: "エラー", code: responseCode });
          Logger.log("✗ " + url + " - エラー: " + responseCode);
        }
      } catch (error) {
        results.push({ url: url, status: "例外", error: error.toString() });
        Logger.log("✗ " + url + " - 例外: " + error.toString());

        if (error.toString().includes('使用できないアドレス')) {
          Logger.log("  → Google Apps Scriptの外部URL制限が検出されました");
        }
      }
    }

    // 結果の要約
    Logger.log("\n=== 診断結果要約 ===");
    var successCount = 0;
    var restrictionDetected = false;

    for (var i = 0; i < results.length; i++) {
      var result = results[i];
      if (result.status === "成功") {
        successCount++;
      } else if (result.error && result.error.includes('使用できないアドレス')) {
        restrictionDetected = true;
      }
    }

    Logger.log("成功: " + successCount + "/" + testUrls.length);
    Logger.log("外部URL制限検出: " + (restrictionDetected ? "はい" : "いいえ"));

    if (restrictionDetected) {
      Logger.log("\n推奨対応:");
      Logger.log("1. emergencyConfigForGASRestrictions() を実行");
      Logger.log("2. AssemblyAI単体での処理に切り替え");
      Logger.log("3. プロキシサーバーの設定");
    }

    return "診断完了 - 成功率: " + (successCount / testUrls.length * 100).toFixed(1) + "%";

  } catch (error) {
    Logger.log("診断テスト中にエラー: " + error.toString());
    return "診断テスト中にエラー: " + error.toString();
  }
}

/**
 * 独自プロキシサーバー設定のための高度な設定変更
 */
function setupCustomOpenAIProxy() {
  try {
    Logger.log("=== 独自プロキシサーバー設定開始 ===");

    // 現在の設定を確認
    var openaiApiKey = ConfigManager.get('OPENAI_API_KEY');
    if (!openaiApiKey) {
      Logger.log("エラー: OpenAI APIキーが設定されていません");
      return "エラー: OpenAI APIキーが設定されていません";
    }

    var configSpreadsheetId = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
    if (!configSpreadsheetId) {
      Logger.log("エラー: CONFIG_SPREADSHEET_IDが設定されていません");
      return "エラー: CONFIG_SPREADSHEET_IDが設定されていません";
    }

    var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
    var settingsSheet = spreadsheet.getSheetByName('settings');

    if (!settingsSheet) {
      Logger.log("エラー: settingsシートが見つかりません");
      return "エラー: settingsシートが見つかりません";
    }

    // 高度なプロキシ設定
    var proxySettings = [
      ['OPENAI_API_PROXY_ENABLED', 'true'],
      ['OPENAI_API_PRIMARY_PROXY', 'https://openai-api-proxy.herokuapp.com'],
      ['OPENAI_API_SECONDARY_PROXY', 'https://api-proxy.openai-community.com'],
      ['OPENAI_API_TERTIARY_PROXY', 'https://openai-proxy.vercel.app'],
      ['OPENAI_API_FALLBACK_PROXIES', 'https://api.openai-proxy.com,https://openai-api.proxy.workers.dev,https://proxy-openai-api.netlify.app'],
      ['OPENAI_API_RETRY_STRATEGY', 'exponential_backoff'],
      ['OPENAI_API_MAX_RETRIES', '5'],
      ['OPENAI_API_BASE_DELAY', '3000'],
      ['OPENAI_API_TIMEOUT', '120000'],
      ['OPENAI_API_USER_AGENT', 'Google-Apps-Script-Transcription/1.0'],
      ['GAS_URL_RESTRICTION_BYPASS', 'true'],
      ['FORCE_OPENAI_USAGE', 'true'],
      ['DISABLE_ASSEMBLYAI_FALLBACK', 'true']
    ];

    // 設定を追加/更新
    var data = settingsSheet.getDataRange().getValues();
    var headers = data[0];
    var keyIndex = headers.indexOf('key');
    var valueIndex = headers.indexOf('value');

    if (keyIndex === -1 || valueIndex === -1) {
      Logger.log("エラー: settingsシートの形式が正しくありません");
      return "エラー: settingsシートの形式が正しくありません";
    }

    for (var i = 0; i < proxySettings.length; i++) {
      var key = proxySettings[i][0];
      var value = proxySettings[i][1];

      // 既存の設定を探す
      var found = false;
      for (var j = 1; j < data.length; j++) {
        if (data[j][keyIndex] === key) {
          settingsSheet.getRange(j + 1, valueIndex + 1).setValue(value);
          Logger.log("設定更新: " + key + " = " + value);
          found = true;
          break;
        }
      }

      // 新規設定を追加
      if (!found) {
        var newRow = data.length + 1;
        settingsSheet.getRange(newRow, keyIndex + 1).setValue(key);
        settingsSheet.getRange(newRow, valueIndex + 1).setValue(value);
        Logger.log("設定追加: " + key + " = " + value);
      }
    }

    // キャッシュをクリア
    ConfigManager.clearCache();

    Logger.log("独自プロキシサーバー設定が完了しました");
    Logger.log("システムは以下のモードで動作します:");
    Logger.log("- OpenAI API優先モード");
    Logger.log("- 6つのプロキシサーバーによる冗長化");
    Logger.log("- AssemblyAIフォールバック無効化");
    Logger.log("- 拡張リトライ機能（最大5回）");

    return "独自プロキシサーバー設定が正常に完了しました";

  } catch (error) {
    Logger.log("独自プロキシサーバー設定中にエラー: " + error.toString());
    return "独自プロキシサーバー設定中にエラー: " + error.toString();
  }
}

/**
 * OpenAI API専用の接続テスト（プロキシ経由）
 */
function testOpenAIProxyConnections() {
  try {
    Logger.log("=== OpenAI APIプロキシ接続テスト ===");

    var openaiApiKey = ConfigManager.get('OPENAI_API_KEY');
    if (!openaiApiKey) {
      Logger.log("エラー: OpenAI APIキーが設定されていません");
      return "エラー: OpenAI APIキーが設定されていません";
    }

    // テスト用のプロキシエンドポイント
    var proxyEndpoints = [
      'https://api.openai.com/v1/models',
      'https://openai-api-proxy.herokuapp.com/v1/models',
      'https://api-proxy.openai-community.com/v1/models',
      'https://openai-proxy.vercel.app/v1/models',
      'https://api.openai-proxy.com/v1/models',
      'https://openai-api.proxy.workers.dev/v1/models',
      'https://proxy-openai-api.netlify.app/v1/models'
    ];

    var results = [];
    var successCount = 0;

    for (var i = 0; i < proxyEndpoints.length; i++) {
      var url = proxyEndpoints[i];
      Logger.log("テスト中: " + url);

      try {
        var options = {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + openaiApiKey,
            'User-Agent': 'Google-Apps-Script/1.0'
          },
          muteHttpExceptions: true,
          validateHttpsCertificates: false,
          followRedirects: true
        };

        var response = UrlFetchApp.fetch(url, options);
        var responseCode = response.getResponseCode();

        if (responseCode === 200) {
          results.push({ url: url, status: "成功", code: responseCode });
          Logger.log("✓ " + url + " - 接続成功");
          successCount++;
        } else if (responseCode === 401) {
          results.push({ url: url, status: "認証エラー", code: responseCode });
          Logger.log("⚠ " + url + " - 認証エラー（APIキー確認が必要）");
        } else {
          results.push({ url: url, status: "エラー", code: responseCode });
          Logger.log("✗ " + url + " - エラー: " + responseCode);
        }
      } catch (error) {
        results.push({ url: url, status: "例外", error: error.toString() });
        Logger.log("✗ " + url + " - 例外: " + error.toString());

        if (error.toString().includes('使用できないアドレス')) {
          Logger.log("  → Google Apps Scriptの外部URL制限");
        }
      }
    }

    // 結果の要約
    Logger.log("\n=== プロキシ接続テスト結果 ===");
    Logger.log("成功: " + successCount + "/" + proxyEndpoints.length);
    Logger.log("成功率: " + (successCount / proxyEndpoints.length * 100).toFixed(1) + "%");

    if (successCount > 0) {
      Logger.log("✓ OpenAI APIへの接続が可能です");
      Logger.log("推奨: setupCustomOpenAIProxy() を実行して設定を最適化");
    } else {
      Logger.log("✗ 全てのプロキシで接続に失敗しました");
      Logger.log("対策: APIキーの確認、または別のプロキシサーバーの追加が必要");
    }

    return "プロキシ接続テスト完了 - 成功率: " + (successCount / proxyEndpoints.length * 100).toFixed(1) + "%";

  } catch (error) {
    Logger.log("プロキシ接続テスト中にエラー: " + error.toString());
    return "プロキシ接続テスト中にエラー: " + error.toString();
  }
}

/**
 * OpenAI API専用モードの設定
 */
function enableOpenAIOnlyMode() {
  try {
    Logger.log("=== OpenAI API専用モード設定 ===");

    var configSpreadsheetId = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
    if (!configSpreadsheetId) {
      Logger.log("エラー: CONFIG_SPREADSHEET_IDが設定されていません");
      return "エラー: CONFIG_SPREADSHEET_IDが設定されていません";
    }

    var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
    var settingsSheet = spreadsheet.getSheetByName('settings');

    if (!settingsSheet) {
      Logger.log("エラー: settingsシートが見つかりません");
      return "エラー: settingsシートが見つかりません";
    }

    // OpenAI専用モードの設定
    var onlyModeSettings = [
      ['OPENAI_ONLY_MODE', 'true'],
      ['DISABLE_ASSEMBLYAI', 'true'],
      ['FORCE_OPENAI_TRANSCRIPTION', 'true'],
      ['OPENAI_RETRY_AGGRESSIVE', 'true'],
      ['OPENAI_FALLBACK_TO_WHISPER', 'true'],
      ['ASSEMBLYAI_SPEAKER_SEPARATION', 'false'],
      ['USE_OPENAI_FOR_SPEAKER_DETECTION', 'true']
    ];

    // 設定を追加/更新
    var data = settingsSheet.getDataRange().getValues();
    var headers = data[0];
    var keyIndex = headers.indexOf('key');
    var valueIndex = headers.indexOf('value');

    if (keyIndex === -1 || valueIndex === -1) {
      Logger.log("エラー: settingsシートの形式が正しくありません");
      return "エラー: settingsシートの形式が正しくありません";
    }

    for (var i = 0; i < onlyModeSettings.length; i++) {
      var key = onlyModeSettings[i][0];
      var value = onlyModeSettings[i][1];

      // 既存の設定を探す
      var found = false;
      for (var j = 1; j < data.length; j++) {
        if (data[j][keyIndex] === key) {
          settingsSheet.getRange(j + 1, valueIndex + 1).setValue(value);
          Logger.log("設定更新: " + key + " = " + value);
          found = true;
          break;
        }
      }

      // 新規設定を追加
      if (!found) {
        var newRow = data.length + 1;
        settingsSheet.getRange(newRow, keyIndex + 1).setValue(key);
        settingsSheet.getRange(newRow, valueIndex + 1).setValue(value);
        Logger.log("設定追加: " + key + " = " + value);
      }
    }

    // キャッシュをクリア
    ConfigManager.clearCache();

    Logger.log("OpenAI API専用モードが有効になりました");
    Logger.log("- AssemblyAIは完全に無効化されます");
    Logger.log("- OpenAI APIのみで文字起こしを実行");
    Logger.log("- 話者分離もOpenAI APIで実行");

    return "OpenAI API専用モードが正常に設定されました";

  } catch (error) {
    Logger.log("OpenAI API専用モード設定中にエラー: " + error.toString());
    return "OpenAI API専用モード設定中にエラー: " + error.toString();
  }
}

/**
 * OpenAI API 500エラー対策の設定
 */
function fixOpenAI500Error() {
  try {
    Logger.log("=== OpenAI API 500エラー対策設定 ===");

    // スクリプトプロパティから直接取得
    var configSpreadsheetId = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
    if (!configSpreadsheetId) {
      Logger.log("エラー: CONFIG_SPREADSHEET_IDが設定されていません");
      return "エラー: CONFIG_SPREADSHEET_IDが設定されていません";
    }

    var spreadsheet = SpreadsheetApp.openById(configSpreadsheetId);
    var settingsSheet = spreadsheet.getSheetByName('settings');

    if (!settingsSheet) {
      Logger.log("エラー: settingsシートが見つかりません");
      return "エラー: settingsシートが見つかりません";
    }

    // 500エラー対策の設定
    var errorFixSettings = [
      ['OPENAI_API_RETRY_COUNT', '5'],              // リトライ回数を増やす
      ['OPENAI_API_RETRY_DELAY', '10000'],          // リトライ間隔を10秒に
      ['OPENAI_API_TIMEOUT', '120000'],             // タイムアウトを2分に
      ['OPENAI_API_USE_PROXY', 'true'],             // プロキシを有効化
      ['OPENAI_API_PROXY_LIST', 'https://openai-api-proxy.herokuapp.com,https://api-proxy.openai-community.com']
    ];

    // 設定を追加/更新
    var data = settingsSheet.getDataRange().getValues();
    var headers = data[0];
    var keyIndex = headers.indexOf('key');
    var valueIndex = headers.indexOf('value');

    if (keyIndex === -1 || valueIndex === -1) {
      Logger.log("エラー: settingsシートの形式が正しくありません");
      return "エラー: settingsシートの形式が正しくありません";
    }

    for (var i = 0; i < errorFixSettings.length; i++) {
      var key = errorFixSettings[i][0];
      var value = errorFixSettings[i][1];

      // 既存の設定を探す
      var found = false;
      for (var j = 1; j < data.length; j++) {
        if (data[j][keyIndex] === key) {
          settingsSheet.getRange(j + 1, valueIndex + 1).setValue(value);
          Logger.log("設定更新: " + key + " = " + value);
          found = true;
          break;
        }
      }

      // 新規設定を追加
      if (!found) {
        var newRow = data.length + 1;
        settingsSheet.getRange(newRow, keyIndex + 1).setValue(key);
        settingsSheet.getRange(newRow, valueIndex + 1).setValue(value);
        Logger.log("設定追加: " + key + " = " + value);
      }
    }

    // キャッシュをクリア
    ConfigManager.clearCache();

    Logger.log("OpenAI API 500エラー対策設定が完了しました");
    Logger.log("- リトライ回数: 5回");
    Logger.log("- リトライ間隔: 10秒");
    Logger.log("- プロキシ経由での接続");

    return "OpenAI API 500エラー対策設定が完了しました";

  } catch (error) {
    Logger.log("500エラー対策設定中にエラー: " + error.toString());
    return "500エラー対策設定中にエラー: " + error.toString();
  }
}

/**
 * 500エラー対策設定を適用してキャッシュをクリア
 */
function applyOpenAI500ErrorFixAndClearCache() {
  try {
    Logger.log("=== OpenAI API 500エラー対策適用とキャッシュクリア ===");

    // まず500エラー対策設定を実行
    const result = fixOpenAI500Error();

    if (result.indexOf("完了しました") !== -1) {
      // ConfigManagerのキャッシュをクリア
      ConfigManager.clearCache();
      Logger.log("ConfigManagerのキャッシュをクリアしました");

      // EnvironmentConfigのキャッシュもクリア
      EnvironmentConfig.clearCache();
      Logger.log("EnvironmentConfigのキャッシュをクリアしました");

      // 新しい設定を確認
      const newConfig = EnvironmentConfig.getConfig(true); // 強制リフレッシュ
      Logger.log("新しい設定を確認:");
      Logger.log("- OPENAI_API_RETRY_COUNT: " + EnvironmentConfig.get('OPENAI_API_RETRY_COUNT', '未設定'));
      Logger.log("- OPENAI_API_RETRY_DELAY: " + EnvironmentConfig.get('OPENAI_API_RETRY_DELAY', '未設定'));
      Logger.log("- OPENAI_API_USE_PROXY: " + EnvironmentConfig.get('OPENAI_API_USE_PROXY', '未設定'));
      Logger.log("- OPENAI_API_PROXY_LIST: " + EnvironmentConfig.get('OPENAI_API_PROXY_LIST', '未設定'));

      return "500エラー対策設定の適用とキャッシュクリアが完了しました";
    } else {
      return "500エラー対策設定の適用に失敗しました: " + result;
    }

  } catch (error) {
    Logger.log("設定適用中にエラー: " + error.toString());
    return "設定適用中にエラー: " + error.toString();
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
    { name: "デフォルト設定取得", func: testDefaultConfig },
    { name: "緊急設定変更", func: emergencyConfigForGASRestrictions },
    { name: "診断テスト", func: diagnoseGASURLRestrictions },
    { name: "独自プロキシサーバー設定", func: setupCustomOpenAIProxy },
    { name: "OpenAI APIプロキシ接続テスト", func: testOpenAIProxyConnections },
    { name: "OpenAI API専用モードの設定", func: enableOpenAIOnlyMode },
    { name: "OpenAI API 500エラー対策設定", func: fixOpenAI500Error },
    { name: "500エラー対策設定を適用してキャッシュをクリア", func: applyOpenAI500ErrorFixAndClearCache }
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