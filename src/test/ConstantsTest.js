/**
 * Constantsモジュールのテスト用スクリプト
 * 定数定義とユーティリティ関数をテストする
 */

/**
 * 定数定義のテスト
 */
function testConstantDefinitions() {
  try {
    Logger.log("=== 定数定義テスト開始 ===");

    // ステータス定数のテスト
    Logger.log("STATUS定数:");
    Logger.log("  SUCCESS: " + Constants.STATUS.SUCCESS);
    Logger.log("  ERROR: " + Constants.STATUS.ERROR);
    Logger.log("  PROCESSING: " + Constants.STATUS.PROCESSING);

    // リトライマーク定数のテスト
    Logger.log("RETRY_MARKS定数:");
    Logger.log("  RETRIED: " + Constants.RETRY_MARKS.RETRIED);
    Logger.log("  FORCE_RETRY: " + Constants.RETRY_MARKS.FORCE_RETRY);

    // エラーパターン定数のテスト
    Logger.log("ERROR_PATTERNS定数 (件数: " + Constants.ERROR_PATTERNS.length + "):");
    for (var i = 0; i < Math.min(3, Constants.ERROR_PATTERNS.length); i++) {
      var pattern = Constants.ERROR_PATTERNS[i];
      Logger.log("  " + (i + 1) + ". " + pattern.pattern + " -> " + pattern.issue);
    }

    // シート名定数のテスト
    Logger.log("SHEET_NAMES定数:");
    Logger.log("  RECORDINGS: " + Constants.SHEET_NAMES.RECORDINGS);
    Logger.log("  CALL_RECORDS: " + Constants.SHEET_NAMES.CALL_RECORDS);

    // カラム名定数のテスト
    Logger.log("COLUMNS定数:");
    Logger.log("  RECORDINGS.ID: " + Constants.COLUMNS.RECORDINGS.ID);
    Logger.log("  CALL_RECORDS.RECORD_ID: " + Constants.COLUMNS.CALL_RECORDS.RECORD_ID);

    Logger.log("=== 定数定義テスト完了 ===");
    return "定数定義のテストに成功しました";
  } catch (error) {
    Logger.log("定数定義テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * メッセージフォーマット機能のテスト
 */
function testFormatMessage() {
  try {
    Logger.log("=== メッセージフォーマット機能テスト開始 ===");

    // 基本的なフォーマットテスト
    var template1 = "こんにちは、{name}さん！今日は{date}です。";
    var values1 = { name: "田中", date: "2024年1月1日" };
    var result1 = Constants.formatMessage(template1, values1);
    Logger.log("テスト1 結果: " + result1);

    // エラーメッセージテンプレートのテスト
    var errorTemplate = Constants.ERROR_MESSAGES.SHEET_NOT_FOUND;
    var errorValues = { sheetName: "test_sheet" };
    var errorResult = Constants.formatMessage(errorTemplate, errorValues);
    Logger.log("エラーメッセージテスト: " + errorResult);

    // 複数の置換テスト
    var template2 = "{user}が{action}を{count}回実行しました。";
    var values2 = { user: "ユーザーA", action: "ログイン", count: "3" };
    var result2 = Constants.formatMessage(template2, values2);
    Logger.log("複数置換テスト: " + result2);

    // 存在しないキーのテスト
    var template3 = "Hello {name}, your {missing} is ready.";
    var values3 = { name: "John" };
    var result3 = Constants.formatMessage(template3, values3);
    Logger.log("存在しないキーテスト: " + result3);

    Logger.log("=== メッセージフォーマット機能テスト完了 ===");
    return "メッセージフォーマット機能のテストに成功しました";
  } catch (error) {
    Logger.log("メッセージフォーマット機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * 音声ファイル判定機能のテスト
 */
function testIsAudioFile() {
  try {
    Logger.log("=== 音声ファイル判定機能テスト開始 ===");

    // モックファイルオブジェクトを作成
    var testCases = [
      {
        name: "test.mp3",
        mimeType: "audio/mpeg",
        expected: true,
        description: "MP3ファイル（MIMEタイプ）"
      },
      {
        name: "recording.wav",
        mimeType: "audio/wav",
        expected: true,
        description: "WAVファイル（MIMEタイプ）"
      },
      {
        name: "voice.m4a",
        mimeType: "application/octet-stream",
        expected: true,
        description: "M4Aファイル（拡張子判定）"
      },
      {
        name: "document.pdf",
        mimeType: "application/pdf",
        expected: false,
        description: "PDFファイル（非音声）"
      },
      {
        name: "image.jpg",
        mimeType: "image/jpeg",
        expected: false,
        description: "JPEGファイル（非音声）"
      }
    ];

    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var mockFile = {
        getName: function () { return testCase.name; },
        getMimeType: function () { return testCase.mimeType; }
      };

      var result = Constants.isAudioFile(mockFile);
      var status = result === testCase.expected ? "成功" : "失敗";
      Logger.log(testCase.description + ": " + result + " (" + status + ")");
    }

    // null/undefinedのテスト
    var nullResult = Constants.isAudioFile(null);
    Logger.log("nullファイルテスト: " + nullResult + " (期待値: false)");

    Logger.log("=== 音声ファイル判定機能テスト完了 ===");
    return "音声ファイル判定機能のテストに成功しました";
  } catch (error) {
    Logger.log("音声ファイル判定機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * レコードID抽出機能のテスト
 */
function testExtractRecordIdFromFileName() {
  try {
    Logger.log("=== レコードID抽出機能テスト開始 ===");

    var testCases = [
      {
        fileName: "zoom_call_123_abc123def456.mp3",
        expected: "abc123def456",
        description: "正常なZoomファイル名"
      },
      {
        fileName: "zoom_call_456_xyz789.mp3",
        expected: "xyz789",
        description: "短いレコードID"
      },
      {
        fileName: "ZOOM_CALL_789_ABC123DEF456GHI789.MP3",
        expected: "ABC123DEF456GHI789",
        description: "大文字のファイル名"
      },
      {
        fileName: "regular_file.mp3",
        expected: null,
        description: "Zoom形式でないファイル名"
      },
      {
        fileName: "zoom_call_invalid.mp3",
        expected: null,
        description: "不正な形式のファイル名"
      },
      {
        fileName: "",
        expected: null,
        description: "空のファイル名"
      }
    ];

    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var result = Constants.extractRecordIdFromFileName(testCase.fileName);
      var status = result === testCase.expected ? "成功" : "失敗";
      Logger.log(testCase.description + ": '" + result + "' (" + status + ")");
    }

    // nullのテスト
    var nullResult = Constants.extractRecordIdFromFileName(null);
    Logger.log("nullファイル名テスト: " + nullResult + " (期待値: null)");

    Logger.log("=== レコードID抽出機能テスト完了 ===");
    return "レコードID抽出機能のテストに成功しました";
  } catch (error) {
    Logger.log("レコードID抽出機能テストでエラー: " + error);
    return "テスト中にエラーが発生しました: " + error;
  }
}

/**
 * Constantsモジュールの全テストを実行
 */
function runAllConstantsTests() {
  var startTime = new Date();
  Logger.log("====== Constants 全テスト開始 ======");
  Logger.log("開始時刻: " + startTime);

  var results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };

  var tests = [
    { name: "定数定義", func: testConstantDefinitions },
    { name: "メッセージフォーマット", func: testFormatMessage },
    { name: "音声ファイル判定", func: testIsAudioFile },
    { name: "レコードID抽出", func: testExtractRecordIdFromFileName }
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

  Logger.log("\n====== Constants テスト結果 ======");
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

  Logger.log("====== Constants 全テスト完了 ======");
  return results;
} 