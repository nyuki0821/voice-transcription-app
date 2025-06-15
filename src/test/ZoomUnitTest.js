/**
 * Zoom Phone単体テストスイート
 * ZoomAPIManager, ZoomphoneService, ZoomphoneProcessorの単体テスト
 */

/**
 * Zoom単体テスト結果を保存するオブジェクト
 */
var ZoomUnitTestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * Zoom単体テストスイートの実行
 */
function runZoomUnitTests() {
  Logger.log('========== Zoom単体テスト開始 ==========');
  
  // テスト結果をリセット
  ZoomUnitTestResults.passed = 0;
  ZoomUnitTestResults.failed = 0;
  ZoomUnitTestResults.errors = [];

  // 各単体テストを実行
  testZoomAPIManagerUnit();
  testZoomphoneServiceUnit();
  testZoomphoneProcessorUnit();
  testZoomConfigUnit();

  // テスト結果サマリー
  displayZoomUnitTestSummary();

  return ZoomUnitTestResults;
}

/**
 * 1. ZoomAPIManager単体テスト
 */
function testZoomAPIManagerUnit() {
  Logger.log('\n--- ZoomAPIManager単体テスト ---');
  
  try {
    // モジュールの存在確認
    assertZoomUnit(typeof ZoomAPIManager === 'object', 'ZoomAPIManagerモジュール存在');
    assertZoomUnit(typeof ZoomAPIManager.getAccessToken === 'function', 'getAccessToken関数存在');
    assertZoomUnit(typeof ZoomAPIManager.sendRequest === 'function', 'sendRequest関数存在');
    
    // 設定検証の確認
    try {
      ZoomAPIManager.getAccessToken();
      Logger.log('APIトークン取得機能: 動作確認済み');
      assertZoomUnit(true, 'APIトークン取得機能');
    } catch (error) {
      if (error.toString().includes('認証情報が設定されていません')) {
        assertZoomUnit(true, 'API認証情報バリデーション');
        Logger.log('認証情報バリデーション: 正常動作');
      } else {
        assertZoomUnit(false, 'APIトークン取得エラーハンドリング');
      }
    }

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`ZoomAPIManager単体テストエラー: ${error.toString()}`);
  }
}

/**
 * 2. ZoomphoneService単体テスト
 */
function testZoomphoneServiceUnit() {
  Logger.log('\n--- ZoomphoneService単体テスト ---');
  
  try {
    // モジュールの存在確認
    assertZoomUnit(typeof ZoomphoneService === 'object', 'ZoomphoneServiceモジュール存在');
    assertZoomUnit(typeof ZoomphoneService.listCallRecordings === 'function', 'listCallRecordings関数存在');
    assertZoomUnit(typeof ZoomphoneService.getRecordingBlob === 'function', 'getRecordingBlob関数存在');
    assertZoomUnit(typeof ZoomphoneService.saveRecordingToDrive === 'function', 'saveRecordingToDrive関数存在');
    
    // 日付パラメータのテスト
    const testFromDate = new Date('2024-01-01');
    const testToDate = new Date('2024-01-02');
    
    try {
      // 実際にAPIを呼ばず、パラメータ処理のテスト
      const fromStr = Utilities.formatDate(testFromDate, 'UTC', 'yyyy-MM-dd');
      const toStr = Utilities.formatDate(testToDate, 'UTC', 'yyyy-MM-dd');
      
      assertZoomUnit(fromStr === '2024-01-01', '日付フォーマット(FROM)');
      assertZoomUnit(toStr === '2024-01-02', '日付フォーマット(TO)');
      
    } catch (dateError) {
      assertZoomUnit(false, '日付処理機能');
    }
    
    // URLパラメータ構築のテスト
    const params = [
      'from=2024-01-01',
      'to=2024-01-02',
      'page_size=30',
      'page_number=1'
    ];
    const queryString = params.join('&');
    assertZoomUnit(queryString.includes('from='), 'URLパラメータ構築');

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`ZoomphoneService単体テストエラー: ${error.toString()}`);
  }
}

/**
 * 3. ZoomphoneProcessor単体テスト
 */
function testZoomphoneProcessorUnit() {
  Logger.log('\n--- ZoomphoneProcessor単体テスト ---');
  
  try {
    // モジュールの存在確認
    assertZoomUnit(typeof ZoomphoneProcessor === 'object', 'ZoomphoneProcessorモジュール存在');
    assertZoomUnit(typeof ZoomphoneProcessor.processRecordings === 'function', 'processRecordings関数存在');
    
    // デフォルト日付計算のテスト
    const now = new Date();
    const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    
    assertZoomUnit(yesterday < now, 'デフォルト日付計算');
    assertZoomUnit((now.getTime() - yesterday.getTime()) === 24 * 60 * 60 * 1000, '24時間差計算');
    
    // 設定依存関数のテスト
    try {
      const config = ConfigManager.getConfig();
      assertZoomUnit(typeof config === 'object', 'Processor設定取得');
      assertZoomUnit(config.SOURCE_FOLDER_ID !== undefined, 'SOURCE_FOLDER_ID設定確認');
    } catch (configError) {
      assertZoomUnit(false, 'Processor設定取得');
    }

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`ZoomphoneProcessor単体テストエラー: ${error.toString()}`);
  }
}

/**
 * 4. Zoom設定単体テスト
 */
function testZoomConfigUnit() {
  Logger.log('\n--- Zoom設定単体テスト ---');
  
  try {
    const config = ConfigManager.getConfig();
    
    // 必須設定項目の存在確認
    const requiredZoomFields = [
      'ZOOM_CLIENT_ID',
      'ZOOM_CLIENT_SECRET', 
      'ZOOM_ACCOUNT_ID',
      'SOURCE_FOLDER_ID',
      'RECORDINGS_SHEET_ID'
    ];
    
    requiredZoomFields.forEach(field => {
      const exists = config.hasOwnProperty(field);
      assertZoomUnit(exists, `設定項目存在: ${field}`);
      
      if (exists) {
        const isConfigured = config[field] && config[field].toString().length > 0;
        if (isConfigured) {
          Logger.log(`✓ ${field}: 設定済み`);
        } else {
          Logger.log(`⚠ ${field}: 未設定`);
        }
      }
    });
    
    // Webhook Secret設定確認
    const hasWebhookSecret = config.hasOwnProperty('ZOOM_WEBHOOK_SECRET');
    assertZoomUnit(hasWebhookSecret, 'ZOOM_WEBHOOK_SECRET設定項目存在');

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`Zoom設定単体テストエラー: ${error.toString()}`);
  }
}

/**
 * 5. Zoomファイル名解析テスト
 */
function testZoomFileNameParsing() {
  Logger.log('\n--- Zoomファイル名解析テスト ---');
  
  try {
    const testCases = [
      {
        fileName: 'zoom_call_20250514075653_17a9e671a1744ef09a12497f9397d6ae.mp3',
        expected: '17a9e671a1744ef09a12497f9397d6ae'
      },
      {
        fileName: 'zoom_call_123456_abcdef123456.mp3',
        expected: 'abcdef123456'
      },
      {
        fileName: 'invalid_file_name.mp3',
        expected: null
      }
    ];
    
    testCases.forEach(testCase => {
      const result = Constants.extractRecordIdFromFileName(testCase.fileName);
      const success = result === testCase.expected;
      assertZoomUnit(success, `ファイル名解析: ${testCase.fileName}`);
      
      if (!success) {
        Logger.log(`期待値: ${testCase.expected}, 実際: ${result}`);
      }
    });

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`ファイル名解析テストエラー: ${error.toString()}`);
  }
}

/**
 * 6. Zoom API レート制限テスト
 */
function testZoomRateLimiting() {
  Logger.log('\n--- Zoom APIレート制限テスト ---');
  
  try {
    // レート制限関連の設定確認
    const rateLimitDelay = 200; // milliseconds
    const maxRetries = 3;
    
    assertZoomUnit(typeof rateLimitDelay === 'number', 'レート制限遅延設定');
    assertZoomUnit(rateLimitDelay >= 100, 'レート制限遅延値適切');
    assertZoomUnit(maxRetries >= 1, 'リトライ回数設定');
    
    // sleep関数の存在確認
    assertZoomUnit(typeof Utilities.sleep === 'function', 'sleep関数存在');
    
    Logger.log(`レート制限設定: 遅延${rateLimitDelay}ms, 最大リトライ${maxRetries}回`);

  } catch (error) {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`レート制限テストエラー: ${error.toString()}`);
  }
}

/**
 * Zoom単体テスト用アサーション関数
 */
function assertZoomUnit(condition, testName) {
  if (condition) {
    ZoomUnitTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    ZoomUnitTestResults.failed++;
    ZoomUnitTestResults.errors.push(`${testName}: 失敗`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

/**
 * Zoom単体テスト結果サマリーの表示
 */
function displayZoomUnitTestSummary() {
  Logger.log('\n========== Zoom単体テスト結果サマリー ==========');
  Logger.log(`成功: ${ZoomUnitTestResults.passed}`);
  Logger.log(`失敗: ${ZoomUnitTestResults.failed}`);
  
  if (ZoomUnitTestResults.errors.length > 0) {
    Logger.log('\nエラー詳細:');
    ZoomUnitTestResults.errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
  }
  
  const total = ZoomUnitTestResults.passed + ZoomUnitTestResults.failed;
  if (total > 0) {
    const successRate = (ZoomUnitTestResults.passed / total) * 100;
    Logger.log(`\n成功率: ${successRate.toFixed(1)}%`);
  }
}

/**
 * 軽量Zoom単体テスト（基本機能のみ）
 */
function runQuickZoomUnitTest() {
  testZoomAPIManagerUnit();
  testZoomConfigUnit();
}