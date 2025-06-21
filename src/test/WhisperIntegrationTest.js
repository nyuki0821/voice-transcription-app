/**
 * Whisperベース文字起こしシステムの統合テストスイート
 * ファイル入力から最終出力までの全体フローを検証
 */

/**
 * 統合テスト結果を保存するオブジェクト
 */
var WhisperIntegrationTestResults = {
  passed: 0,
  failed: 0,
  errors: [],
  performanceMetrics: {}
};

/**
 * 統合テストスイートの実行
 */
function runWhisperIntegrationTests() {
  Logger.log('========== Whisper統合テスト開始 ==========');
  
  // テスト結果をリセット
  WhisperIntegrationTestResults.passed = 0;
  WhisperIntegrationTestResults.failed = 0;
  WhisperIntegrationTestResults.errors = [];
  WhisperIntegrationTestResults.performanceMetrics = {};

  // 設定の確認
  const config = getTestConfig();
  if (!config.OPENAI_API_KEY || !config.TEST_FOLDER_ID) {
    Logger.log('エラー: 必要な設定が不足しています');
    return;
  }

  // 各統合テストを実行
  testFullTranscriptionFlow(config);
  testErrorHandling(config);
  testPerformanceMetrics(config);
  testOutputQuality(config);
  testErrorRecovery(config);

  // テスト結果サマリー
  displayIntegrationTestSummary();

  return WhisperIntegrationTestResults;
}

/**
 * テスト用設定の取得
 */
function getTestConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    OPENAI_API_KEY: scriptProperties.getProperty('OPENAI_API_KEY'),
    TEST_FOLDER_ID: scriptProperties.getProperty('TEST_FOLDER_ID')
  };
}

/**
 * テスト用ファイルの取得（サイズ制限を考慮）
 */
function getTestFile(folder, maxSizeMB = 1) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileSizeMB = file.getSize() / 1024 / 1024;
    if (fileSizeMB <= maxSizeMB) {
      Logger.log(`適切なサイズのテストファイルを発見: ${file.getName()} (${fileSizeMB.toFixed(2)}MB)`);
      return file;
    }
  }
  return null;
}

/**
 * 1. 完全な文字起こしフローのテスト
 */
function testFullTranscriptionFlow(config) {
  Logger.log('\n--- 完全文字起こしフローテスト ---');
  
  try {
    // テストフォルダから適切なサイズの音声ファイルを取得
    const folder = DriveApp.getFolderById(config.TEST_FOLDER_ID);
    const testFile = getTestFile(folder, 1); // 1MB以下のファイルを探す
    
    if (!testFile) {
      Logger.log('1MB以下のテストファイルが見つからないため、モックファイルでテスト');
      const mockFile = createMockAudioFile('test_flow.mp3', 512);
      
      try {
        const result = TranscriptionService.transcribe(mockFile, config.OPENAI_API_KEY);
        // モックファイルではAPI呼び出しが失敗することが期待される
        WhisperIntegrationTestResults.failed++;
        WhisperIntegrationTestResults.errors.push('完全フローテスト: モックファイルでのAPI呼び出しが予期せず成功');
      } catch (mockError) {
        Logger.log('モックファイルでの予期されたエラー: ' + mockError.toString());
        // モックファイルでのエラーは予期された動作
        const isExpectedError = mockError.toString().includes('file could not be decoded') ||
                               mockError.toString().includes('Invalid file format');
        assertEqual(isExpectedError, true, '完全フロー実行（モックファイル・予期されたエラー）');
      }
      return;
    }

    const startTime = new Date().getTime();
    
    // メインのtranscribe関数を実行
    const result = TranscriptionService.transcribe(
      testFile,
      config.OPENAI_API_KEY
    );
    
    const endTime = new Date().getTime();
    const processingTime = (endTime - startTime) / 1000;

    // 結果の検証
    assertIntegrationResult(result, '完全フロー実行');
    assertNotNull(result.text, '文字起こしテキスト');
    assertNotNull(result.processingMode, '処理モード');
    
    // Whisperベースで処理されたか確認
    assertEqual(result.processingMode, 'whisper_based', 'Whisperベース処理の確認');
    
    // パフォーマンスメトリクスの記録
    WhisperIntegrationTestResults.performanceMetrics.fullFlowTime = processingTime;
    Logger.log(`処理時間: ${processingTime}秒`);

  } catch (error) {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`完全フローテストエラー: ${error.toString()}`);
  }
}

/**
 * 2. エラーハンドリングのテスト
 */
function testErrorHandling(config) {
  Logger.log('\n--- エラーハンドリングテスト ---');
  
  try {
    // 無効なAPIキーでエラーテスト
    const mockFile = createMockAudioFile('error_test.mp3');
    const invalidApiKey = 'invalid-api-key';
    
    let errorHandled = false;
    try {
      const result = TranscriptionService.transcribe(
        mockFile,
        invalidApiKey
      );
    } catch (error) {
      errorHandled = true;
      Logger.log(`想定内のエラー: ${error.toString()}`);
    }

    assertEqual(errorHandled, true, 'エラーハンドリング機能の動作');

  } catch (error) {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`エラーハンドリングテストエラー: ${error.toString()}`);
  }
}

/**
 * 3. パフォーマンス測定テスト
 */
function testPerformanceMetrics(config) {
  Logger.log('\n--- パフォーマンス測定テスト ---');
  
  try {
    const folder = DriveApp.getFolderById(config.TEST_FOLDER_ID);
    const files = folder.getFiles();
    
    if (!files.hasNext()) {
      Logger.log('パフォーマンステスト: ファイルなし');
      return;
    }

    const testFile = files.next();
    const fileSize = testFile.getSize() / 1024 / 1024; // MB単位
    
    // 各処理ステップの時間を測定
    const metrics = {
      whisperTime: 0,
      separationTime: 0,
      refinementTime: 0,
      totalTime: 0
    };

    const totalStart = new Date().getTime();
    
    // Whisper処理時間の測定
    const whisperStart = new Date().getTime();
    const whisperResult = TranscriptionService.performWhisperTranscription(
      testFile,
      config.OPENAI_API_KEY
    );
    metrics.whisperTime = (new Date().getTime() - whisperStart) / 1000;

    // 話者分離時間の測定
    const separationStart = new Date().getTime();
    const separationResult = TranscriptionService.performSpeakerSeparationWithWhisper(
      whisperResult,
      config.OPENAI_API_KEY
    );
    metrics.separationTime = (new Date().getTime() - separationStart) / 1000;

    metrics.totalTime = (new Date().getTime() - totalStart) / 1000;

    // メトリクスの記録
    WhisperIntegrationTestResults.performanceMetrics = {
      ...WhisperIntegrationTestResults.performanceMetrics,
      ...metrics,
      fileSize: fileSize,
      processingSpeed: fileSize / metrics.totalTime // MB/秒
    };

    Logger.log(`ファイルサイズ: ${fileSize.toFixed(2)} MB`);
    Logger.log(`Whisper処理: ${metrics.whisperTime.toFixed(2)}秒`);
    Logger.log(`話者分離: ${metrics.separationTime.toFixed(2)}秒`);
    Logger.log(`合計時間: ${metrics.totalTime.toFixed(2)}秒`);
    Logger.log(`処理速度: ${(fileSize / metrics.totalTime).toFixed(2)} MB/秒`);

    // パフォーマンス基準の検証（Whisper APIの現実的な処理速度: 0.05 MB/秒以上）
    const processingSpeed = fileSize / metrics.totalTime;
    const isPerformanceAcceptable = processingSpeed > 0.05;
    assertEqual(isPerformanceAcceptable, true, 'パフォーマンス基準');

  } catch (error) {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`パフォーマンステストエラー: ${error.toString()}`);
  }
}

/**
 * 4. 出力品質のテスト
 */
function testOutputQuality(config) {
  Logger.log('\n--- 出力品質テスト ---');
  
  try {
    const folder = DriveApp.getFolderById(config.TEST_FOLDER_ID);
    const files = folder.getFiles();
    
    if (!files.hasNext()) {
      Logger.log('品質テスト: ファイルなし');
      return;
    }

    const testFile = files.next();
    const result = TranscriptionService.transcribe(
      testFile,
      config.OPENAI_API_KEY
    );

    // 品質チェック項目
    const qualityChecks = {
      hasText: result.text && result.text.length > 0,
      hasSpeakerLabels: result.text.includes('【') && result.text.includes('】'),
      hasMultipleSpeakers: (result.text.match(/【[^】]+】/g) || []).length > 1,
      hasProperFormatting: result.text.includes('\n\n'),
      hasEntityInfo: result.text.includes('エンティティ情報') || 
                     Object.keys(result.speakerInfo || {}).length > 0
    };

    // 各品質項目の検証
    assertEqual(qualityChecks.hasText, true, '文字起こしテキストの存在');
    assertEqual(qualityChecks.hasSpeakerLabels, true, '話者ラベルの存在');
    assertEqual(qualityChecks.hasMultipleSpeakers, true, '複数話者の識別');
    assertEqual(qualityChecks.hasProperFormatting, true, '適切なフォーマット');
    
    // 品質スコアの計算（0-100）
    const qualityScore = Object.values(qualityChecks).filter(v => v).length * 20;
    Logger.log(`品質スコア: ${qualityScore}/100`);
    
    WhisperIntegrationTestResults.performanceMetrics.qualityScore = qualityScore;

  } catch (error) {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`品質テストエラー: ${error.toString()}`);
  }
}

/**
 * 5. エラーリカバリーのテスト
 */
function testErrorRecovery(config) {
  Logger.log('\n--- エラーリカバリーテスト ---');
  
  try {
    // 様々なエラーシナリオをテスト
    const errorScenarios = [
      {
        name: '空ファイル処理',
        test: () => {
          const emptyFile = createMockAudioFile('empty.mp3', 0);
          try {
            TranscriptionService.transcribe(emptyFile, config.OPENAI_API_KEY);
            return false; // モックファイルが成功すべきでない
          } catch (e) {
            // 期待されるエラー（ファイル形式の問題、空ファイル等）
            const isExpectedError = e.toString().includes('file could not be decoded') ||
                                   e.toString().includes('Invalid file format') ||
                                   e.toString().includes('ファイルが指定されていません');
            Logger.log('期待されるエラー: ' + e.toString());
            return isExpectedError;
          }
        }
      },
      {
        name: '大容量ファイル処理',
        test: () => {
          // 25MB超過のテスト（Whisper APIの上限テスト）
          const largeFile = createMockAudioFile('large.mp3', 512);
          largeFile.getSize = function() { return 26 * 1024 * 1024; }; // 26MBとして報告（制限超過）
          
          try {
            TranscriptionService.transcribe(largeFile, config.OPENAI_API_KEY);
            return false; // 制限超過ファイルが成功すべきでない
          } catch (e) {
            // サイズ制限エラーまたはファイル形式エラーが期待される
            const isExpectedError = e.toString().includes('ファイルサイズが25MB制限を超えています') ||
                                   e.toString().includes('ファイルサイズが大きすぎます') ||
                                   e.toString().includes('file could not be decoded');
            Logger.log('想定内のエラー: ' + e.toString());
            return isExpectedError;
          }
        }
      }
    ];

    // 各シナリオを実行
    errorScenarios.forEach(scenario => {
      const recovered = scenario.test();
      assertEqual(recovered, true, `エラーリカバリー: ${scenario.name}`);
    });

  } catch (error) {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`エラーリカバリーテストエラー: ${error.toString()}`);
  }
}

/**
 * モック音声ファイルの作成（テスト用の小さいサイズ）
 */
function createMockAudioFile(fileName, size = 256) {
  return {
    getBlob: function() {
      return {
        getContentType: function() { return 'audio/mpeg'; },
        getBytes: function() { 
          // MP3のヘッダーを模擬した基本的なバイト配列
          const bytes = [];
          // MP3ファイルの基本ヘッダー（ID3v2.3）
          if (size > 0) {
            bytes[0] = 0x49; // 'I'
            bytes[1] = 0x44; // 'D'
            bytes[2] = 0x33; // '3'
            bytes[3] = 0x03; // version
            bytes[4] = 0x00; // revision
            bytes[5] = 0x00; // flags
          }
          // 残りのバイトを埋める
          for (let i = 6; i < size && i < 512; i++) {
            bytes[i] = (i * 17) % 256; // よりランダムなパターン
          }
          return bytes;
        }
      };
    },
    getName: function() { return fileName; },
    getSize: function() { return Math.min(size, 512); } // 最大512バイト
  };
}

/**
 * 統合テスト用アサーション関数
 */
function assertIntegrationResult(result, testName) {
  if (result && typeof result === 'object') {
    WhisperIntegrationTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`${testName}: 結果オブジェクトが無効`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

function assertEqual(actual, expected, testName) {
  if (actual === expected) {
    WhisperIntegrationTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(
      `${testName}: 期待値 "${expected}" に対して "${actual}" を取得`
    );
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

function assertNotNull(value, testName) {
  if (value !== null && value !== undefined) {
    WhisperIntegrationTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperIntegrationTestResults.failed++;
    WhisperIntegrationTestResults.errors.push(`${testName}: null または undefined`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

/**
 * テスト結果サマリーの表示
 */
function displayIntegrationTestSummary() {
  Logger.log('\n========== 統合テスト結果サマリー ==========');
  Logger.log(`成功: ${WhisperIntegrationTestResults.passed}`);
  Logger.log(`失敗: ${WhisperIntegrationTestResults.failed}`);
  
  if (WhisperIntegrationTestResults.errors.length > 0) {
    Logger.log('\nエラー詳細:');
    WhisperIntegrationTestResults.errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
  }

  if (Object.keys(WhisperIntegrationTestResults.performanceMetrics).length > 0) {
    Logger.log('\nパフォーマンスメトリクス:');
    Object.entries(WhisperIntegrationTestResults.performanceMetrics).forEach(([key, value]) => {
      Logger.log(`  ${key}: ${typeof value === 'number' ? value.toFixed(2) : value}`);
    });
  }
}

/**
 * 簡易実行用関数
 */
function runQuickIntegrationTest() {
  const config = getTestConfig();
  testFullTranscriptionFlow(config);
}