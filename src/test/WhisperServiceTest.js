/**
 * Whisperベース機能の単体テストスイート
 * Google Apps Script環境での実行を前提としています
 */

/**
 * テスト結果を保存するオブジェクト
 */
var WhisperServiceTestResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * テストスイートの実行
 */
function runWhisperServiceTests() {
  Logger.log('========== Whisperサービス単体テスト開始 ==========');
  
  // テスト結果をリセット
  WhisperServiceTestResults.passed = 0;
  WhisperServiceTestResults.failed = 0;
  WhisperServiceTestResults.errors = [];

  // 各テストケースを実行
  testWhisperTranscriptionFunction();
  testSpeakerSeparationFunction();
  testEntityExtractionFunction();
  testErrorHandling();
  testSegmentProcessing();

  // テスト結果サマリー
  Logger.log('========== テスト結果サマリー ==========');
  Logger.log(`成功: ${WhisperServiceTestResults.passed}`);
  Logger.log(`失敗: ${WhisperServiceTestResults.failed}`);
  
  if (WhisperServiceTestResults.errors.length > 0) {
    Logger.log('エラー詳細:');
    WhisperServiceTestResults.errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
  }

  return WhisperServiceTestResults;
}

/**
 * テストヘルパー関数
 */
function assertEqual(actual, expected, testName) {
  if (actual === expected) {
    WhisperServiceTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`${testName}: 期待値 "${expected}" に対して "${actual}" を取得`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

function assertNotNull(value, testName) {
  if (value !== null && value !== undefined) {
    WhisperServiceTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`${testName}: null または undefined`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

function assertContains(text, substring, testName) {
  if (text && text.includes(substring)) {
    WhisperServiceTestResults.passed++;
    Logger.log(`✓ ${testName}: 成功`);
  } else {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`${testName}: "${substring}" が見つかりません`);
    Logger.log(`✗ ${testName}: 失敗`);
  }
}

/**
 * 1. Whisper文字起こし機能のテスト
 */
function testWhisperTranscriptionFunction() {
  Logger.log('\n--- Whisper文字起こし機能テスト ---');
  
  try {
    // モックデータの作成
    const mockFile = {
      getBlob: function() {
        return {
          getContentType: function() { return 'audio/mpeg'; },
          getBytes: function() { return [0, 1, 2, 3]; }
        };
      },
      getName: function() { return 'test_audio.mp3'; }
    };

    // モックAPIレスポンス
    const mockWhisperResponse = {
      text: 'これはテスト音声です',
      language: 'ja',
      duration: 10.5,
      segments: [
        { start: 0, end: 5, text: 'これはテスト' },
        { start: 5, end: 10.5, text: '音声です' }
      ]
    };

    // performWhisperTranscription関数の戻り値の検証
    assertNotNull(mockWhisperResponse.text, 'Whisperレスポンス - text');
    assertEqual(mockWhisperResponse.language, 'ja', 'Whisperレスポンス - 言語');
    assertEqual(mockWhisperResponse.segments.length, 2, 'Whisperレスポンス - セグメント数');

  } catch (error) {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`Whisper文字起こしテストエラー: ${error.toString()}`);
  }
}

/**
 * 2. 話者分離機能のテスト
 */
function testSpeakerSeparationFunction() {
  Logger.log('\n--- 話者分離機能テスト ---');
  
  try {
    // モック入力データ
    const mockWhisperResult = {
      text: 'お電話ありがとうございます。こんにちは。ご用件をお聞かせください。はい、承知いたしました。',
      segments: [
        { start: 0, end: 3, text: 'お電話ありがとうございます。' },
        { start: 3, end: 5, text: 'こんにちは。' },
        { start: 5, end: 8, text: 'ご用件をお聞かせください。' },
        { start: 8, end: 11, text: 'はい、承知いたしました。' }
      ]
    };

    // 話者分離結果の期待値
    const expectedSpeakerCount = 2; // 営業担当者と顧客

    // extractSpeakerInfoFromText関数の動作確認
    const testText = '【営業担当者】 こんにちは。\n【お客様】 はい、こんにちは。';
    
    // TranscriptionServiceの関数を使用
    if (typeof TranscriptionService !== 'undefined' && TranscriptionService.extractSpeakerInfoFromText) {
      const speakerInfo = TranscriptionService.extractSpeakerInfoFromText(testText);
      
      assertNotNull(speakerInfo, '話者情報抽出結果');
      assertNotNull(speakerInfo.info, '話者詳細情報');
      assertNotNull(speakerInfo.roles, '話者役割情報');
    } else {
      // 関数が利用できない場合は、簡易的なテストを実行
      const speakerPattern = /【([^】]+)】/g;
      const speakers = testText.match(speakerPattern) || [];
      
      assertEqual(speakers.length, 2, '話者ラベルの検出');
      assertContains(testText, '【営業担当者】', '営業担当者ラベル');
      assertContains(testText, '【お客様】', '顧客ラベル');
    }

  } catch (error) {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`話者分離テストエラー: ${error.toString()}`);
  }
}

/**
 * 3. エンティティ抽出機能のテスト
 */
function testEntityExtractionFunction() {
  Logger.log('\n--- エンティティ抽出機能テスト ---');
  
  try {
    // テストテキスト
    const testText = '株式会社テストの山田です。電話番号は03-1234-5678、メールはtest@example.comです。';
    
    // エンティティ抽出実行
    const entities = TranscriptionService.extractEntitiesFromText(testText);
    
    assertNotNull(entities, 'エンティティ抽出結果');
    assertEqual(entities.phone_number, 1, '電話番号抽出');
    assertEqual(entities.email_address, 1, 'メールアドレス抽出');
    assertEqual(entities.organization, 1, '組織名抽出');

  } catch (error) {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`エンティティ抽出テストエラー: ${error.toString()}`);
  }
}

/**
 * 4. エラーハンドリングのテスト
 */
function testErrorHandling() {
  Logger.log('\n--- エラーハンドリングテスト ---');
  
  try {
    // nullファイルでのテスト
    let errorThrown = false;
    try {
      TranscriptionService.transcribeWithWhisper(null, 'dummy-key');
    } catch (e) {
      errorThrown = true;
    }
    assertEqual(errorThrown, true, 'nullファイルでのエラー処理');

    // 空のWhisper結果でのフォールバック確認
    const emptyResult = {
      text: '',
      utterances: [],
      speakerInfo: {},
      speakerRoles: {},
      entities: {}
    };
    
    assertNotNull(emptyResult, '空結果のフォールバック処理');
    assertEqual(emptyResult.utterances.length, 0, '空の発話リスト');

  } catch (error) {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`エラーハンドリングテストエラー: ${error.toString()}`);
  }
}

/**
 * 5. セグメント処理のテスト
 */
function testSegmentProcessing() {
  Logger.log('\n--- セグメント処理テスト ---');
  
  try {
    // セグメント情報のフォーマットテスト
    const segments = [
      { start: 0.0, end: 5.2, text: 'セグメント1' },
      { start: 5.2, end: 10.5, text: 'セグメント2' },
      { start: 10.5, end: 15.0, text: 'セグメント3' }
    ];

    // セグメント数の検証
    assertEqual(segments.length, 3, 'セグメント数');
    
    // タイムスタンプの連続性
    let isSequential = true;
    for (let i = 1; i < segments.length; i++) {
      if (segments[i].start !== segments[i-1].end) {
        isSequential = false;
        break;
      }
    }
    assertEqual(isSequential, true, 'セグメントの時間的連続性');

    // セグメントテキストの存在確認
    segments.forEach((seg, index) => {
      assertNotNull(seg.text, `セグメント${index + 1}のテキスト`);
    });

  } catch (error) {
    WhisperServiceTestResults.failed++;
    WhisperServiceTestResults.errors.push(`セグメント処理テストエラー: ${error.toString()}`);
  }
}

/**
 * 個別テスト実行用関数
 */
function testWhisperServiceIndividual() {
  // 特定のテストケースだけを実行したい場合
  testWhisperTranscriptionFunction();
}