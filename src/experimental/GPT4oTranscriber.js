/**
 * Whisper Transcribeを使用した文字起こし実験用スクリプト
 * - OpenAI Whisper APIを使用して音声ファイルから文字起こしと話者分離を行う
 */

// 設定値を取得（統一されたConfigManagerを使用）
function getConfig() {
  const config = ConfigManager.getConfig();
  return {
    OPENAI_API_KEY: config.OPENAI_API_KEY,
    TEST_FOLDER_ID: config.TEST_FOLDER_ID || EnvironmentConfig.get('TEST_FOLDER_ID', '')
  };
}

/**
 * テストフォルダ内の音声ファイルを処理するメイン関数
 * - 最新のファイルから処理する
 * - ファイル名に 'processed_' が含まれるものはスキップ
 */
function processTestFiles() {
  const config = getConfig();

  // 設定値の検証
  if (!config.OPENAI_API_KEY) {
    throw new Error('OpenAI APIキーが設定されていません。スクリプトプロパティで設定してください。');
  }
  if (!config.TEST_FOLDER_ID) {
    throw new Error('テストフォルダIDが設定されていません。スクリプトプロパティで設定してください。');
  }

  // テストフォルダを取得
  const folder = DriveApp.getFolderById(config.TEST_FOLDER_ID);
  if (!folder) {
    throw new Error('指定されたテストフォルダIDが見つかりません: ' + config.TEST_FOLDER_ID);
  }

  // デバッグ情報
  Logger.log('テストフォルダ情報: ' + folder.getName() + ' (ID: ' + config.TEST_FOLDER_ID + ')');

  // ファイル一覧を取得（すべてのファイルを取得し、後でフィルタリング）
  const files = folder.getFiles();
  if (!files.hasNext()) {
    Logger.log('処理対象のファイルがありません。テストフォルダに音声ファイルを追加してください。');
    return;
  }

  // 処理対象ファイルを配列に保存
  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // デバッグ情報
    Logger.log('フォルダ内ファイル: ' + fileName + ' (' + mimeType + ')');

    // 音声ファイルのみを処理対象とする
    if (isAudioFile(fileName, mimeType)) {
      fileList.push(file);
      Logger.log('処理対象ファイル検出: ' + fileName + ' (' + mimeType + ')');
    } else {
      Logger.log('音声ファイルではないためスキップ: ' + fileName + ' (' + mimeType + ')');
    }
  }

  // 処理対象ファイルがない場合
  if (fileList.length === 0) {
    Logger.log('処理対象の音声ファイルがありません。テストフォルダに音声ファイルを追加してください。');
    return;
  }

  // 最新のファイルを処理
  const latestFile = fileList[fileList.length - 1];
  Logger.log('処理開始: ' + latestFile.getName());

  try {
    // Whisperベース処理
    Logger.log('Whisperベース処理を開始します...');
    
    // Whisper APIで高精度文字起こし・話者分離
    const whisperResult = TranscriptionService.transcribeWithWhisper(latestFile, config.OPENAI_API_KEY);
    Logger.log('Whisper API 文字起こし・話者分離完了');

    // 結果の洗練
    const enhancedResult = TranscriptionService.enhanceDialogueWithGPT4Mini({text: whisperResult.text}, whisperResult, config.OPENAI_API_KEY);
    Logger.log('GPT-4.1 mini洗練処理完了');

    const finalResult = {
      text: enhancedResult.text,
      rawText: whisperResult.text,
      utterances: whisperResult.utterances || [],
      speakerInfo: whisperResult.speakerInfo || {},
      speakerRoles: whisperResult.speakerRoles || {},
      fileName: latestFile.getName(),
      processingMode: 'whisper_based'
    };

    // 処理結果をログに出力
    Logger.log('======= 処理結果 (Whisperベース) =======');
    Logger.log(finalResult.text);

    return finalResult;

  } catch (error) {
    Logger.log('処理中にエラーが発生しました: ' + error.toString());
    throw error;
  }
}

/**
 * ファイルが音声ファイルかどうかを判定する
 * @param {string} fileName - ファイル名
 * @param {string} mimeType - MIMEタイプ
 * @return {boolean} - 音声ファイルかどうか
 */
function isAudioFile(fileName, mimeType) {
  // 拡張子で判定
  const audioExtensions = ['.mp3', '.wav', '.m4a', '.aac', '.ogg', '.flac', '.wma'];
  const extension = fileName.toLowerCase().substring(fileName.lastIndexOf('.'));
  
  if (audioExtensions.includes(extension)) {
    return true;
  }
  
  // MIMEタイプで判定
  if (mimeType && mimeType.startsWith('audio/')) {
    return true;
  }
  
  return false;
}

/**
 * Whisperベースの文字起こし処理をテストする
 */
function testWhisperBasedProcessing() {
  Logger.log('========== Whisperベース処理テスト開始 ==========');
  
  try {
    const result = processTestFiles();
    
    if (result) {
      Logger.log('テスト成功: ' + result.fileName);
      Logger.log('処理モード: ' + result.processingMode);
      Logger.log('発話数: ' + (result.utterances ? result.utterances.length : 0));
      Logger.log('話者情報: ' + JSON.stringify(result.speakerInfo));
      Logger.log('話者役割: ' + JSON.stringify(result.speakerRoles));
    } else {
      Logger.log('テスト失敗: 結果が取得できませんでした');
    }
    
  } catch (error) {
    Logger.log('テスト中にエラー: ' + error.toString());
  }
  
  Logger.log('========== Whisperベース処理テスト終了 ==========');
}

/**
 * 複数ファイルの一括処理テスト
 */
function processBatch() {
  const config = getConfig();
  
  if (!config.OPENAI_API_KEY || !config.TEST_FOLDER_ID) {
    Logger.log('エラー: 必要な設定が不足しています');
    return;
  }

  const folder = DriveApp.getFolderById(config.TEST_FOLDER_ID);
  const files = folder.getFiles();
  
  const audioFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    if (isAudioFile(file.getName(), file.getMimeType())) {
      audioFiles.push(file);
    }
  }
  
  Logger.log('====== 一括処理開始 ======');
  Logger.log('処理対象ファイル数: ' + audioFiles.length);
  
  const results = [];
  const startTime = new Date();
  
  for (let i = 0; i < audioFiles.length; i++) {
    const file = audioFiles[i];
    Logger.log(`\n[${i + 1}/${audioFiles.length}] 処理中: ${file.getName()}`);
    
    try {
      const whisperResult = TranscriptionService.transcribeWithWhisper(file, config.OPENAI_API_KEY);
      const enhancedResult = TranscriptionService.enhanceDialogueWithGPT4Mini({text: whisperResult.text}, whisperResult, config.OPENAI_API_KEY);
      
      const result = {
        fileName: file.getName(),
        text: enhancedResult.text,
        utteranceCount: whisperResult.utterances ? whisperResult.utterances.length : 0,
        speakerInfo: whisperResult.speakerInfo,
        speakerRoles: whisperResult.speakerRoles,
        entities: whisperResult.entities
      };
      
      results.push(result);
      Logger.log(`完了: ${result.utteranceCount}発話, 文字数: ${result.text.length}`);
      
    } catch (error) {
      Logger.log(`エラー: ${file.getName()} - ${error.toString()}`);
      results.push({
        fileName: file.getName(),
        error: error.toString()
      });
    }
  }
  
  const endTime = new Date();
  const totalTime = (endTime - startTime) / 1000;
  
  Logger.log('\n====== 一括処理結果 ======');
  Logger.log(`総処理時間: ${totalTime}秒`);
  Logger.log(`成功: ${results.filter(r => !r.error).length}件`);
  Logger.log(`失敗: ${results.filter(r => r.error).length}件`);
  
  // 詳細結果
  results.forEach((result, index) => {
    Logger.log(`\n[${index + 1}] ${result.fileName}`);
    if (result.error) {
      Logger.log(`  エラー: ${result.error}`);
    } else {
      Logger.log(`  文字数: ${result.text.length}, 発話数: ${result.utteranceCount}`);
      if (result.speakerInfo && Object.keys(result.speakerInfo).length > 0) {
        Logger.log(`  話者情報: ${JSON.stringify(result.speakerInfo)}`);
      }
    }
  });
  
  return results;
}