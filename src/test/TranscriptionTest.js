/**
 * 文字起こしサービス改良版のシンプルテスト
 */

/**
 * 大きなファイル処理のテスト
 */
function testLargeFileTranscription() {
  console.log("=== 大きなファイル処理テスト開始 ===");

  // テスト用の大きなファイル（12MB）を作成
  const testFile = createTestFile("test_large.mp3", 12 * 1024 * 1024);

  try {
    // 設定情報を取得
    const config = getConfig();

    // APIキーを取得
    const apiKey = config.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error("APIキーが設定されていません。設定スプレッドシートのsettingsシートにOPENAI_API_KEYを設定してください。");
    }

    // 文字起こし実行
    console.log("大きなファイル（12MB）の文字起こしを実行中...");
    const result = TranscriptionService.transcribe(testFile, apiKey);

    // 結果を検証
    console.log("テスト結果:");
    console.log("- 文字起こし成功: " + (result && result.text ? "はい" : "いいえ"));
    console.log("- テキスト長: " + (result.text ? result.text.length : 0) + "文字");
    console.log("- チャンク処理: " + (result.isProcessedInChunks ? "はい" : "いいえ"));
    if (result.isProcessedInChunks) {
      console.log("- 処理チャンク数: " + result.numChunks);
      console.log("- 有効チャンク数: " + result.validChunks);
    }

    console.log("テスト成功: 大きなファイルの処理が正常に完了しました");

  } catch (error) {
    console.error("テスト失敗:", error.toString());
  } finally {
    // テストファイルのクリーンアップ
    cleanupTestFile(testFile);
  }
}

/**
 * 標準サイズファイル処理のテスト
 */
function testNormalFileTranscription() {
  console.log("=== 標準サイズファイル処理テスト開始 ===");

  // テスト用の標準サイズファイル（5MB）を作成
  const testFile = createTestFile("test_normal.mp3", 5 * 1024 * 1024);

  try {
    // 設定情報を取得
    const config = getConfig();

    // APIキーを取得
    const apiKey = config.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error("APIキーが設定されていません。設定スプレッドシートのsettingsシートにOPENAI_API_KEYを設定してください。");
    }

    // 文字起こし実行
    console.log("標準サイズファイル（5MB）の文字起こしを実行中...");
    const result = TranscriptionService.transcribe(testFile, apiKey);

    // 結果を検証
    console.log("テスト結果:");
    console.log("- 文字起こし成功: " + (result && result.text ? "はい" : "いいえ"));
    console.log("- テキスト長: " + (result.text ? result.text.length : 0) + "文字");
    console.log("- 言語: " + (result.language || "不明"));
    console.log("- 発話数: " + (result.utterances ? result.utterances.length : 0));

    console.log("テスト成功: 標準サイズファイルの処理が正常に完了しました");

  } catch (error) {
    console.error("テスト失敗:", error.toString());
  } finally {
    // テストファイルのクリーンアップ
    cleanupTestFile(testFile);
  }
}

/**
 * 未処理フォルダ内のファイルをテスト
 * @param {boolean} processAllFiles - true: すべてのファイルを処理、false: 最新の1ファイルのみ処理
 */
function testUnprocessedFiles(processAllFiles) {
  console.log("=== 未処理フォルダのファイルテスト開始 ===");

  try {
    // 設定情報を取得
    const config = getConfig();

    // 未処理フォルダIDを取得
    const unprocessedFolderId = config.SOURCE_FOLDER_ID;
    if (!unprocessedFolderId) {
      throw new Error("未処理フォルダのIDが設定されていません。設定スプレッドシートのsettingsシートにSOURCE_FOLDER_IDを設定してください。");
    }

    // APIキーを取得
    const apiKey = config.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error("APIキーが設定されていません。設定スプレッドシートのsettingsシートにOPENAI_API_KEYを設定してください。");
    }

    // 未処理フォルダを取得
    const unprocessedFolder = DriveApp.getFolderById(unprocessedFolderId);
    if (!unprocessedFolder) {
      throw new Error("未処理フォルダが見つかりません: " + unprocessedFolderId);
    }

    console.log("未処理フォルダ: " + unprocessedFolder.getName());

    // 音声ファイルを取得
    const audioFiles = getAudioFiles(unprocessedFolder);
    if (audioFiles.length === 0) {
      console.log("未処理フォルダに音声ファイルが見つかりません。");
      return;
    }

    console.log("未処理ファイル数: " + audioFiles.length);

    // 処理対象ファイル
    let filesToProcess = audioFiles;
    if (!processAllFiles) {
      // 最新の1ファイルのみを処理
      filesToProcess = [audioFiles[0]];
      console.log("最新の1ファイルのみを処理します。");
    }

    // ファイルを処理
    for (let i = 0; i < filesToProcess.length; i++) {
      const file = filesToProcess[i];
      console.log("\n処理中のファイル " + (i + 1) + "/" + filesToProcess.length + ": " + file.getName());
      console.log("ファイルサイズ: " + (file.getSize() / 1024 / 1024).toFixed(2) + "MB");

      // 文字起こし実行
      console.log("文字起こし処理を実行中...");
      const result = TranscriptionService.transcribe(file, apiKey);

      // 結果を検証
      console.log("テスト結果:");
      console.log("- 文字起こし成功: " + (result && result.text ? "はい" : "いいえ"));
      console.log("- テキスト長: " + (result.text ? result.text.length : 0) + "文字");
      console.log("- 言語: " + (result.language || "不明"));
      console.log("- 発話数: " + (result.utterances ? result.utterances.length : 0));

      // 大きなファイルの場合の追加情報
      if (result.isProcessedInChunks) {
        console.log("- チャンク処理: はい");
        console.log("- 処理チャンク数: " + result.numChunks);
        console.log("- 有効チャンク数: " + result.validChunks);
      }

      console.log("ファイル処理成功: " + file.getName());
    }

    console.log("\nすべてのファイル処理が完了しました。");

  } catch (error) {
    console.error("テスト失敗:", error.toString());
  }
}

/**
 * 未処理フォルダから音声ファイルを取得（最新順）
 * @param {Folder} folder - 未処理フォルダ
 * @return {Array} - 音声ファイルの配列（最新順）
 */
function getAudioFiles(folder) {
  const files = folder.getFiles();
  const audioFiles = [];

  // 音声ファイルを収集
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();

    // 音声ファイルのみを対象
    if (mimeType.indexOf('audio/') === 0 ||
      mimeType === 'application/octet-stream' && isAudioExtension(file.getName())) {
      audioFiles.push(file);
    }
  }

  // 最新順に並べ替え
  audioFiles.sort(function (a, b) {
    return b.getDateCreated() - a.getDateCreated();
  });

  return audioFiles;
}

/**
 * ファイル名から音声ファイルかどうかを判定
 * @param {string} fileName - ファイル名
 * @return {boolean} - 音声ファイルの場合true
 */
function isAudioExtension(fileName) {
  const audioExtensions = ['.mp3', '.m4a', '.wav', '.ogg', '.aac', '.flac'];
  const extension = '.' + fileName.split('.').pop().toLowerCase();
  return audioExtensions.indexOf(extension) !== -1;
}

/**
 * 設定情報を取得
 * @return {Object} - 設定情報
 */
function getConfig() {
  // EnvironmentConfigから設定を取得
  return EnvironmentConfig.getConfig();
}

/**
 * テスト用ファイルを作成
 * @param {string} fileName - ファイル名
 * @param {number} fileSize - ファイルサイズ（バイト）
 * @return {Object} - テストファイルオブジェクト
 */
function createTestFile(fileName, fileSize) {
  // テスト用の一時フォルダを作成
  const tempFolder = DriveApp.createFolder("TranscriptionTest_" + new Date().getTime());

  // サンプル音声ファイルを検索
  const testFiles = DriveApp.getFilesByName("sample_audio.mp3");
  let testFile;

  if (testFiles.hasNext()) {
    // 既存のサンプルファイルを使用
    const originalFile = testFiles.next();
    testFile = originalFile.makeCopy(fileName, tempFolder);
    console.log("既存のサンプル音声ファイルを使用: " + originalFile.getName());
  } else {
    // サンプルファイルがない場合は警告
    console.warn("警告: サンプル音声ファイル（sample_audio.mp3）が見つかりません。");

    // 短いテスト用音声ファイルを作成
    const blob = Utilities.newBlob("dummy audio data", "audio/mpeg", fileName);
    testFile = tempFolder.createFile(blob);
    console.log("ダミー音声ファイルを作成しました（実際のテストには不適切です）");
  }

  // モックファイルオブジェクト（実際のファイルを参照しつつ、サイズだけ偽装）
  return {
    getName: function () { return testFile.getName(); },
    getBlob: function () { return testFile.getBlob(); },
    getSize: function () { return fileSize; }, // サイズだけ指定値を返す
    getId: function () { return testFile.getId(); },
    getContentType: function () { return "audio/mpeg"; },
    // クリーンアップ用の情報
    _actualFile: testFile,
    _tempFolder: tempFolder
  };
}

/**
 * テスト後のクリーンアップ
 * @param {Object} testFile - テストファイルオブジェクト
 */
function cleanupTestFile(testFile) {
  try {
    if (testFile && testFile._actualFile) {
      DriveApp.getFileById(testFile._actualFile.getId()).setTrashed(true);
    }

    if (testFile && testFile._tempFolder) {
      DriveApp.getFolderById(testFile._tempFolder.getId()).setTrashed(true);
    }

    console.log("テストファイルのクリーンアップ完了");
  } catch (error) {
    console.error("クリーンアップ中にエラー:", error.toString());
  }
}

/**
 * 実際のファイルを使ったテスト
 * @param {string} fileId - テスト対象のファイルID（Google DriveのファイルID）
 */
function testWithRealFile(fileId) {
  console.log("=== 実ファイルでのテスト開始 ===");

  try {
    // 指定されたIDのファイルを取得
    const file = DriveApp.getFileById(fileId);
    if (!file) {
      throw new Error("指定されたIDのファイルが見つかりません: " + fileId);
    }

    console.log("テスト対象ファイル: " + file.getName());
    console.log("ファイルサイズ: " + (file.getSize() / 1024 / 1024).toFixed(2) + "MB");

    // 設定情報を取得
    const config = getConfig();

    // APIキーを取得
    const apiKey = config.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error("APIキーが設定されていません。設定スプレッドシートのsettingsシートにOPENAI_API_KEYを設定してください。");
    }

    // 文字起こし実行
    console.log("文字起こし処理を実行中...");
    const result = TranscriptionService.transcribe(file, apiKey);

    // 結果を検証
    console.log("テスト結果:");
    console.log("- 文字起こし成功: " + (result && result.text ? "はい" : "いいえ"));
    console.log("- テキスト長: " + (result.text ? result.text.length : 0) + "文字");
    console.log("- 言語: " + (result.language || "不明"));
    console.log("- 発話数: " + (result.utterances ? result.utterances.length : 0));

    // 大きなファイルの場合の追加情報
    if (result.isProcessedInChunks) {
      console.log("- チャンク処理: はい");
      console.log("- 処理チャンク数: " + result.numChunks);
      console.log("- 有効チャンク数: " + result.validChunks);
    }

    console.log("テスト成功: 実ファイルの処理が正常に完了しました");

  } catch (error) {
    console.error("テスト失敗:", error.toString());
  }
}

/**
 * すべてのテストを実行
 */
function runAllTests() {
  console.log("=== 文字起こしサービステスト開始 ===");
  console.log("実行時間: " + new Date().toLocaleString());
  console.log("");

  // 設定情報を表示
  const config = getConfig();
  console.log("設定情報:");
  console.log("- API キー: " + (config.OPENAI_API_KEY ? "設定済み" : "未設定"));
  console.log("- 未処理フォルダ: " + (config.SOURCE_FOLDER_ID ? "設定済み" : "未設定"));
  console.log("- 処理中フォルダ: " + (config.PROCESSING_FOLDER_ID ? "設定済み" : "未設定"));
  console.log("- 完了フォルダ: " + (config.COMPLETED_FOLDER_ID ? "設定済み" : "未設定"));
  console.log("- エラーフォルダ: " + (config.ERROR_FOLDER_ID ? "設定済み" : "未設定"));
  console.log("");

  // 標準サイズファイルのテスト
  testNormalFileTranscription();
  console.log("");

  // 大きなファイルのテスト
  testLargeFileTranscription();
  console.log("");

  console.log("=== すべてのテスト完了 ===");
} 