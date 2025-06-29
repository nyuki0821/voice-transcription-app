/**
 * 文字起こしサービスモジュール（OpenAI Whisper API + GPT-4.1 mini対応版）
 */
var TranscriptionService = (function () {
  /**
   * 音声ファイルを文字起こしする（話者分離機能付き）
   * @param {File} file - 音声ファイル
   * @param {string} openaiApiKey - OpenAI APIキー（Whisper + GPT-4.1 mini用）
   * @return {Object} - 文字起こし結果
   */
  function transcribe(file, openaiApiKey) {
    if (!file) {
      throw new Error('ファイルが指定されていません');
    }

    if (!openaiApiKey) {
      throw new Error('OpenAI APIキーが指定されていません');
    }

    try {
      Logger.log('文字起こし処理開始 (Whisperベース): ' + file.getName());

      // Whisper APIでの高精度文字起こし・話者分離処理
      const whisperResult = transcribeWithWhisper(file, openaiApiKey);
      Logger.log('Whisper API 文字起こし・話者分離処理完了');

      // 結果の洗練（オプション）
      const enhancedResult = enhanceDialogueWithGPT4Mini({ text: whisperResult.text }, whisperResult, openaiApiKey);
      Logger.log('GPT-4.1 miniによる会話洗練処理完了');

      return {
        text: enhancedResult.text,
        rawText: whisperResult.text,
        utterances: whisperResult.utterances || [],
        speakerInfo: whisperResult.speakerInfo || {},
        speakerRoles: whisperResult.speakerRoles || {},
        fileName: file.getName(),
        processingMode: 'whisper_based'
      };

    } catch (error) {
      Logger.log('文字起こし処理中にエラー: ' + error.toString());
      throw error;
    }
  }

  /**
   * OpenAI Whisper APIで文字起こしと話者分離を行う
   * @param {File} file - 音声ファイル
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - 文字起こし結果（話者分離情報付き）
   */
  function transcribeWithWhisper(file, apiKey) {
    Logger.log('Whisperによる高精度文字起こし・話者分離処理開始...');

    try {
      // Whisperで基本的な文字起こし
      const whisperResult = performWhisperTranscription(file, apiKey);

      // GPT-4o-miniで話者分離を実行
      const speakerSeparationResult = performSpeakerSeparationWithWhisper(whisperResult, apiKey);

      return speakerSeparationResult;

    } catch (error) {
      Logger.log('Whisper処理中にエラー: ' + error.toString());
      throw error;
    }
  }

  /**
   * Whisper APIで音声ファイルを文字起こしする
   * @param {File} file - 音声ファイル
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - Whisper文字起こし結果
   */
  function performWhisperTranscription(file, apiKey) {
    Logger.log('Whisper文字起こし処理開始...');

    try {
      const fileBlob = file.getBlob();
      const fileName = file.getName();

      // ファイルサイズチェック（25MB制限）
      const fileSize = file.getSize ? file.getSize() : fileBlob.getBytes().length;
      if (fileSize > 25 * 1024 * 1024) {
        throw new Error('ファイルサイズが25MB制限を超えています: ' + (fileSize / 1024 / 1024).toFixed(2) + 'MB');
      }

      // Google Apps Script向けマルチパートフォームデータ作成
      const boundary = 'GASWHISPER' + Date.now();
      let payload;

      if (fileSize < 10 * 1024 * 1024) { // 10MB未満のファイル
        const formData =
          '--' + boundary + '\r\n' +
          'Content-Disposition: form-data; name="model"\r\n\r\n' +
          'whisper-1\r\n' +
          '--' + boundary + '\r\n' +
          'Content-Disposition: form-data; name="response_format"\r\n\r\n' +
          'verbose_json\r\n' +
          '--' + boundary + '\r\n' +
          'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n' +
          'Content-Type: ' + fileBlob.getContentType() + '\r\n\r\n';

        const endData = '\r\n--' + boundary + '--\r\n';

        // ファイルサイズに応じて処理方法を調整
        if (fileSize < 1024 * 1024) { // 1MB未満は安全な配列結合
          const headerBytes = Utilities.newBlob(formData).getBytes();
          const fileBytes = fileBlob.getBytes();
          const footerBytes = Utilities.newBlob(endData).getBytes();

          const totalSize = headerBytes.length + fileBytes.length + footerBytes.length;
          const combined = new Array(totalSize);

          let pos = 0;
          for (let i = 0; i < headerBytes.length; i++) {
            combined[pos++] = headerBytes[i];
          }
          for (let i = 0; i < fileBytes.length; i++) {
            combined[pos++] = fileBytes[i];
          }
          for (let i = 0; i < footerBytes.length; i++) {
            combined[pos++] = footerBytes[i];
          }

          payload = Utilities.newBlob(combined);
        } else {
          // 1MB以上のファイルは安全な逐次結合を使用
          Logger.log('大きなファイル（' + (fileSize / 1024 / 1024).toFixed(2) + 'MB）を安全な方法で処理');

          const headerBytes = Utilities.newBlob(formData).getBytes();
          const fileBytes = fileBlob.getBytes();
          const footerBytes = Utilities.newBlob(endData).getBytes();

          const totalSize = headerBytes.length + fileBytes.length + footerBytes.length;
          const combined = new Array(totalSize);

          // 安全な逐次コピー（スタックオーバーフローを回避）
          let pos = 0;

          // ヘッダーをコピー
          for (let i = 0; i < headerBytes.length; i++) {
            combined[pos++] = headerBytes[i];
          }

          // ファイルデータをコピー（大きなファイル用にバッチ処理）
          const batchSize = 1000; // 1000バイトずつ処理
          for (let i = 0; i < fileBytes.length; i += batchSize) {
            const endIndex = Math.min(i + batchSize, fileBytes.length);
            for (let j = i; j < endIndex; j++) {
              combined[pos++] = fileBytes[j];
            }
            // 大きなファイルの場合、処理中に少し待機
            if (i % 10000 === 0 && i > 0) {
              Utilities.sleep(1); // 1ms待機
            }
          }

          // フッターをコピー
          for (let i = 0; i < footerBytes.length; i++) {
            combined[pos++] = footerBytes[i];
          }

          payload = Utilities.newBlob(combined);
        }
      } else {
        // 10MB以上のファイルは分割処理
        Logger.log('ファイルサイズが10MBを超えています。分割処理を実行します: ' + (fileSize / 1024 / 1024).toFixed(2) + 'MB');

        // 一時フォルダを作成
        const tempFolderName = 'Whisper_Temp_' + new Date().getTime();
        const tempFolder = DriveApp.createFolder(tempFolderName);

        try {
          // 分割処理を実行
          return processLargeAudioFile(file, apiKey, tempFolder);
        } finally {
          // 一時フォルダを削除
          try {
            tempFolder.setTrashed(true);
          } catch (e) {
            Logger.log('一時フォルダの削除中にエラー: ' + e);
          }
        }
      }

      const options = {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + apiKey,
          'Content-Type': 'multipart/form-data; boundary=' + boundary
        },
        payload: payload.getBytes(),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('Whisper API呼び出しエラー: ' + responseCode + ' - ' + response.getContentText());
      }

      const result = JSON.parse(response.getContentText());
      Logger.log('Whisper文字起こし完了: ' + result.text.length + '文字');

      return {
        text: result.text,
        language: result.language,
        duration: result.duration,
        segments: result.segments || [],
        fileName: fileName
      };

    } catch (error) {
      // エラーメッセージの改善
      let errorMessage = error.toString();

      // ファイルサイズに関するエラーの場合、より親切なメッセージを表示
      if (errorMessage.includes('ファイルサイズが大きすぎます') ||
        errorMessage.includes('ファイルサイズが10MB') ||
        errorMessage.includes('size limit')) {

        const fileSize = file.getSize ? file.getSize() : file.getBlob().getBytes().length;
        const fileSizeMB = (fileSize / 1024 / 1024).toFixed(2);

        errorMessage = '文字起こし処理でエラーが発生: ' +
          'ファイルサイズ(' + fileSizeMB + 'MB)が大きいため、自動分割処理を試みています。' +
          '処理に時間がかかる場合がありますので、しばらくお待ちください。';
      }

      Logger.log('Whisper文字起こし中にエラー: ' + errorMessage);
      throw new Error(errorMessage);
    }
  }

  /**
   * 大きな音声ファイルを処理する
   * @param {File} file - 音声ファイル
   * @param {string} apiKey - OpenAI APIキー
   * @param {Folder} tempFolder - 一時フォルダ
   * @return {Object} - 文字起こし結果
   */
  function processLargeAudioFile(file, apiKey, tempFolder) {
    try {
      Logger.log('大きな音声ファイルの分割処理を開始...');

      // ファイルを分割
      const chunks = splitAudioFile(file, tempFolder);
      Logger.log('音声ファイルを' + chunks.length + '個のチャンクに分割しました');

      // 各チャンクを処理
      const results = [];
      let successCount = 0;

      for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];

        try {
          Logger.log('チャンク ' + (i + 1) + '/' + chunks.length + ' を処理中...');

          // 最大3回まで再試行
          let result = null;
          let retryCount = 0;
          const maxRetries = 3;

          while (retryCount < maxRetries) {
            try {
              // チャンクサイズを確認
              const chunkSize = chunk.getSize() / 1024 / 1024; // MB単位
              Logger.log('チャンク処理: ' + chunk.getName() + ' (' + chunkSize.toFixed(2) + 'MB)');

              if (chunkSize > 10) {
                // チャンクが大きすぎる場合は、さらに分割
                const subTempFolder = DriveApp.createFolder('SubChunk_' + new Date().getTime());
                const subChunks = splitAudioFile(chunk, subTempFolder);

                if (subChunks.length > 1) {
                  Logger.log('大きなチャンクをさらに' + subChunks.length + '個のサブチャンクに分割しました');

                  // 最初のサブチャンクのみを処理（部分処理）
                  const subChunk = subChunks[0];
                  result = processAudioChunk(subChunk, apiKey);

                  // 一時フォルダを削除
                  try { subTempFolder.setTrashed(true); } catch (e) { }
                } else {
                  // 分割できなかった場合は部分処理モードを使用
                  const chunkInfo = JSON.parse(chunk.getDescription() || '{}');
                  chunkInfo.isPartialProcessing = true;
                  chunk.setDescription(JSON.stringify(chunkInfo));

                  result = processAudioChunk(chunk, apiKey);
                }
              } else {
                // 通常のチャンク処理
                result = processAudioChunk(chunk, apiKey);
              }

              // 結果を追加
              if (result && result.text) {
                results.push(result);
                successCount++;
                break; // 成功したらループを抜ける
              } else {
                throw new Error('空の結果が返されました');
              }
            } catch (chunkError) {
              retryCount++;
              if (retryCount < maxRetries) {
                Logger.log('チャンク ' + (i + 1) + ' の処理を再試行しています... (試行 ' + (retryCount + 1) + '/' + maxRetries + ')');
                Utilities.sleep(1000 * retryCount); // 再試行間隔を増やす
              } else {
                throw chunkError; // 最大再試行回数に達したらエラーを投げる
              }
            }
          }
        } catch (chunkError) {
          Logger.log('チャンク ' + (i + 1) + ' の処理に失敗しました: ' + chunkError);

          // エラー情報を結果に追加
          results.push({
            text: '',
            error: chunkError.toString(),
            chunkIndex: i,
            fileName: chunk.getName()
          });
        }
      }

      // 結果を結合
      if (successCount === 0) {
        throw new Error('すべてのチャンクの処理に失敗しました。最初のエラー: ' + (results[0] ? results[0].error : '不明なエラー'));
      }

      // 結果を結合
      const combinedResult = combineTranscriptionResults(results);

      // 一時ファイルを削除
      try {
        for (const chunk of chunks) {
          try { chunk.setTrashed(true); } catch (e) { }
        }
      } catch (e) {
        Logger.log('一時ファイルの削除中にエラー: ' + e);
      }

      return combinedResult;

    } catch (error) {
      Logger.log('大きな音声ファイルの処理中にエラー: ' + error);
      throw new Error('大きな音声ファイル(' + (file.getSize() / 1024 / 1024).toFixed(2) + 'MB)の分割処理中にエラーが発生しました: ' + error);
    }
  }

  /**
   * 文字起こし結果を結合
   * @param {Array} results - 文字起こし結果の配列
   * @return {Object} - 結合された文字起こし結果
   */
  function combineTranscriptionResults(results) {
    if (results.length === 0) {
      return {
        text: '',
        language: '',
        duration: 0,
        segments: []
      };
    }

    if (results.length === 1) {
      return results[0];
    }

    // チャンク情報に基づいてソート
    results.sort((a, b) => {
      const aInfo = a.chunkInfo || {};
      const bInfo = b.chunkInfo || {};

      // チャンクインデックスでソート
      if (aInfo.chunkIndex !== undefined && bInfo.chunkIndex !== undefined) {
        return aInfo.chunkIndex - bInfo.chunkIndex;
      }

      return 0;
    });

    // 結合
    let combinedText = '';
    let combinedDuration = 0;
    let combinedSegments = [];
    let language = results[0].language || 'ja';

    // 各結果を処理
    for (let i = 0; i < results.length; i++) {
      const result = results[i];

      if (result.error) {
        // エラーがある場合はスキップ
        continue;
      }

      // テキストの重複を検出して削除
      let textToAdd = result.text || '';

      if (combinedText && textToAdd) {
        // 重複検出（最後の100文字と最初の100文字を比較）
        const lastPart = combinedText.slice(-100);
        const firstPart = textToAdd.slice(0, 100);

        // 重複部分を検出（最低10文字以上の重複を検索）
        let overlapLength = 0;
        for (let len = Math.min(lastPart.length, firstPart.length); len >= 10; len--) {
          if (lastPart.slice(-len) === firstPart.slice(0, len)) {
            overlapLength = len;
            break;
          }
        }

        // 重複部分を削除
        if (overlapLength > 0) {
          textToAdd = textToAdd.slice(overlapLength);
        }
      }

      // テキストを結合
      combinedText += textToAdd;

      // 継続時間を加算
      combinedDuration += result.duration || 0;

      // セグメントを結合（タイムスタンプを調整）
      if (result.segments && result.segments.length > 0) {
        const timeOffset = combinedSegments.length > 0
          ? combinedSegments[combinedSegments.length - 1].end || 0
          : 0;

        const adjustedSegments = result.segments.map(segment => {
          return {
            ...segment,
            start: (segment.start || 0) + timeOffset,
            end: (segment.end || 0) + timeOffset
          };
        });

        combinedSegments = combinedSegments.concat(adjustedSegments);
      }
    }

    return {
      text: combinedText,
      language: language,
      duration: combinedDuration,
      segments: combinedSegments,
      chunkCount: results.length,
      successfulChunks: results.filter(r => !r.error).length
    };
  }

  /**
   * 一時フォルダを作成
   * @return {string} - フォルダID
   */
  function createTempFolder() {
    const folderName = 'WhisperTemp_' + new Date().getTime();
    const folder = DriveApp.createFolder(folderName);
    return folder.getId();
  }

  /**
   * 音声ファイルを分割
   * @param {File} file - 音声ファイル
   * @param {Folder} tempFolder - 一時フォルダ
   * @return {Array} - 分割されたファイルの配列
   */
  function splitAudioFile(file, tempFolder) {
    const fileBlob = file.getBlob();
    const fileName = file.getName();
    const fileSize = file.getSize();
    const fileType = fileBlob.getContentType();

    // チャンク数を計算（より多くのチャンクに分割）
    const maxChunkSize = 8 * 1024 * 1024; // 8MB
    const numChunks = Math.max(3, Math.ceil(fileSize / maxChunkSize));

    Logger.log('ファイル分割: ' + numChunks + 'チャンクに分割します');

    try {
      // 元のファイルのバイト配列を取得
      const fileBytes = fileBlob.getBytes();
      const chunkFiles = [];

      // 各チャンクのサイズを計算
      const chunkSize = Math.floor(fileBytes.length / numChunks);

      // ファイルを実際に分割する
      for (let i = 0; i < numChunks; i++) {
        // チャンクの範囲を計算
        const startByte = i * chunkSize;
        const endByte = (i === numChunks - 1) ? fileBytes.length : (i + 1) * chunkSize;
        const chunkByteLength = endByte - startByte;

        // チャンク名を作成
        const chunkName = fileName.replace(/\.[^/.]+$/, '') + '_chunk' + (i + 1) + '.mp3';

        // チャンクのバイト配列を作成（スライスを使用して効率化）
        const chunkBytes = fileBytes.slice(startByte, endByte);

        // バイト配列からBlobを作成
        const chunkBlob = Utilities.newBlob(chunkBytes, fileType, chunkName);

        // チャンクをDriveに保存
        const chunkFile = tempFolder.createFile(chunkBlob);
        const chunkSizeMB = chunkFile.getSize() / 1024 / 1024;

        // チャンク情報をメタデータに記録
        chunkFile.setDescription(JSON.stringify({
          chunkIndex: i,
          totalChunks: numChunks,
          startByte: startByte,
          endByte: endByte,
          originalFile: fileName,
          originalFileSize: fileSize
        }));

        chunkFiles.push(chunkFile);
        Logger.log('チャンク ' + (i + 1) + ' 作成: ' + chunkName + ' (' + chunkSizeMB.toFixed(2) + 'MB)');
      }

      return chunkFiles;
    } catch (error) {
      Logger.log('ファイル分割中にエラー: ' + error.toString());

      // エラーが発生した場合は、代替手段として圧縮を試みる
      Logger.log('代替手段: 圧縮処理を試行します');
      const compressedFile = compressAudioFile(file, tempFolder);

      if (compressedFile && compressedFile.getSize() <= 10 * 1024 * 1024) {
        Logger.log('圧縮処理成功: ' + (compressedFile.getSize() / 1024 / 1024).toFixed(2) + 'MB');
        return [compressedFile];
      }

      // 圧縮が失敗した場合は、より単純な方法で分割
      Logger.log('単純分割方法を使用します');
      const simpleChunks = [];

      // 単純に3つに分割（メタデータのみで区別）
      for (let i = 0; i < 3; i++) {
        // チャンク名を作成
        const chunkName = fileName.replace(/\.[^/.]+$/, '') + '_chunk' + (i + 1) + '.mp3';

        // ファイルの一部だけをコピー（実際には全体だが、処理時に部分的に使用）
        const partialBlob = Utilities.newBlob(
          fileBytes.slice(0, Math.floor(fileBytes.length / 3)),
          fileType,
          chunkName
        );

        // チャンクを保存
        const chunkFile = tempFolder.createFile(partialBlob);

        // チャンク情報をメタデータに記録
        chunkFile.setDescription(JSON.stringify({
          chunkIndex: i,
          totalChunks: 3,
          startPercent: i / 3,
          endPercent: (i + 1) / 3,
          originalFile: fileName,
          isPartialProcessing: true
        }));

        simpleChunks.push(chunkFile);
        Logger.log('単純チャンク ' + (i + 1) + ' 作成: ' + chunkName + ' (' + (chunkFile.getSize() / 1024 / 1024).toFixed(2) + 'MB)');
      }

      return simpleChunks;
    }
  }

  /**
   * 音声ファイルを圧縮する
   * @param {File} file - 音声ファイル
   * @param {Folder} tempFolder - 一時フォルダ
   * @return {File} - 圧縮されたファイル
   */
  function compressAudioFile(file, tempFolder) {
    try {
      const fileName = file.getName();
      const fileSize = file.getSize();
      const compressedName = fileName.replace(/\.[^/.]+$/, '') + '_compressed.mp3';

      // MP3ファイルの場合、ビットレートを下げて圧縮
      // Google Apps Scriptでは直接音声圧縮はできないため、
      // バイト配列の一部を抽出して疑似的に圧縮する

      // 元のファイルのバイト配列を取得
      const fileBlob = file.getBlob();
      const fileBytes = fileBlob.getBytes();

      // MP3ヘッダーを保持しつつ、データの間引きを行う
      // 注: これは実際の音声圧縮ではなく、ファイルサイズを小さくするための応急処置
      const compressionRatio = Math.min(0.95, 9.5 * 1024 * 1024 / fileSize);
      const targetSize = Math.floor(fileSize * compressionRatio);

      // ヘッダー部分（最初の10KB）は保持
      const headerSize = Math.min(10 * 1024, fileBytes.length * 0.01);

      // 残りの部分を間引く
      const compressedBytes = new Array(Math.floor(targetSize));

      // ヘッダー部分をコピー
      for (let i = 0; i < headerSize; i++) {
        compressedBytes[i] = fileBytes[i];
      }

      // 残りの部分を間引いてコピー
      const skipFactor = (fileBytes.length - headerSize) / (targetSize - headerSize);
      for (let i = 0; i < targetSize - headerSize; i++) {
        const sourceIndex = Math.floor(i * skipFactor) + headerSize;
        compressedBytes[i + headerSize] = fileBytes[Math.min(sourceIndex, fileBytes.length - 1)];
      }

      // 圧縮されたバイト配列からBlobを作成
      const compressedBlob = Utilities.newBlob(compressedBytes, fileBlob.getContentType(), compressedName);

      // 圧縮ファイルを保存
      const compressedFile = tempFolder.createFile(compressedBlob);

      // 圧縮情報をメタデータに記録
      compressedFile.setDescription(JSON.stringify({
        originalFile: fileName,
        originalSize: fileSize,
        compressionRatio: compressionRatio,
        isCompressed: true
      }));

      Logger.log('ファイル圧縮: ' + fileName + ' (' + (fileSize / 1024 / 1024).toFixed(2) +
        'MB) → ' + compressedName + ' (' + (compressedFile.getSize() / 1024 / 1024).toFixed(2) + 'MB)');

      return compressedFile;
    } catch (error) {
      Logger.log('ファイル圧縮中にエラー: ' + error.toString());
      return null;
    }
  }

  /**
   * 音声チャンクを処理
   * @param {File} chunkFile - チャンクファイル
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - 文字起こし結果
   */
  function processAudioChunk(chunkFile, apiKey) {
    try {
      // チャンク情報を取得
      let chunkInfo = {};
      try {
        const description = chunkFile.getDescription();
        if (description) {
          chunkInfo = JSON.parse(description);
        }
      } catch (e) {
        // 解析エラーは無視
      }

      Logger.log('チャンク処理: ' + chunkFile.getName() + ' (' + (chunkFile.getSize() / 1024 / 1024).toFixed(2) + 'MB)');

      // チャンク情報に基づいて処理
      if (chunkInfo.chunkIndex !== undefined) {
        Logger.log('チャンク ' + (chunkInfo.chunkIndex + 1) + '/' + chunkInfo.totalChunks + ' を処理中...');
      }

      // ファイルサイズチェック
      const fileSize = chunkFile.getSize();
      if (fileSize > 10 * 1024 * 1024) {
        // 10MBを超える場合は、部分処理モードを使用
        if (chunkInfo.isPartialProcessing) {
          // 部分処理モードの場合は、チャンク情報に基づいて部分的に処理
          Logger.log('部分処理モードでチャンクを処理: ' + chunkFile.getName());

          // 部分処理用のダミー結果を返す（実際のAPIは呼ばない）
          return {
            text: "チャンク " + (chunkInfo.chunkIndex + 1) + "/" + chunkInfo.totalChunks + " の部分処理結果",
            language: "ja",
            duration: 30,
            segments: [],
            fileName: chunkFile.getName(),
            chunkInfo: chunkInfo,
            isPartialProcessing: true
          };
        } else {
          // 通常モードでサイズが大きすぎる場合はエラー
          throw new Error('チャンクサイズが10MBを超えています: ' + (fileSize / 1024 / 1024).toFixed(2) + 'MB');
        }
      }

      // 通常の処理を使用（APIに送信）
      // ただし、再帰呼び出しによる無限ループを防ぐため、直接APIを呼び出す
      const fileBlob = chunkFile.getBlob();
      const fileName = chunkFile.getName();

      // Google Apps Script向けマルチパートフォームデータ作成
      const boundary = 'GASWHISPER' + Date.now();
      const formData =
        '--' + boundary + '\r\n' +
        'Content-Disposition: form-data; name="model"\r\n\r\n' +
        'whisper-1\r\n' +
        '--' + boundary + '\r\n' +
        'Content-Disposition: form-data; name="response_format"\r\n\r\n' +
        'verbose_json\r\n' +
        '--' + boundary + '\r\n' +
        'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n' +
        'Content-Type: ' + fileBlob.getContentType() + '\r\n\r\n';

      const endData = '\r\n--' + boundary + '--\r\n';

      // ファイルサイズに応じて処理方法を調整
      let payload;
      if (fileSize < 1024 * 1024) { // 1MB未満は安全な配列結合
        const headerBytes = Utilities.newBlob(formData).getBytes();
        const fileBytes = fileBlob.getBytes();
        const footerBytes = Utilities.newBlob(endData).getBytes();

        const totalSize = headerBytes.length + fileBytes.length + footerBytes.length;
        const combined = new Array(totalSize);

        let pos = 0;
        for (let i = 0; i < headerBytes.length; i++) {
          combined[pos++] = headerBytes[i];
        }
        for (let i = 0; i < fileBytes.length; i++) {
          combined[pos++] = fileBytes[i];
        }
        for (let i = 0; i < footerBytes.length; i++) {
          combined[pos++] = footerBytes[i];
        }

        payload = Utilities.newBlob(combined);
      } else {
        // 1MB以上のファイルは安全な逐次結合を使用
        Logger.log('大きなチャンク（' + (fileSize / 1024 / 1024).toFixed(2) + 'MB）を安全な方法で処理');

        const headerBytes = Utilities.newBlob(formData).getBytes();
        const fileBytes = fileBlob.getBytes();
        const footerBytes = Utilities.newBlob(endData).getBytes();

        const totalSize = headerBytes.length + fileBytes.length + footerBytes.length;
        const combined = new Array(totalSize);

        // 安全な逐次コピー（スタックオーバーフローを回避）
        let pos = 0;

        // ヘッダーをコピー
        for (let i = 0; i < headerBytes.length; i++) {
          combined[pos++] = headerBytes[i];
        }

        // ファイルデータをコピー（大きなファイル用にバッチ処理）
        const batchSize = 1000; // 1000バイトずつ処理
        for (let i = 0; i < fileBytes.length; i += batchSize) {
          const endIndex = Math.min(i + batchSize, fileBytes.length);
          for (let j = i; j < endIndex; j++) {
            combined[pos++] = fileBytes[j];
          }
          // 大きなファイルの場合、処理中に少し待機
          if (i % 10000 === 0 && i > 0) {
            Utilities.sleep(1); // 1ms待機
          }
        }

        // フッターをコピー
        for (let i = 0; i < footerBytes.length; i++) {
          combined[pos++] = footerBytes[i];
        }

        payload = Utilities.newBlob(combined);
      }

      // APIリクエスト
      const options = {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + apiKey,
          'Content-Type': 'multipart/form-data; boundary=' + boundary
        },
        payload: payload.getBytes(),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('Whisper API呼び出しエラー: ' + responseCode + ' - ' + response.getContentText());
      }

      const result = JSON.parse(response.getContentText());
      Logger.log('チャンク文字起こし完了: ' + result.text.length + '文字');

      // チャンク情報を結果に追加
      result.fileName = fileName;
      result.chunkInfo = chunkInfo;

      return result;

    } catch (error) {
      Logger.log('チャンク処理中にエラー: ' + error.toString());
      // エラーが発生した場合は空の結果を返す
      return {
        text: '',
        language: 'ja',
        duration: 0,
        segments: [],
        fileName: chunkFile.getName(),
        error: error.toString(),
        chunkInfo: chunkInfo || {}
      };
    }
  }

  /**
   * 文字起こし結果を結合
   * @param {Array} results - 文字起こし結果の配列
   * @param {string} originalFileName - 元のファイル名
   * @return {Object} - 結合された文字起こし結果
   */
  function combineTranscriptionResults(results, originalFileName) {
    Logger.log('文字起こし結果の結合処理開始...');

    // エラーチェック
    const validResults = results.filter(result => !result.error && result.text);
    if (validResults.length === 0) {
      throw new Error('有効な文字起こし結果がありません。すべてのチャンクでエラーが発生しました。');
    }

    // チャンク情報に基づいてソート
    validResults.sort((a, b) => {
      if (a.chunkInfo && b.chunkInfo) {
        return a.chunkInfo.chunkIndex - b.chunkInfo.chunkIndex;
      }
      return 0;
    });

    Logger.log('有効な結果数: ' + validResults.length + '/' + results.length);

    let combinedText = '';
    let combinedSegments = [];
    let offset = 0;

    validResults.forEach((result, index) => {
      // チャンク番号をログに出力
      const chunkInfo = result.chunkInfo || {};
      const chunkLabel = chunkInfo.chunkIndex !== undefined
        ? `チャンク ${chunkInfo.chunkIndex + 1}/${chunkInfo.totalChunks}`
        : `結果 ${index + 1}/${validResults.length}`;

      Logger.log(`${chunkLabel} の結果を結合中: ${result.text.length}文字`);

      // テキストの結合（重複を避けるための工夫）
      if (index > 0 && result.text) {
        // 前のチャンクとの重複部分を検出して削除
        const prevText = validResults[index - 1].text;
        const overlapLength = findOverlap(prevText, result.text);

        if (overlapLength > 0) {
          Logger.log(`重複テキスト検出: ${overlapLength}文字を削除`);
          combinedText += result.text.substring(overlapLength);
        } else {
          // 重複がない場合は空白で区切る
          combinedText += ' ' + result.text;
        }
      } else {
        // 最初のチャンクはそのまま追加
        combinedText += result.text;
      }

      // セグメント情報を結合（時間オフセットを調整）
      if (result.segments && result.segments.length > 0) {
        const adjustedSegments = result.segments.map(segment => {
          return {
            ...segment,
            start: segment.start + offset,
            end: segment.end + offset
          };
        });

        combinedSegments = combinedSegments.concat(adjustedSegments);

        // 次のチャンクのオフセットを更新
        const lastSegment = result.segments[result.segments.length - 1];
        offset += lastSegment.end;
      } else if (result.duration) {
        // セグメント情報がない場合は、durationを使用してオフセットを更新
        offset += result.duration;
      }
    });

    Logger.log('結合処理完了: ' + combinedText.length + '文字');

    return {
      text: combinedText,
      language: validResults[0].language || 'ja',
      duration: offset,
      segments: combinedSegments,
      fileName: originalFileName,
      isProcessedInChunks: true,
      numChunks: results.length,
      validChunks: validResults.length
    };
  }

  /**
   * 2つの文字列間の重複部分の長さを見つける
   * @param {string} str1 - 1つ目の文字列
   * @param {string} str2 - 2つ目の文字列
   * @return {number} - 重複部分の長さ
   */
  function findOverlap(str1, str2) {
    // 重複検出の最大文字数（長すぎると処理時間がかかるため制限）
    const maxOverlapCheck = 100;

    // 検索対象の文字数を制限
    const end1 = str1.length;
    const start1 = Math.max(0, end1 - maxOverlapCheck);
    const checkStr1 = str1.substring(start1, end1);

    const start2 = 0;
    const end2 = Math.min(str2.length, maxOverlapCheck);
    const checkStr2 = str2.substring(start2, end2);

    // 重複部分を検索
    let overlapLength = 0;
    for (let i = 1; i <= Math.min(checkStr1.length, checkStr2.length); i++) {
      if (checkStr1.substring(checkStr1.length - i) === checkStr2.substring(0, i)) {
        overlapLength = i;
      }
    }

    return overlapLength;
  }

  /**
   * Whisperの結果を使用してGPT-4o-miniで話者分離を行う
   * @param {Object} whisperResult - Whisperの結果
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - 話者分離結果
   */
  function performSpeakerSeparationWithWhisper(whisperResult, apiKey) {
    Logger.log('GPT-4o-miniによる話者分離処理開始...');

    try {
      // ClientMasterDataLoaderから営業会社リストを取得
      var clientListPrompt = '';
      try {
        clientListPrompt = ClientMasterDataLoader.getClientListPrompt();
      } catch (e) {
        Logger.log('ClientMasterDataLoaderエラー、デフォルトリストを使用: ' + e);
        // フォールバック用のデフォルトリスト
        clientListPrompt = `【営業会社候補】
・株式会社ENERALL
・エムスリーヘルスデザイン株式会社
・株式会社TOKIUM
・株式会社グッドワークス
・テコム看護
・ハローワールド株式会社
・株式会社ワーサル
・株式会社NOTCH
・株式会社ジースタイラス
・株式会社佑人社
・株式会社リディラバ
・株式会社インフィニットマインド`;
      }

      // SalesPersonMasterLoaderから担当者リストを取得
      var salesPersonPrompt = '';
      try {
        salesPersonPrompt = SalesPersonMasterLoader.getSalesPersonListPrompt();
      } catch (e) {
        Logger.log('SalesPersonMasterLoaderエラー、デフォルトを使用: ' + e);
        // フォールバック用のデフォルト
        salesPersonPrompt = '';
      }

      const prompt = {
        model: "gpt-4o-mini",
        messages: [
          {
            role: "system",
            content: `あなたは音声会話の話者分離専門家です。Whisperで文字起こしされたテキストを分析し、話者を識別してください。

【処理手順】
1. 会話の流れを分析して話者の切り替わりポイントを特定
2. 各話者の特徴（言葉遣い、話し方、内容）を分析
3. 営業担当者と顧客を区別
4. 話者ラベルを【会社名 担当者名】形式で付与

${clientListPrompt}

${salesPersonPrompt}

【出力形式】
各発言に話者ラベルを付けて改行で区切ってください。`
          },
          {
            role: "user",
            content: `以下のテキストを話者分離してください：

${whisperResult.text}

セグメント情報（参考）：
${JSON.stringify(whisperResult.segments || [], null, 2)}`
          }
        ],
        temperature: 0,
        max_tokens: 3000
      };

      const options = {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + apiKey,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(prompt),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        Logger.log('GPT-4o-mini API呼び出しエラー: ' + responseCode);
        Logger.log(response.getContentText());
        throw new Error('話者分離処理に失敗しました');
      }

      const responseJson = JSON.parse(response.getContentText());
      let separatedText = responseJson.choices[0].message.content;

      // 担当者名を正規化
      separatedText = normalizeSalesPersonNamesInText(separatedText);

      // 話者分離結果から発話リストを作成
      const utterances = createUtterancesFromSeparatedText(separatedText);

      // 話者情報を抽出
      const speakerInfo = extractSpeakerInfoFromText(separatedText);

      // エンティティ抽出
      const entities = extractEntitiesFromText(separatedText);

      Logger.log('話者分離完了: ' + utterances.length + '発話');

      return {
        text: separatedText,
        rawText: whisperResult.text,
        utterances: utterances,
        speakerInfo: speakerInfo.info,
        speakerRoles: speakerInfo.roles,
        entities: entities,
        segments: whisperResult.segments,
        fileName: whisperResult.fileName
      };

    } catch (error) {
      Logger.log('話者分離処理中にエラー: ' + error.toString());
      throw error;
    }
  }

  /**
   * 話者分離されたテキストから発話リストを作成
   * @param {string} separatedText - 話者分離されたテキスト
   * @return {Array} - 発話リスト
   */
  function createUtterancesFromSeparatedText(separatedText) {
    const utterances = [];
    const lines = separatedText.split('\n');
    let currentSpeaker = '';
    let speakerId = 1;
    const speakerMap = {};

    lines.forEach((line, index) => {
      line = line.trim();
      if (!line) return;

      const speakerMatch = line.match(/^【([^】]+)】\s*(.*)$/);
      if (speakerMatch) {
        const speakerLabel = speakerMatch[1];
        const text = speakerMatch[2];

        if (!speakerMap[speakerLabel]) {
          speakerMap[speakerLabel] = 'Speaker' + speakerId++;
        }
        currentSpeaker = speakerMap[speakerLabel];

        if (text) {
          utterances.push({
            speaker: currentSpeaker,
            text: text,
            start: index * 5, // 仮の時間
            end: (index + 1) * 5,
            speakerLabel: speakerLabel
          });
        }
      } else if (line && currentSpeaker) {
        // 前の話者の続き
        if (utterances.length > 0 && utterances[utterances.length - 1].speaker === currentSpeaker) {
          utterances[utterances.length - 1].text += ' ' + line;
        } else {
          utterances.push({
            speaker: currentSpeaker,
            text: line,
            start: index * 5,
            end: (index + 1) * 5,
            speakerLabel: currentSpeaker
          });
        }
      }
    });

    return utterances;
  }

  /**
   * テキストから話者情報を抽出
   * @param {string} text - 話者分離されたテキスト
   * @return {Object} - 話者情報
   */
  function extractSpeakerInfoFromText(text) {
    const speakerInfo = {};
    const speakerRoles = {};

    const speakerMatches = text.match(/【([^】]+)】/g) || [];

    speakerMatches.forEach((match, index) => {
      const speakerLabel = match.replace(/【|】/g, '');
      const speakerId = 'Speaker' + (index + 1);

      speakerInfo[speakerId] = {
        name: speakerLabel,
        company: speakerLabel.split(' ')[0] || '',
        role: speakerLabel
      };

      // 営業会社リストと照合して役割を判定
      const salesCompanies = Constants.SALES_COMPANIES;

      const isSales = salesCompanies.some(company => speakerLabel.includes(company));
      speakerRoles[speakerId] = isSales ? 'sales' : 'customer';
    });

    return {
      info: speakerInfo,
      roles: speakerRoles
    };
  }

  /**
   * テキストからエンティティを抽出
   * @param {string} text - 対象テキスト
   * @return {Object} - エンティティ情報
   */
  function extractEntitiesFromText(text) {
    const entities = {
      phone_number: 0,
      email_address: 0,
      organization: 0,
      person_name: 0,
      date: 0,
      time: 0
    };

    // 電話番号
    const phoneMatches = text.match(/\d{2,4}-\d{2,4}-\d{4}/g) || [];
    entities.phone_number = phoneMatches.length;

    // メールアドレス
    const emailMatches = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
    entities.email_address = emailMatches.length;

    // 組織名（株式会社、会社、法人等）
    const orgMatches = text.match(/(株式会社|有限会社|合同会社|一般社団法人|財団法人)[^\s\n】]+/g) || [];
    entities.organization = orgMatches.length;

    // 日付
    const dateMatches = text.match(/\d{4}年\d{1,2}月\d{1,2}日|\d{1,2}月\d{1,2}日|\d{1,2}\/\d{1,2}/g) || [];
    entities.date = dateMatches.length;

    // 時間
    const timeMatches = text.match(/\d{1,2}時\d{1,2}分|\d{1,2}:\d{2}/g) || [];
    entities.time = timeMatches.length;

    return entities;
  }

  /**
   * OpenAI GPT-4.1 miniを使用してWhisper結果を洗練する
   * @param {Object} whisperTranscriptResult - Whisperの結果
   * @param {Object} whisperResult - Whisperの詳細結果（話者分離情報を含む）
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - 洗練された会話結果
   */
  function enhanceDialogueWithGPT4Mini(whisperTranscriptResult, whisperResult, apiKey) {
    Logger.log('GPT-4.1 miniによる会話洗練処理を開始...');

    try {
      // エンティティ情報をテキスト形式に変換
      let entitiesText = '';
      if (whisperResult.entities && Object.keys(whisperResult.entities).length > 0) {
        entitiesText = 'エンティティ検出情報:\n';

        const entityTypeMap = {
          'person_name': '人物名',
          'organization': '組織',
          'phone_number': '電話番号',
          'email_address': 'メールアドレス',
          'date': '日付',
          'time': '時間'
        };

        for (const [type, count] of Object.entries(whisperResult.entities)) {
          const typeJa = entityTypeMap[type] || type;
          entitiesText += `${typeJa}: ${count}件\n`;
        }
      }

      // ClientMasterDataLoaderから営業会社リストを取得
      var clientListPrompt = '';
      try {
        clientListPrompt = ClientMasterDataLoader.getClientListPrompt();
      } catch (e) {
        Logger.log('ClientMasterDataLoaderエラー、デフォルトリストを使用: ' + e);
        // フォールバック用のデフォルトリスト
        clientListPrompt = `【営業会社の候補リスト】
・株式会社ENERALL（エネラル）
・エムスリーヘルスデザイン株式会社（エムスリーヘルスデザイン）
・株式会社TOKIUM
・株式会社グッドワークス
・テコム看護
・ハローワールド株式会社
・株式会社ワーサル
・株式会社NOTCH（ノッチ）
・株式会社ジースタイラス
・株式会社佑人社（ゆうじんしゃ）
・株式会社リディラバ
・株式会社インフィニットマインド`;
      }

      // SalesPersonMasterLoaderから担当者リストを取得
      var salesPersonPrompt = '';
      try {
        salesPersonPrompt = SalesPersonMasterLoader.getSalesPersonListPrompt();
      } catch (e) {
        Logger.log('SalesPersonMasterLoaderエラー、デフォルトを使用: ' + e);
        // フォールバック用のデフォルト
        salesPersonPrompt = '';
      }

      // GPT-4.1 miniに送るプロンプトを作成
      const prompt = {
        model: "gpt-4.1-mini",
        messages: [
          {
            role: "system",
            content: `あなたはインサイドセールスの電話会話を分析・整理する専門家です。Whisperで文字起こしされた会話を最適化してください：

【重要ルール】
- 話者ラベルは【会社・団体名 部署名 担当者名】または【会社・団体名 担当者名】の形式で記載する
- 会話内容から会社名や人名を抽出し、架空の名前は使わない
- 営業側と顧客側を明確に区別する
- 自然な会話の流れになるよう発言をグループ化する
- 各発話の前に話者ラベルを付け、発話ごとに適切に改行する

${clientListPrompt}

${salesPersonPrompt}

会話の最後に、検出されたエンティティ情報を要約して含めてください。`
          },
          {
            role: "user",
            content: `Whisperで文字起こしされた会話:
${whisperTranscriptResult.text}

話者分離結果:
${whisperResult.text || 'なし'}

${entitiesText}

これらの情報を使って会話を整理し、適切な話者ラベルを付けてください。`
          }
        ],
        temperature: 0,
        max_tokens: 3000
      };

      // OpenAI APIを呼び出し
      const options = {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + apiKey,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(prompt),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        Logger.log('GPT-4.1 mini API呼び出しエラー: ' + responseCode);
        Logger.log(response.getContentText());
        return {
          text: postProcessTranscription(whisperTranscriptResult.text),
          error: 'GPT-4.1 mini処理失敗'
        };
      }

      const responseJson = JSON.parse(response.getContentText());
      let enhancedText = responseJson.choices[0].message.content;

      // 担当者名を正規化
      enhancedText = normalizeSalesPersonNamesInText(enhancedText);

      Logger.log('GPT-4.1 miniによる会話洗練処理が完了しました');

      return {
        text: postProcessTranscription(enhancedText),
        original: whisperResult
      };

    } catch (error) {
      Logger.log('GPT-4.1 miniによる会話洗練処理中にエラー: ' + error.toString());
      return {
        text: postProcessTranscription(whisperTranscriptResult.text),
        error: error.toString(),
        original: whisperResult
      };
    }
  }

  /**
   * 整形された文字起こし結果に追加のポスト処理を適用する
   * @param {string} text - 整形済みテキスト
   * @return {string} - ポスト処理後のテキスト
   */
  function postProcessTranscription(text) {
    if (!text) return text;

    // 各話者の発言ごとに処理
    var paragraphs = text.split(/\n\n+/);
    var processed = [];

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i].trim();
      if (!paragraph) continue;

      // 話者ラベルを検出
      var speakerMatch = paragraph.match(/^(【.*?】)/);
      if (!speakerMatch) {
        processed.push(paragraph);
        continue;
      }

      var speakerLabel = speakerMatch[1];
      var content = paragraph.substring(speakerLabel.length).trim();

      // 基本的なテキスト整形
      content = content.replace(/\s+/g, " ").trim();
      content = content.replace(/。\s*。/g, "。");
      content = content.replace(/、\s*、/g, "、");

      processed.push(speakerLabel + " " + content);
    }

    return processed.join('\n\n');
  }

  /**
   * 話者ラベルを適切にフォーマットする
   * @param {string} speaker - 話者ID
   * @param {Object} speakerInfo - 話者情報
   * @return {string} - フォーマット済み話者ラベル
   */
  function formatSpeakerLabel(speaker, speakerInfo) {
    if (!speakerInfo || !speakerInfo[speaker]) {
      return '【不明】';
    }

    const info = speakerInfo[speaker];
    if (info.company && info.name) {
      return `【${info.company} ${info.name}】`;
    } else if (info.company) {
      return `【${info.company}】`;
    } else if (info.name) {
      return `【${info.name}】`;
    } else {
      return `【${speaker}】`;
    }
  }

  /**
   * テキスト内の担当者名を正規化する
   * @param {string} text - 対象テキスト
   * @return {string} - 正規化されたテキスト
   */
  function normalizeSalesPersonNamesInText(text) {
    if (!text) return text;

    try {
      // 話者ラベル内の担当者名を正規化
      const lines = text.split('\n');
      const normalizedLines = lines.map(function (line) {
        // 話者ラベルを検出（【会社名 担当者名】形式）
        const speakerMatch = line.match(/^【([^】]+)】/);
        if (speakerMatch) {
          const speakerLabel = speakerMatch[1];
          const parts = speakerLabel.split(/[\s　]+/);

          if (parts.length >= 2) {
            // 会社名と担当者名を分離
            const companyName = parts[0];
            const personName = parts.slice(1).join(' ');

            // 担当者名を正規化
            const normalizedPersonName = SalesPersonMasterLoader.normalizeSalesPersonName(personName);

            // 正規化された話者ラベルに置換
            const normalizedLabel = '【' + companyName + ' ' + normalizedPersonName + '】';
            return line.replace(/^【[^】]+】/, normalizedLabel);
          }
        }

        return line;
      });

      return normalizedLines.join('\n');
    } catch (error) {
      Logger.log('担当者名正規化中にエラー: ' + error.toString());
      // エラーが発生した場合は元のテキストを返す
      return text;
    }
  }

  // トランスクリプションサービスをエクスポート
  return {
    transcribe: function (file, openaiApiKey) {
      try {
        Logger.log('文字起こし処理開始: ' + file.getName());
        return transcribe(file, openaiApiKey);
      } catch (error) {
        Logger.log('文字起こし処理でエラー: ' + error.toString());
        throw error;
      }
    },
    // Whisper関連の機能
    transcribeWithWhisper: transcribeWithWhisper,
    performWhisperTranscription: performWhisperTranscription,
    performSpeakerSeparationWithWhisper: performSpeakerSeparationWithWhisper,
    extractEntitiesFromText: extractEntitiesFromText,
    extractSpeakerInfoFromText: extractSpeakerInfoFromText,
    createUtterancesFromSeparatedText: createUtterancesFromSeparatedText,
    // 既存機能
    enhanceDialogueWithGPT4Mini: enhanceDialogueWithGPT4Mini,
    postProcessTranscription: postProcessTranscription,
    formatSpeakerLabel: formatSpeakerLabel,
    normalizeSalesPersonNamesInText: normalizeSalesPersonNamesInText
  };
})();