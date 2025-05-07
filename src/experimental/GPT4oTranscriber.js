/**
 * GPT-4o Transcribeを使用した文字起こし実験用スクリプト
 * - OpenAIのAPIを使用して音声ファイルから文字起こしを行う
 * - AssemblyAIで話者分離情報を取得し、両者をマージ
 */

// スクリプトプロパティから設定値を取得
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    OPENAI_API_KEY: scriptProperties.getProperty('OPENAI_API_KEY'),
    ASSEMBLYAI_API_KEY: scriptProperties.getProperty('ASSEMBLYAI_API_KEY'),
    TEST_FOLDER_ID: scriptProperties.getProperty('TEST_FOLDER_ID')
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
  if (!config.ASSEMBLYAI_API_KEY) {
    throw new Error('AssemblyAI APIキーが設定されていません。スクリプトプロパティで設定してください。');
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

  // ファイルを処理（音声ファイルのみを抽出）
  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // デバッグ情報
    Logger.log('フォルダ内ファイル: ' + fileName + ' (' + mimeType + ')');

    // すでに処理済みのファイルもテスト用に処理する（一時的な措置）
    // if (fileName.includes('processed_')) {
    //   Logger.log('処理済みファイルをスキップ: ' + fileName);
    //   continue;
    // }

    // 音声ファイルのみを処理対象とする（ファイル名の制限を緩和）
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
    // GPT-4o Transcribeで文字起こし
    const openaiResult = transcribeWithGPT4o(latestFile, config.OPENAI_API_KEY);
    Logger.log('GPT-4o Transcribe処理完了');

    try {
      // AssemblyAIで話者分離
      const assemblyAIResult = transcribeWithAssemblyAI(latestFile, config.ASSEMBLYAI_API_KEY);
      Logger.log('AssemblyAI処理完了');

      // 結果をマージ
      const mergedResult = mergeTranscriptions(openaiResult, assemblyAIResult);
      Logger.log('マージ処理完了');

      // 処理結果をログに出力
      Logger.log('======= 処理結果 =======');
      Logger.log(mergedResult.text);

      // ファイル名の先頭に 'processed_' を付けて処理済みとしてマーク（テスト中は無効化）
      // latestFile.setName('processed_' + latestFile.getName());

      return mergedResult;
    } catch (assemblyError) {
      // AssemblyAIでエラーが発生した場合、別の形式で試す
      Logger.log('AssemblyAI処理中にエラーが発生しました: ' + assemblyError.toString());
      Logger.log('別の形式でリクエストを試みます...');

      // 別の形式でリクエストを試す
      const altSuccess = tryAlternativeAssemblyAIRequest(latestFile, config.ASSEMBLYAI_API_KEY);

      if (altSuccess) {
        Logger.log('代替リクエスト方式が成功しました。デバッグ用のログを確認してください。');
      } else {
        Logger.log('代替リクエスト方式も失敗しました。さらにV3 APIを試します...');

        // V3 APIも試す
        const v3Success = tryAssemblyAIV3Request(latestFile, config.ASSEMBLYAI_API_KEY);

        if (v3Success) {
          Logger.log('V3 API方式が成功しました。デバッグ用のログを確認してください。');
        } else {
          Logger.log('すべての試行が失敗しました。GPT-4oの結果のみを返します。');
        }
      }

      // GPT-4oの結果だけで簡易フォーマット
      const simpleResult = {
        text: formatSimpleText(openaiResult.text),
        error: assemblyError.toString(),
        original: {
          openai: openaiResult,
          assemblyAI: { error: assemblyError.toString() }
        }
      };

      // 処理結果をログに出力
      Logger.log('======= 処理結果 (GPT-4oのみ) =======');
      Logger.log(simpleResult.text);

      // ファイル名の先頭に 'processed_' を付けて処理済みとしてマーク（テスト中は無効化）
      // latestFile.setName('processed_' + latestFile.getName());

      return simpleResult;
    }
  } catch (error) {
    Logger.log('処理中にエラーが発生しました: ' + error.toString());
    throw error;
  }
}

/**
 * ファイルが音声ファイルかどうかを判定する
 * @param {string} fileName - ファイル名
 * @param {string} mimeType - MIMEタイプ
 * @return {boolean} - 音声ファイルならtrue
 */
function isAudioFile(fileName, mimeType) {
  // デバッグ情報
  Logger.log('ファイル判定: ' + fileName + ', MIMEタイプ: ' + mimeType);

  // MIMEタイプによる判定
  if (mimeType && (
    mimeType.includes('audio/') ||
    mimeType.includes('video/') ||  // 動画ファイルも音声として扱う
    mimeType === 'application/octet-stream'  // バイナリファイルの場合は拡張子で判断
  )) {

    // 拡張子による判断（バイナリファイルの場合など）
    if (mimeType === 'application/octet-stream') {
      const ext = fileName.split('.').pop().toLowerCase();
      const isAudio = ['mp3', 'wav', 'm4a', 'aac', 'ogg', 'flac', 'mp4', 'mpeg', 'mpga', 'webm'].includes(ext);
      Logger.log('バイナリファイル拡張子判定: ' + fileName + ' => ' + (isAudio ? '音声ファイル' : '非音声ファイル'));
      return isAudio;
    }

    Logger.log('MIMEタイプに基づき音声ファイルと判定: ' + fileName);
    return true;
  }

  // 拡張子による判断（バックアップ）
  const ext = fileName.split('.').pop().toLowerCase();
  const isAudio = ['mp3', 'wav', 'm4a', 'aac', 'ogg', 'flac', 'mp4', 'mpeg', 'mpga', 'webm'].includes(ext);
  Logger.log('拡張子による判断: ' + fileName + ' => ' + (isAudio ? '音声ファイル' : '非音声ファイル'));
  return isAudio;
}

/**
 * GPT-4o Transcribeを使用して文字起こしを行う
 * @param {File} file - 音声ファイル
 * @param {string} apiKey - OpenAI APIキー
 * @return {Object} - 文字起こし結果
 */
function transcribeWithGPT4o(file, apiKey) {
  Logger.log('GPT-4o Transcribeでの文字起こし開始...');

  try {
    // ファイルのバイナリデータを取得
    const fileBlob = file.getBlob();
    const fileName = file.getName();

    // OpenAI APIのエンドポイントURL
    const url = 'https://api.openai.com/v1/audio/transcriptions';

    // マルチパートフォームデータの作成
    const boundary = Utilities.getUuid();
    const contentType = 'multipart/form-data; boundary=' + boundary;

    // ファイルパートの作成
    let data = '';
    data += '--' + boundary + '\r\n';
    data += 'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n';
    data += 'Content-Type: ' + fileBlob.getContentType() + '\r\n\r\n';

    // バイナリデータの前後にテキストパートを追加
    const fileBytes = fileBlob.getBytes();

    // modelパートの追加
    let modelPart = '\r\n--' + boundary + '\r\n';
    modelPart += 'Content-Disposition: form-data; name="model"\r\n\r\n';
    modelPart += 'gpt-4o-mini-transcribe' + '\r\n';

    // language部分の追加（日本語を指定）
    let langPart = '--' + boundary + '\r\n';
    langPart += 'Content-Disposition: form-data; name="language"\r\n\r\n';
    langPart += 'ja' + '\r\n';

    // responseフォーマット部分の追加（JSONを取得）
    let responsePart = '--' + boundary + '\r\n';
    responsePart += 'Content-Disposition: form-data; name="response_format"\r\n\r\n';
    responsePart += 'json' + '\r\n';

    // timestampsパラメータ（タイムスタンプを有効化）
    let timestampsPart = '--' + boundary + '\r\n';
    timestampsPart += 'Content-Disposition: form-data; name="timestamp_granularities[]"\r\n\r\n';
    timestampsPart += 'segment' + '\r\n';

    // 最後の境界を追加
    let endPart = '--' + boundary + '--\r\n';

    // パート結合のためのバイト配列を作成
    const head = Utilities.newBlob(data).getBytes();
    const modelPartBytes = Utilities.newBlob(modelPart).getBytes();
    const langPartBytes = Utilities.newBlob(langPart).getBytes();
    const responsePartBytes = Utilities.newBlob(responsePart).getBytes();
    const timestampsPartBytes = Utilities.newBlob(timestampsPart).getBytes();
    const tail = Utilities.newBlob(endPart).getBytes();

    // 全てのバイト配列を結合
    const payload = [].concat(
      head,
      fileBytes,
      modelPartBytes,
      langPartBytes,
      responsePartBytes,
      timestampsPartBytes,
      tail
    );

    // APIリクエストオプション
    const options = {
      method: 'post',
      contentType: contentType,
      headers: {
        'Authorization': 'Bearer ' + apiKey
      },
      payload: payload,
      muteHttpExceptions: true
    };

    // APIリクエスト送信
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    // レスポンスのログ記録
    Logger.log('GPT-4o APIレスポンスコード: ' + responseCode);

    if (responseCode !== 200) {
      Logger.log('エラーレスポンス: ' + response.getContentText());
      throw new Error('GPT-4o API呼び出しエラー: ' + responseCode);
    }

    // レスポンスをJSONとしてパース
    const responseJson = JSON.parse(response.getContentText());

    // ログに記録（全文出力するように変更）
    Logger.log('GPT-4o処理完了: 全文出力');
    Logger.log(responseJson.text);

    return {
      text: responseJson.text,
      segments: responseJson.segments || [], // セグメント情報が含まれていなくても空配列を返す
      language: responseJson.language,
      duration: responseJson.duration,
      fileName: fileName
    };

  } catch (error) {
    Logger.log('GPT-4o Transcribe処理中にエラー: ' + error.toString());
    throw new Error('GPT-4o Transcribe処理中にエラー: ' + error.toString());
  }
}

/**
 * AssemblyAIを使用して話者分離情報を取得する
 * @param {File} file - 音声ファイル
 * @param {string} apiKey - AssemblyAI APIキー
 * @return {Object} - 文字起こし結果（話者分離情報付き）
 */
function transcribeWithAssemblyAI(file, apiKey) {
  Logger.log('AssemblyAIでの話者分離処理開始...');

  try {
    // ファイルのバイナリデータを取得
    const fileBlob = file.getBlob();
    const fileName = file.getName();

    // Assembly AI APIのエンドポイントURL
    const uploadUrl = 'https://api.assemblyai.com/v2/upload';
    const transcriptUrl = 'https://api.assemblyai.com/v2/transcript';

    // ファイルをAssembly AIにアップロード
    const uploadOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/octet-stream'
      },
      payload: fileBlob.getBytes(),
      muteHttpExceptions: true
    };

    const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
    const uploadResponseCode = uploadResponse.getResponseCode();

    if (uploadResponseCode !== 200) {
      throw new Error('ファイルのアップロードに失敗しました。レスポンスコード: ' + uploadResponseCode);
    }

    const uploadResponseJson = JSON.parse(uploadResponse.getContentText());
    const audioUrl = uploadResponseJson.upload_url;

    // 文字起こしリクエストを送信（話者分離を有効化）
    // 最新のAPIドキュメントに基づいた形式に修正
    const transcriptOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        audio_url: audioUrl,
        language_code: 'ja',
        speaker_labels: true  // この形式で試す
      }),
      muteHttpExceptions: true
    };

    const transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);

    // レスポンスコードの確認
    const transcriptResponseCode = transcriptResponse.getResponseCode();
    Logger.log('AssemblyAI 文字起こしリクエストレスポンスコード: ' + transcriptResponseCode);

    if (transcriptResponseCode !== 200) {
      Logger.log('AssemblyAI 文字起こしリクエストエラー: ' + transcriptResponse.getContentText());
      throw new Error('AssemblyAI 文字起こしリクエストエラー: ' + transcriptResponseCode);
    }

    const transcriptResponseJson = JSON.parse(transcriptResponse.getContentText());

    // レスポンス全体を確認用にログ出力
    Logger.log('AssemblyAI 文字起こしリクエストレスポンス: ' + JSON.stringify(transcriptResponseJson));

    const transcriptId = transcriptResponseJson.id;

    // トランスクリプションIDの取得確認をログに出力
    Logger.log('AssemblyAI トランスクリプションID取得: ' + transcriptId);

    // IDが取得できない場合はエラー
    if (!transcriptId) {
      Logger.log('AssemblyAI エラー: トランスクリプションIDが取得できませんでした');
      Logger.log('レスポンス内容: ' + JSON.stringify(transcriptResponseJson));
      throw new Error('AssemblyAI トランスクリプションIDが取得できませんでした');
    }

    // 文字起こし結果のポーリング
    const pollingUrl = transcriptUrl + '/' + transcriptId;
    const pollingOptions = {
      method: 'get',
      headers: {
        'Authorization': apiKey
      },
      muteHttpExceptions: true
    };

    // ポーリング開始前に少し待機（AssemblyAI側の処理開始を待つ）
    Logger.log('AssemblyAI 処理開始待機中...');
    Utilities.sleep(10000);  // 10秒待機

    let pollingResponseJson;
    let status = '';
    const pollingInterval = 5000; // 5秒ごとにポーリングに変更（サーバー負荷軽減）
    const maxPollingAttempts = 120; // 最大ポーリング回数（約10分に延長）
    let pollingAttempts = 0;

    while (status !== 'completed' && status !== 'error' && pollingAttempts < maxPollingAttempts) {
      // ポーリング間隔を待機
      Utilities.sleep(pollingInterval);

      try {
        const pollingResponse = UrlFetchApp.fetch(pollingUrl, pollingOptions);
        const responseCode = pollingResponse.getResponseCode();

        // レスポンスコードをログに出力
        if (pollingAttempts % 5 === 0 || responseCode !== 200) {
          Logger.log('AssemblyAI ポーリングレスポンスコード: ' + responseCode);
        }

        if (responseCode !== 200) {
          const errorText = pollingResponse.getContentText();
          Logger.log('AssemblyAI ポーリングエラー: ' + errorText);

          // ID not foundエラーの場合、トランスクリプションIDが無効になっている可能性がある
          if (errorText.includes('not found')) {
            if (pollingAttempts > 5) {  // 何度か試行してもダメなら諦める
              throw new Error('AssemblyAI トランスクリプションIDが無効です: ' + transcriptId);
            }
          } else if (pollingAttempts > 10) {  // その他のエラーが続く場合
            throw new Error('AssemblyAI ポーリングエラー: レスポンスコード ' + responseCode);
          }

          // エラーの場合は少し長めに待機
          Utilities.sleep(10000);
          pollingAttempts++;
          continue;
        }

        pollingResponseJson = JSON.parse(pollingResponse.getContentText());
        status = pollingResponseJson.status;

        // 処理に進展があった場合のみログに記録
        if (status) {
          Logger.log('AssemblyAI文字起こし進捗: ステータス=' + status + ', 試行回数=' + pollingAttempts + '/' + maxPollingAttempts);
        } else {
          if (pollingAttempts % 5 === 0) {
            Logger.log('AssemblyAI文字起こし進捗: ステータス不明, 試行回数=' + pollingAttempts + '/' + maxPollingAttempts);
            Logger.log('レスポンス内容: ' + pollingResponse.getContentText().substring(0, 200) + '...');
          }
        }
      } catch (pollingError) {
        Logger.log('AssemblyAI ポーリング中の例外: ' + pollingError.toString());
        // エラーの場合は少し長めに待機してから再試行
        Utilities.sleep(10000);
      }

      pollingAttempts++;
    }

    // 処理中または完了でない場合でもエラーをスロー
    if (status !== 'completed') {
      throw new Error('AssemblyAI文字起こし処理がタイムアウトまたはエラーが発生しました。ステータス: ' + status);
    }

    // 話者数をログに記録
    const utteranceCount = pollingResponseJson.utterances ? pollingResponseJson.utterances.length : 0;
    const speakerCount = getSpeakerCount(pollingResponseJson.utterances || []);
    Logger.log('AssemblyAI処理完了: 話者数=' + speakerCount + ', 発話数=' + utteranceCount);

    return {
      text: pollingResponseJson.text || '',
      utterances: pollingResponseJson.utterances || [],
      fileName: fileName
    };

  } catch (error) {
    // エラー情報を詳細にログ出力
    Logger.log('AssemblyAI処理中に致命的エラー: ' + error.toString());
    Logger.log('スタックトレース: ' + (error.stack || 'スタックトレースなし'));

    // エラー情報を含むオブジェクトを返す
    return {
      error: error.toString(),
      errorDetails: error.stack,
      fileName: file.getName()
    };
  }
}

/**
 * 話者の数を抽出する
 * @param {Array} utterances - 発話情報の配列
 * @return {number} - 話者の数
 */
function getSpeakerCount(utterances) {
  const speakerSet = new Set();

  for (let i = 0; i < utterances.length; i++) {
    if (utterances[i].speaker) {
      speakerSet.add(utterances[i].speaker);
    }
  }

  return speakerSet.size;
}

/**
 * 両サービスの文字起こし結果をマージする
 * @param {Object} openaiResult - GPT-4o Transcribeの結果
 * @param {Object} assemblyAIResult - AssemblyAIの結果（話者分離情報付き）
 * @return {Object} - マージした結果
 */
function mergeTranscriptions(openaiResult, assemblyAIResult) {
  Logger.log('文字起こし結果のマージ処理開始...');

  try {
    // AssemblyAIの結果が空またはエラーの場合はOpenAIの結果のみ返す
    if (!assemblyAIResult || assemblyAIResult.error) {
      Logger.log('AssemblyAIの結果がありません。GPT-4oの結果のみを返します。');
      // 単純なフォーマットに変換して返す
      return {
        text: formatSimpleText(openaiResult.text),
        original: {
          openai: openaiResult,
          assemblyAI: assemblyAIResult || { error: 'No result' }
        }
      };
    }

    // AssemblyAIから話者情報を抽出
    const utterances = assemblyAIResult.utterances || [];

    // 話者分離情報がない場合は単純にOpenAIの結果を返す
    if (!utterances || utterances.length === 0) {
      Logger.log('AssemblyAIからの話者分離情報がありません。GPT-4oの結果をそのまま返します。');
      return openaiResult;
    }

    // GPT-4oのセグメント情報を取得
    // 注意: gpt-4o-transcribeがセグメント情報を返さない場合がある
    const segments = openaiResult.segments || [];

    // セグメント情報がない場合は代替処理
    if (!segments || segments.length === 0) {
      Logger.log('GPT-4oからのセグメント情報がありません。文章ベースのマージ処理を実行します。');

      // GPT-4o-miniによる高度なマージを試みる（OpenAI APIキーを使用）
      try {
        Logger.log('GPT-4o-miniによる高度なマージを試みます...');
        // getConfig()からAPIキーを再取得
        const config = getConfig();
        return enhanceDialogueWithGPT4Mini(openaiResult, assemblyAIResult, config.OPENAI_API_KEY);
      } catch (miniError) {
        Logger.log('GPT-4o-miniによるマージでエラー: ' + miniError.toString());
        Logger.log('通常の文章ベースマージにフォールバックします。');
        return textBasedMerge(openaiResult, assemblyAIResult);
      }
    }

    // 各タイムスタンプ区間ごとに話者情報を割り当て
    const speakerMap = createSpeakerTimeMap(utterances);

    // 話者IDを役割に変換（A->営業担当者、B->お客様など）
    const speakerRoles = identifySpeakerRoles(utterances);

    // 話者情報を付与したテキストを生成
    let mergedText = '';
    let currentSpeaker = null;
    let currentSegments = [];

    // セグメントごとに話者を判定してマージ
    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i];
      const startTime = segment.start;
      const endTime = segment.end;

      // この時間帯の話者を特定
      const speaker = findSpeakerAtTime(speakerMap, startTime);

      // 話者が変わったら出力
      if (speaker !== currentSpeaker) {
        // 前の話者のテキストを出力
        if (currentSpeaker !== null && currentSegments.length > 0) {
          const speakerLabel = formatSpeakerLabel(currentSpeaker, speakerRoles);
          const text = currentSegments.join(' ');
          mergedText += speakerLabel + ' ' + text + '\n\n';
        }

        // 新しい話者のセグメントを開始
        currentSpeaker = speaker;
        currentSegments = [segment.text];
      } else {
        // 同じ話者の場合は追加
        currentSegments.push(segment.text);
      }
    }

    // 最後の話者のテキストを出力
    if (currentSpeaker !== null && currentSegments.length > 0) {
      const speakerLabel = formatSpeakerLabel(currentSpeaker, speakerRoles);
      const text = currentSegments.join(' ');
      mergedText += speakerLabel + ' ' + text + '\n\n';
    }

    return {
      text: mergedText,
      original: {
        openai: openaiResult,
        assemblyAI: assemblyAIResult
      }
    };

  } catch (error) {
    Logger.log('マージ処理中にエラー: ' + error.toString());
    // エラー時はGPT-4oの結果を簡易フォーマットして返す
    return {
      text: formatSimpleText(openaiResult.text),
      error: error.toString(),
      original: {
        openai: openaiResult,
        assemblyAI: assemblyAIResult || { error: 'Processing failed' }
      }
    };
  }
}

/**
 * 単純なテキストを段落ごとにフォーマットする
 * @param {string} text - 元のテキスト
 * @return {string} - フォーマット済みテキスト
 */
function formatSimpleText(text) {
  // 空の場合は空文字を返す
  if (!text) return '';

  // 文章ごとに分割して整形
  const sentences = text.split(/(?<=[。．！？])\s*/);
  const paragraphs = [];
  let currentParagraph = '';

  // 3-4文ごとに段落を分ける
  for (let i = 0; i < sentences.length; i++) {
    currentParagraph += sentences[i];

    // 3-4文ごと、または最後の文の場合は段落を確定
    if ((i + 1) % 4 === 0 || i === sentences.length - 1) {
      paragraphs.push(currentParagraph.trim());
      currentParagraph = '';
    }
  }

  // 段落をつなげて返す
  return paragraphs.join('\n\n');
}

/**
 * 話者情報をタイムスタンプマップに変換
 * @param {Array} utterances - 発話情報の配列
 * @return {Array} - 時間ごとの話者マップ
 */
function createSpeakerTimeMap(utterances) {
  const speakerMap = [];

  for (let i = 0; i < utterances.length; i++) {
    const u = utterances[i];
    if (u.speaker && u.start && u.end) {
      speakerMap.push({
        start: u.start,
        end: u.end,
        speaker: u.speaker
      });
    }
  }

  // 開始時間でソート
  speakerMap.sort((a, b) => a.start - b.start);

  return speakerMap;
}

/**
 * 特定の時間における話者を見つける
 * @param {Array} speakerMap - 話者マップ
 * @param {number} time - 時間（ミリ秒）
 * @return {string|null} - 話者ID、見つからない場合はnull
 */
function findSpeakerAtTime(speakerMap, time) {
  for (let i = 0; i < speakerMap.length; i++) {
    if (time >= speakerMap[i].start && time <= speakerMap[i].end) {
      return speakerMap[i].speaker;
    }
  }

  // 近い時間帯の話者を探す（50ms以内）
  const tolerance = 50;
  for (let i = 0; i < speakerMap.length; i++) {
    if (Math.abs(time - speakerMap[i].start) <= tolerance ||
      Math.abs(time - speakerMap[i].end) <= tolerance) {
      return speakerMap[i].speaker;
    }
  }

  return null;
}

/**
 * 話者ラベルをフォーマットする
 * @param {string} speaker - 話者ID
 * @param {Object} speakerRoles - 話者の役割マップ
 * @return {string} - フォーマットされた話者ラベル
 */
function formatSpeakerLabel(speaker, speakerRoles) {
  if (!speaker) return '【不明】';

  // 役割情報がある場合
  if (speakerRoles && speakerRoles[speaker]) {
    return '【' + speakerRoles[speaker] + '】';
  }

  // デフォルトのフォーマット
  return '【話者' + speaker + '】';
}

/**
 * 話者の役割を特定する
 * @param {Array} utterances - 発話情報の配列
 * @return {Object} - 話者IDと役割のマッピング
 */
function identifySpeakerRoles(utterances) {
  // 話者ごとの発話回数と発話の長さを集計
  const speakerStats = {};
  const speakerUtterances = {};

  for (let i = 0; i < utterances.length; i++) {
    const u = utterances[i];
    const speaker = u.speaker;

    if (!speaker) continue;

    if (!speakerStats[speaker]) {
      speakerStats[speaker] = {
        count: 0,
        totalLength: 0,
        totalDuration: 0
      };
      speakerUtterances[speaker] = [];
    }

    speakerStats[speaker].count++;
    speakerStats[speaker].totalLength += u.text ? u.text.length : 0;
    speakerStats[speaker].totalDuration += (u.end - u.start);
    speakerUtterances[speaker].push(u.text || '');
  }

  // 役割を特定するためのヒューリスティック
  // 1. 話す比率が高い（長い）ほうが説明する側（営業担当者など）
  // 2. 質問が多い方がお客様の可能性が高い

  // 話者のリストを取得
  const speakers = Object.keys(speakerStats);

  // 話者が1人以下の場合は役割を特定しない
  if (speakers.length <= 1) {
    const result = {};
    if (speakers.length === 1) {
      result[speakers[0]] = '話者A';
    }
    return result;
  }

  // 話者が2人以上の場合、話す時間の長い方を「担当者」、短い方を「お客様」とする
  let staff = null;
  let customer = null;

  // 話す時間で比較
  if (speakers.length === 2) {
    if (speakerStats[speakers[0]].totalDuration > speakerStats[speakers[1]].totalDuration) {
      staff = speakers[0];
      customer = speakers[1];
    } else {
      staff = speakers[1];
      customer = speakers[0];
    }

    return {
      [staff]: '担当者',
      [customer]: 'お客様'
    };
  }

  // 3人以上の場合はA, B, C...と表記
  const result = {};
  for (let i = 0; i < speakers.length; i++) {
    result[speakers[i]] = '話者' + String.fromCharCode(65 + i); // A, B, C...
  }

  return result;
}

/**
 * セグメント情報がない場合に、テキストベースでマージする
 * @param {Object} openaiResult - GPT-4o Transcribeの結果
 * @param {Object} assemblyAIResult - AssemblyAIの結果
 * @return {Object} - マージした結果
 */
function textBasedMerge(openaiResult, assemblyAIResult) {
  Logger.log('テキストベースのマージ処理を実行します...');

  try {
    // AssemblyAIの結果が空またはエラーの場合はOpenAIの結果のみ返す
    if (!assemblyAIResult || assemblyAIResult.error) {
      return {
        text: formatSimpleText(openaiResult.text),
        original: {
          openai: openaiResult,
          assemblyAI: assemblyAIResult || { error: 'No result' }
        }
      };
    }

    // AssemblyAIから話者情報を抽出
    const utterances = assemblyAIResult.utterances || [];

    // 話者分離情報がない場合は単純にOpenAIの結果を返す
    if (!utterances || utterances.length === 0) {
      return {
        text: formatSimpleText(openaiResult.text),
        original: {
          openai: openaiResult,
          assemblyAI: assemblyAIResult
        }
      };
    }

    // 話者IDを役割に変換
    const speakerRoles = identifySpeakerRoles(utterances);

    // OpenAIのテキストを文ごとに分割
    const sentences = openaiResult.text.split(/(?<=[。．！？])\s*/);

    // 文章を話者のタイミングに基づいて分割
    const speakerSegments = splitTextBySpeakerTiming(sentences, utterances);

    // 話者情報を付与したテキストを生成
    let mergedText = '';

    for (let i = 0; i < speakerSegments.length; i++) {
      const segment = speakerSegments[i];
      const speakerLabel = formatSpeakerLabel(segment.speaker, speakerRoles);
      mergedText += speakerLabel + ' ' + segment.text + '\n\n';
    }

    return {
      text: mergedText,
      original: {
        openai: openaiResult,
        assemblyAI: assemblyAIResult
      }
    };

  } catch (error) {
    Logger.log('テキストベースのマージ処理中にエラー: ' + error.toString());
    // エラー時はGPT-4oの結果を簡易フォーマットして返す
    return {
      text: formatSimpleText(openaiResult.text),
      error: error.toString(),
      original: {
        openai: openaiResult,
        assemblyAI: assemblyAIResult
      }
    };
  }
}

/**
 * 文章を話者のタイミングに基づいて分割する
 * @param {Array} sentences - 文章の配列
 * @param {Array} utterances - 発話情報の配列
 * @return {Array} - 話者ごとの文章セグメント
 */
function splitTextBySpeakerTiming(sentences, utterances) {
  const result = [];

  // 文章数が少ない場合
  if (sentences.length <= 1) {
    // 最初の話者を取得
    const firstSpeaker = utterances.length > 0 ? utterances[0].speaker : null;

    result.push({
      speaker: firstSpeaker,
      text: sentences.join(' ')
    });

    return result;
  }

  // 発話数が文章数より多い場合、単純に発話ごとに分割
  if (utterances.length >= sentences.length) {
    let currentSpeaker = null;
    let currentText = '';

    for (let i = 0; i < utterances.length; i++) {
      const u = utterances[i];

      // 話者が変わった場合
      if (u.speaker !== currentSpeaker && currentText) {
        result.push({
          speaker: currentSpeaker,
          text: currentText.trim()
        });
        currentText = '';
      }

      currentSpeaker = u.speaker;
      currentText += (u.text || '') + ' ';
    }

    // 最後の話者のテキストを追加
    if (currentText) {
      result.push({
        speaker: currentSpeaker,
        text: currentText.trim()
      });
    }

    return result;
  }

  // 文章数が発話数より多い場合、文章を話者に割り当てる
  const totalUtteranceTime = utterances[utterances.length - 1].end - utterances[0].start;
  const avgSentenceTime = totalUtteranceTime / sentences.length;

  let currentSpeaker = null;
  let currentText = '';
  let sentenceIndex = 0;

  for (let i = 0; i < utterances.length; i++) {
    const u = utterances[i];
    const utteranceDuration = u.end - u.start;
    const sentencesInUtterance = Math.round(utteranceDuration / avgSentenceTime);

    // 話者が変わった場合に結果を追加
    if (u.speaker !== currentSpeaker && currentText) {
      result.push({
        speaker: currentSpeaker,
        text: currentText.trim()
      });
      currentText = '';
    }

    currentSpeaker = u.speaker;

    // この発話に含まれる文章を追加
    for (let j = 0; j < sentencesInUtterance && sentenceIndex < sentences.length; j++) {
      currentText += sentences[sentenceIndex] + ' ';
      sentenceIndex++;
    }
  }

  // 残りの文章を最後の話者に割り当て
  while (sentenceIndex < sentences.length) {
    currentText += sentences[sentenceIndex] + ' ';
    sentenceIndex++;
  }

  // 最後の話者のテキストを追加
  if (currentText) {
    result.push({
      speaker: currentSpeaker,
      text: currentText.trim()
    });
  }

  return result;
}

// AssemblyAIへの別の形式でのリクエスト送信を試す関数（バックアップ）
function tryAlternativeAssemblyAIRequest(file, apiKey) {
  Logger.log('別の形式でのAssemblyAI接続を試行...');

  try {
    // ファイルのバイナリデータを取得
    const fileBlob = file.getBlob();
    const fileName = file.getName();

    // Assembly AI APIのエンドポイントURL
    const uploadUrl = 'https://api.assemblyai.com/v2/upload';
    const transcriptUrl = 'https://api.assemblyai.com/v2/transcript';

    // ファイルをAssembly AIにアップロード
    const uploadOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/octet-stream'
      },
      payload: fileBlob.getBytes(),
      muteHttpExceptions: true
    };

    const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
    const uploadResponseCode = uploadResponse.getResponseCode();

    if (uploadResponseCode !== 200) {
      throw new Error('ファイルのアップロードに失敗しました。レスポンスコード: ' + uploadResponseCode);
    }

    const uploadResponseJson = JSON.parse(uploadResponse.getContentText());
    const audioUrl = uploadResponseJson.upload_url;

    // 別の形式で試す（最もシンプルな形式）
    const transcriptOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        audio_url: audioUrl,
        speaker_labels: true
      }),
      muteHttpExceptions: true
    };

    const transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);
    const transcriptResponseCode = transcriptResponse.getResponseCode();
    Logger.log('代替リクエスト方式のレスポンスコード: ' + transcriptResponseCode);

    if (transcriptResponseCode !== 200) {
      Logger.log('代替リクエスト方式のエラー: ' + transcriptResponse.getContentText());
      return false;
    }

    return true;
  } catch (error) {
    Logger.log('代替リクエスト方式の例外: ' + error.toString());
    return false;
  }
}

// もう一つの代替形式（v3エンドポイントを試す）
function tryAssemblyAIV3Request(file, apiKey) {
  Logger.log('AssemblyAI v3エンドポイントを試行...');

  try {
    // ファイルのバイナリデータを取得
    const fileBlob = file.getBlob();
    const fileName = file.getName();

    // Assembly AI v3 APIのエンドポイントURL（もし存在すれば）
    const uploadUrl = 'https://api.assemblyai.com/v3/upload';
    const transcriptUrl = 'https://api.assemblyai.com/v3/transcript';

    // ファイルをAssembly AIにアップロード
    const uploadOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/octet-stream'
      },
      payload: fileBlob.getBytes(),
      muteHttpExceptions: true
    };

    const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
    const uploadResponseCode = uploadResponse.getResponseCode();

    if (uploadResponseCode !== 200) {
      Logger.log('V3 APIのファイルアップロードエラー: ' + uploadResponseCode);
      return false;
    }

    const uploadResponseJson = JSON.parse(uploadResponse.getContentText());
    const audioUrl = uploadResponseJson.upload_url;

    // V3形式で試す
    const transcriptOptions = {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        audio_url: audioUrl,
        diarization: true,
        speakers_expected: 2
      }),
      muteHttpExceptions: true
    };

    const transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);
    const transcriptResponseCode = transcriptResponse.getResponseCode();
    Logger.log('V3 API試行のレスポンスコード: ' + transcriptResponseCode);

    if (transcriptResponseCode !== 200) {
      Logger.log('V3 API試行のエラー: ' + transcriptResponse.getContentText());
      return false;
    }

    return true;
  } catch (error) {
    Logger.log('V3 API試行の例外: ' + error.toString());
    return false;
  }
}

/**
 * OpenAI GPT-4o-miniを使用して話者分離マージを洗練する
 * @param {Object} openaiTranscriptResult - GPT-4o Transcribeの結果
 * @param {Object} assemblyAIResult - AssemblyAIの結果
 * @param {string} apiKey - OpenAI APIキー
 * @return {Object} - マージした結果
 */
function enhanceDialogueWithGPT4Mini(openaiTranscriptResult, assemblyAIResult, apiKey) {
  Logger.log('GPT-4o-miniによる会話洗練処理を開始...');

  try {
    // AssemblyAIの話者情報を抽出
    const utterances = assemblyAIResult.utterances || [];

    // 話者分離情報がない場合はシンプルなマージ結果を返す
    if (!utterances || utterances.length === 0) {
      return textBasedMerge(openaiTranscriptResult, assemblyAIResult);
    }

    // 現在のマージ結果を取得
    const currentMergeResult = textBasedMerge(openaiTranscriptResult, assemblyAIResult);

    // 話者ごとのセリフを抽出
    const speakerRoles = identifySpeakerRoles(utterances);
    const speakerLabels = {};

    Object.keys(speakerRoles).forEach(speakerId => {
      speakerLabels[speakerId] = formatSpeakerLabel(speakerId, speakerRoles);
    });

    // GPT-4o-miniに送るプロンプトを作成
    const prompt = {
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content: `あなたはインサイドセールスの電話会話を整理する専門家です。
以下の会話をより読みやすく整理してください：

1. 会話はインサイドセールスの電話交換です。自動音声案内が最初にある場合もあります。
2. 会話の内容や事実を絶対に変更しないでください。
3. 話者の区別を明確にしてください。AssemblyAIによる話者分離情報とGPT-4o Transcribeの文字起こしを組み合わせて最適化します。
4. 各話者を【自動音声】【営業担当】【お客様】などの役割ラベルで示してください。
5. 出力形式は各発話の前に話者ラベルを付け、発話ごとに改行を入れてください。
6. 自然な会話の流れになるよう整理してください。`
        },
        {
          role: "user",
          content: `元の文字起こし文章:
${openaiTranscriptResult.text}

現在の話者分離結果:
${currentMergeResult.text}

AssemblyAIが検出した話者情報:
${JSON.stringify(speakerLabels)}

これらの情報を組み合わせて、会話をより自然で読みやすく整理してください。話者の役割を適切に判断して【自動音声】【営業担当】【お客様】などのラベルを付けてください。内容を変更せず、話者の切り替わりだけを最適化してください。`
        }
      ],
      temperature: 0.2,
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
      Logger.log('GPT-4o-mini API呼び出しエラー: ' + responseCode);
      Logger.log(response.getContentText());
      // エラー時は通常のマージ結果を返す
      return currentMergeResult;
    }

    const responseJson = JSON.parse(response.getContentText());
    const enhancedText = responseJson.choices[0].message.content;

    Logger.log('GPT-4o-miniによる会話洗練処理が完了しました');

    return {
      text: enhancedText,
      original: {
        openai: openaiTranscriptResult,
        assemblyAI: assemblyAIResult,
        basicMerge: currentMergeResult
      }
    };

  } catch (error) {
    Logger.log('GPT-4o-miniによる会話洗練処理中にエラー: ' + error.toString());
    // エラー時は通常のマージ結果を返す
    return textBasedMerge(openaiTranscriptResult, assemblyAIResult);
  }
} 