/**
 * 文字起こしサービスモジュール（Assembly AI APIとOpenAI GPT-4.1対応版）
 */
var TranscriptionService = (function () {
  /**
   * 音声ファイルを文字起こしする（話者分離機能付き）
   * @param {File} file - 音声ファイル
   * @param {string} assemblyAiApiKey - Assembly AI APIキー
   * @param {string} openaiApiKey - OpenAI APIキー（GPT-4.1 mini用）
   * @return {Object} - 文字起こし結果
   */
  function transcribe(file, assemblyAiApiKey, openaiApiKey) {
    if (!file) {
      throw new Error('ファイルが指定されていません');
    }

    if (!assemblyAiApiKey) {
      throw new Error('Assembly AI APIキーが指定されていません');
    }

    try {
      // ファイルのバイナリデータを取得
      var fileBlob = file.getBlob();
      var fileName = file.getName();

      // Assembly AI APIのエンドポイントURL
      var uploadUrl = 'https://api.assemblyai.com/v2/upload';
      var transcriptUrl = 'https://api.assemblyai.com/v2/transcript';

      // ファイルをAssembly AIにアップロード
      var uploadOptions = {
        method: 'post',
        headers: {
          'Authorization': assemblyAiApiKey,
          'Content-Type': 'application/octet-stream'
        },
        payload: fileBlob.getBytes(),
        muteHttpExceptions: true
      };

      var uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
      var uploadResponseCode = uploadResponse.getResponseCode();

      if (uploadResponseCode !== 200) {
        throw new Error('ファイルのアップロードに失敗しました。レスポンスコード: ' + uploadResponseCode);
      }

      var uploadResponseJson = JSON.parse(uploadResponse.getContentText());
      var audioUrl = uploadResponseJson.upload_url;

      // 文字起こしリクエストを送信
      var transcriptOptions = {
        method: 'post',
        headers: {
          'Authorization': assemblyAiApiKey,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          audio_url: audioUrl,
          language_code: 'ja',
          speaker_labels: true, // 話者分離を有効化
          speakers_expected: 2  // 想定される話者数（オプション）
        }),
        muteHttpExceptions: true
      };

      var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);
      var transcriptResponseJson = JSON.parse(transcriptResponse.getContentText());
      var transcriptId = transcriptResponseJson.id;

      // 文字起こし結果のポーリング
      var pollingUrl = transcriptUrl + '/' + transcriptId;
      var pollingOptions = {
        method: 'get',
        headers: {
          'Authorization': assemblyAiApiKey
        },
        muteHttpExceptions: true
      };

      var pollingResponseJson;
      var status = '';
      var pollingInterval = 3000; // 3秒ごとにポーリング
      var maxPollingAttempts = 60; // 最大ポーリング回数（約3分）
      var pollingAttempts = 0;

      while (status !== 'completed' && status !== 'error' && pollingAttempts < maxPollingAttempts) {
        // ポーリング間隔を待機
        Utilities.sleep(pollingInterval);

        var pollingResponse = UrlFetchApp.fetch(pollingUrl, pollingOptions);
        pollingResponseJson = JSON.parse(pollingResponse.getContentText());
        status = pollingResponseJson.status;

        pollingAttempts++;
      }

      if (status !== 'completed') {
        throw new Error('文字起こし処理がタイムアウトまたはエラーが発生しました。ステータス: ' + status);
      }

      // Assembly AIからの文字起こし結果を取得
      var text = pollingResponseJson.text || '';
      var utterances = pollingResponseJson.utterances || [];

      // 会話から話者情報を抽出
      var speakerInfo = InformationExtractor.extractSpeakerInfoFromConversation(utterances);
      var speakerRoles = InformationExtractor.identifySpeakerRoles(utterances);

      // 話者情報を含む形式でテキストを整形
      var formattedText = '';
      for (var i = 0; i < utterances.length; i++) {
        var utterance = utterances[i];
        var speaker = utterance.speaker;
        var speakerDisplay = formatSpeakerInfo(speaker, speakerInfo, speakerRoles);

        formattedText += speakerDisplay + ' ' + utterance.text + '\n\n';
      }

      // OpenAI GPT-4.1で文章を自然化（APIキーが提供されている場合）
      var enhancedText = formattedText;
      if (openaiApiKey) {
        enhancedText = enhanceTranscriptionWithOpenAI(formattedText, openaiApiKey);
      }

      return {
        text: enhancedText,
        rawText: formattedText, // 元の整形テキストも保持
        utterances: utterances,
        speakerInfo: speakerInfo, // 抽出した話者情報
        speakerRoles: speakerRoles, // 識別した話者の役割
        fileName: fileName
      };
    } catch (error) {
      throw new Error('文字起こし処理中にエラー: ' + error.toString());
    }
  }

  /**
   * 話者IDを会社名と担当者名に変換する
   * @param {string} speakerId - 話者ID（例: "A"）
   * @param {Object} speakerInfo - 抽出した話者情報
   * @param {Object} speakerRoles - 話者の役割情報
   * @return {string} - フォーマットされた話者情報
   */
  function formatSpeakerInfo(speakerId, speakerInfo, speakerRoles) {
    // 役割情報を取得
    var role = (speakerRoles && speakerRoles[speakerId]) ? speakerRoles[speakerId] : null;
    var roleLabel = role === 'sales' ? '営業担当者' : (role === 'customer' ? 'お客様' : 'Speaker ' + speakerId);

    // 抽出した話者情報がある場合はそれを使用
    if (speakerInfo && speakerInfo[speakerId]) {
      var info = speakerInfo[speakerId];
      var displayName = '';

      if (info.company && info.name) {
        // 会社名と担当者名の両方がある場合
        displayName = '【' + roleLabel + ': ' + info.company + ' ' + info.name + '】';
      } else if (info.company) {
        // 会社名のみある場合
        displayName = '【' + roleLabel + ': ' + info.company + '】';
      } else if (info.name) {
        // 担当者名のみある場合
        displayName = '【' + roleLabel + ': ' + info.name + '】';
      }

      if (displayName) {
        return displayName;
      }
    }

    // 役割情報のみで表示
    return '【' + roleLabel + '】';
  }

  /**
   * 文字起こし結果をOpenAI GPT-4.1 miniで自然な文章に整形する
   * @param {string} rawTranscription - Assembly AIからの生の文字起こし結果
   * @param {string} openaiApiKey - OpenAI APIキー
   * @return {string} - 整形された文字起こし結果
   */
  function enhanceTranscriptionWithOpenAI(rawTranscription, openaiApiKey) {
    if (!rawTranscription || !openaiApiKey) {
      return rawTranscription;
    }

    // 短すぎるテキストの場合は早期リターン（ハルシネーション防止）
    if (rawTranscription.length < 50) {
      Logger.log('テキストが短すぎるため自然化処理をスキップします: ' + rawTranscription.substring(0, 30) + '...');
      return rawTranscription;
    }

    try {
      // OpenAI APIのエンドポイントURL
      var url = 'https://api.openai.com/v1/chat/completions';

      // システムプロンプトの作成（Plan-and-Solve形式）
      var systemPrompt = "あなたは営業電話の文字起こしを整形するプロフェッショナルです。Plan-and-Solve形式で処理を行ってください。\n\n" +
        "ステップ1 (Plan)：\n" +
        "- 入力された文字起こしテキストを分析する\n" +
        "- 話者の役割を特定する（営業担当者かお客様か）\n" +
        "- 文法的・表現的な違和感がある部分を特定する\n\n" +
        "ステップ2 (Solve)：\n" +
        "- 内容を変えずに文法的に自然な日本語に整形する\n" +
        "- 話者ラベル（【営業担当者】や【お客様】など）を必ず保持する\n" +
        "- 会話の流れを自然に保つ\n\n" +
        "絶対に守るべき原則：\n" +
        "1. 事実や内容は一切変更しない（創作・追加・推測は禁止）\n" +
        "2. 話者の区別（【営業担当者】や【お客様】など）を保持する\n" +
        "3. 不明瞭な部分は推測せず、そのままにする\n" +
        "4. 短い会話は無理に補完せず、そのまま整形する\n\n" +
        "話者役割の判別基準：\n" +
        "- 営業担当者：「弊社は〜」「ご案内させていただきました」「ご提供しております」など提案・紹介をする側\n" +
        "- お客様：「検討します」「必要ありません」「確認します」など返答や判断をする側";

      // ユーザープロンプト（重複を排除）
      var prompt = "以下の営業電話の文字起こし結果を自然な日本語に整えてください：\n\n" + rawTranscription;

      // 文字起こしの長さによってチャンクサイズとトークン数を調整
      var maxTokens = Math.min(4096, rawTranscription.length * 1.5);
      if (maxTokens < 1024) maxTokens = 1024;

      // APIリクエストオプション
      var options = {
        method: 'post',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + openaiApiKey
        },
        payload: JSON.stringify({
          model: "gpt-4.1",
          messages: [
            {
              role: "system",
              content: systemPrompt
            },
            {
              role: "user",
              content: prompt
            }
          ],
          temperature: 0.0, // ハルシネーション防止のため温度を0に設定
          max_tokens: maxTokens
        }),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        Logger.log('OpenAI APIエラー: ' + responseCode + ', ' + response.getContentText());
        return rawTranscription; // 失敗した場合は元の文字起こし結果を返す
      }

      // レスポンスをJSONとしてパース
      var responseJson = JSON.parse(response.getContentText());

      // OpenAI APIからのレスポンス構造に合わせて処理
      if (!responseJson.choices || responseJson.choices.length === 0) {
        return rawTranscription;
      }

      var enhancedText = responseJson.choices[0].message.content;

      // 結果の検証（ハルシネーション検出）
      if (enhancedText.length > rawTranscription.length * 1.5) {
        // 結果が元のテキストに比べて極端に長くなった場合は疑わしい
        Logger.log('自然化結果が過剰に長いため、元のテキストを返します。元:' + rawTranscription.length + '文字, 結果:' + enhancedText.length + '文字');
        return rawTranscription;
      }

      // 話者ラベルが全て保持されているか確認（最低限のチェック）
      var originalSpeakerCount = (rawTranscription.match(/【.*?】/g) || []).length;
      var enhancedSpeakerCount = (enhancedText.match(/【.*?】/g) || []).length;

      if (originalSpeakerCount > 0 && enhancedSpeakerCount < originalSpeakerCount * 0.8) {
        // 話者ラベルが20%以上減少している場合は疑わしい
        Logger.log('自然化結果で話者ラベルが大幅に減少しています。元:' + originalSpeakerCount + '個, 結果:' + enhancedSpeakerCount + '個');
        return rawTranscription;
      }

      return enhancedText;
    } catch (error) {
      Logger.log('文字起こし自然化中にエラー: ' + error.toString());
      return rawTranscription; // エラーが発生した場合は元の文字起こし結果を返す
    }
  }

  // 公開メソッド
  return {
    transcribe: transcribe,
    enhanceTranscriptionWithOpenAI: enhanceTranscriptionWithOpenAI,
    formatSpeakerInfo: formatSpeakerInfo
  };
})();