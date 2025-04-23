/**
 * 文字起こしサービスモジュール（Assembly AI APIとOpenAI GPT-4.1対応版）
 */
var TranscriptionService = (function() {
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
    
    try {
      // OpenAI APIのエンドポイントURL
      var url = 'https://api.openai.com/v1/chat/completions';
      
      // プロンプトの作成
      var prompt = "以下は営業電話の文字起こし結果です。文章を自然な日本語に整えてください。以下の点に特に注意してください：\n\n" +
                 "1. 事実や内容は一切変更しないこと\n" +
                 "2. 話者の区別（【営業担当者】や【お客様】など）は必ず保持すること\n" +
                 "3. 話者役割の判別基準：\n" +
                 "   - 営業担当者：「弊社は〜」「ご案内させていただきました」「ご提供しております」など提案・紹介をする側\n" +
                 "   - お客様：「検討します」「必要ありません」「確認します」など返答や判断をする側\n" +
                 "4. 省略されている主語や述語を補完し、読みやすくすること\n" +
                 "5. 各発言の前後関係を考慮し、会話の流れを自然にすること\n\n" +
                 "重要：これは営業代行の会話であり、様々な会社が営業側として登場します。会社名に関わらず、話者の役割を話し方のパターンから正しく識別してください。\n\n" +
                 rawTranscription;
      
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
              content: "あなたは営業電話の文字起こしを整形する専門家です。典型的な営業代行の通話では、次のような特徴があります：\n\n" +
                      "1. 営業担当者（電話をかける側）の特徴：\n" +
                      "- 「お忙しいところ恐れ入ります」「ご案内でご連絡しました」などの丁寧な表現\n" +
                      "- 自社サービスの紹介（「弊社は～」「ご提供しております」など）\n" +
                      "- 比較的長めの発言（サービス説明など）\n" +
                      "- 「おめでとうございます」など褒める表現\n\n" +
                      "2. お客様（電話を受ける側）の特徴：\n" +
                      "- 「確認します」「少々お待ちください」などの対応フレーズ\n" +
                      "- 「必要ありません」「結構です」などの断りフレーズ\n" +
                      "- 「社内で対応しています」「関連会社があります」など社内情報の言及\n" +
                      "- 比較的短い返答が多い\n\n" +
                      "営業代行サービスでは多様な会社名が営業側として登場します。特定の会社名に依存せず、会話の特徴と話し方のパターンから話者の役割を正確に識別してください。"
            },
            {
              role: "user",
              content: prompt
            }
          ],
          temperature: 0.3,
          max_tokens: 2048
        }),
        muteHttpExceptions: true
      };
      
      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();
      
      if (responseCode !== 200) {
        return rawTranscription; // 失敗した場合は元の文字起こし結果を返す
      }
      
      // レスポンスをJSONとしてパース
      var responseJson = JSON.parse(response.getContentText());
      
      // OpenAI APIからのレスポンス構造に合わせて処理
      if (!responseJson.choices || responseJson.choices.length === 0) {
        return rawTranscription;
      }
      
      var enhancedText = responseJson.choices[0].message.content;
      
      return enhancedText;
    } catch (error) {
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