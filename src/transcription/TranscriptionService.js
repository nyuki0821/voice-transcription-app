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
      var maxPollingAttempts = 100; // 最大ポーリング回数（約5分）- 延長
      var pollingAttempts = 0;

      while (status !== 'completed' && status !== 'error' && pollingAttempts < maxPollingAttempts) {
        // ポーリング間隔を待機
        Utilities.sleep(pollingInterval);

        var pollingResponse = UrlFetchApp.fetch(pollingUrl, pollingOptions);
        pollingResponseJson = JSON.parse(pollingResponse.getContentText());
        status = pollingResponseJson.status;

        pollingAttempts++;

        // 進捗状況をログに記録（10回ごと）
        if (pollingAttempts % 10 === 0) {
          Logger.log('文字起こし進捗: ステータス=' + status + ', 試行回数=' + pollingAttempts + '/' + maxPollingAttempts);
        }
      }

      // 処理中または完了でない場合でも部分的な結果がある場合は使用する
      if (status !== 'completed') {
        Logger.log('文字起こし処理が完全には完了していません。ステータス: ' + status + 'で続行を試みます');

        // 文字起こし結果のURLを保存して後で確認できるようにする
        Logger.log('後で結果を確認するためのURL: ' + pollingUrl);

        // processingステータスでも部分的な結果があればそれを使用
        if (status === 'processing' && pollingResponseJson.text) {
          Logger.log('部分的な文字起こし結果を使用します（処理中）');
          // 警告メッセージを追加
          var warningText = "【注意: この文字起こしは処理が完全に完了していません。精度が低い可能性があります】\n\n";
          pollingResponseJson.text = warningText + (pollingResponseJson.text || '');
        } else {
          // 完全にエラーの場合は例外をスロー
          throw new Error('文字起こし処理がタイムアウトまたはエラーが発生しました。ステータス: ' + status);
        }
      }

      // Assembly AIからの文字起こし結果を取得
      var text = pollingResponseJson.text || '';
      if (!text) {
        throw new Error('テキストが指定されていないか空です');
      }

      var utterances = pollingResponseJson.utterances || [];

      // 会話から話者情報を抽出
      var speakerInfo = InformationExtractor.extractSpeakerInfoFromConversation(utterances);
      var speakerRoles = InformationExtractor.identifySpeakerRoles(utterances);

      // 話者情報を含む形式でテキストを整形
      var formattedText = '';

      // utterancesがある場合は話者区分ありで整形
      if (utterances && utterances.length > 0) {
        for (var i = 0; i < utterances.length; i++) {
          var utterance = utterances[i];
          var speaker = utterance.speaker;
          var speakerDisplay = formatSpeakerInfo(speaker, speakerInfo, speakerRoles);

          formattedText += speakerDisplay + ' ' + utterance.text + '\n\n';
        }
      } else {
        // 話者区分がない場合は単純に文字起こし結果をそのまま使用
        formattedText = text;
      }

      // OpenAI GPT-4.1で文章を自然化（APIキーが提供されている場合）
      var enhancedText = formattedText;
      if (openaiApiKey && formattedText.length > 50) {
        try {
          enhancedText = enhanceTranscriptionWithOpenAI(formattedText, openaiApiKey);
        } catch (enhanceError) {
          Logger.log('文章の自然化処理中にエラーが発生しました: ' + enhanceError.toString());
          // エラーが発生しても元のフォーマット済みテキストを使用して処理を続行
        }
      }

      // 返却前のポスト処理
      enhancedText = postProcessTranscription(enhancedText);

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

      // システムプロンプトの改善版（自然さ優先）
      var systemPrompt = "あなたは営業電話の文字起こしを整形する専門家です。典型的な営業代行の通話では、次のような特徴があります：\n\n" +
        "1. 営業担当者（電話をかける側）の特徴：\n" +
        "- 「お忙しいところ恐れ入ります」「ご案内でご連絡しました」などの丁寧な表現\n" +
        "- 自社サービスの紹介（「弊社は～」「ご提供しております」など）\n" +
        "- 比較的長めの発言（サービス説明など）\n" +
        "- 「おめでとうございます」など褒める表現\n" +
        "- 自己紹介パターン：「私、○○の△△と申します」「○○の者です」など\n" +
        "- 提案・案内フレーズ：「ご案内でご連絡」「ご紹介させていただく」など\n" +
        "- 特徴的なフレーズ：「失礼いたします」「お時間よろしいでしょうか」など\n\n" +
        "2. お客様（電話を受ける側）の特徴：\n" +
        "- 「確認します」「少々お待ちください」などの対応フレーズ\n" +
        "- 「必要ありません」「結構です」などの断りフレーズ\n" +
        "- 「社内で対応しています」「関連会社があります」など社内情報の言及\n" +
        "- 比較的短い返答が多い\n" +
        "- 自己紹介パターン：「こちら担当の△△です」「○○を担当しております」など\n" +
        "- 特徴的な反応：「今は考えていません」「予算が厳しい」「他社と契約中」など\n\n" +
        "営業代行サービスでは多様な会社名が営業側として登場します。特定の会社名に依存せず、会話の特徴と話し方のパターンから話者の役割を正確に識別してください。\n\n" +
        "【会話の自然化ルール】\n" +
        "1. 営業担当者の発言\n" +
        "   - 丁寧な敬語を基本とするが、過度な敬語は避ける\n" +
        "   - 「〜させていただく」は必要最小限に\n" +
        "   - 相槌は「はい」「ええ」など自然な形に\n" +
        "   - 説明は簡潔に、要点を押さえて\n\n" +
        "2. お客様の発言\n" +
        "   - 敬語は控えめに、自然な会話調に\n" +
        "   - 相槌は「はい」「ええ」など自然な形に\n" +
        "   - 断りの表現は柔らかく、丁寧に\n\n" +
        "3. 会話の流れ\n" +
        "   - 不自然な間は適度に補完\n" +
        "   - 唐突な話題転換は避け、自然な流れに\n" +
        "   - 重複する表現は整理\n" +
        "   - 文末は「です」「ます」を適度に使用";

      // ユーザープロンプト（自然さ重視）
      var prompt = "以下は営業電話の文字起こし結果です。文章を自然な日本語に整えてください。以下の点に特に注意してください：\n\n" +
        "1. 事実や内容は一切変更しないこと\n" +
        "2. 話者の区別（【営業担当者】や【お客様】など）は必ず保持すること\n" +
        "3. 話者役割の判別基準：\n" +
        "   - 営業担当者：「弊社は〜」「ご案内させていただきました」「ご提供しております」など提案・紹介をする側\n" +
        "   - お客様：「検討します」「必要ありません」「確認します」など返答や判断をする側\n" +
        "4. 省略されている主語や述語を補完し、読みやすくすること\n" +
        "5. 各発言の前後関係を考慮し、会話の流れを自然にすること\n" +
        "6. 自己紹介や会社名の言及は正確に保持すること\n" +
        "7. 敬語表現は文脈に応じて適切に調整すること\n" +
        "8. 相槌や間投詞は自然な形に修正すること\n\n" +
        "重要：これは営業代行の会話であり、様々な会社が営業側として登場します。会社名に関わらず、話者の役割を話し方のパターンから正しく識別してください。\n\n" +
        rawTranscription;

      // 文字起こしの長さによってチャンクサイズとトークン数を調整
      var maxTokens = Math.min(4096, rawTranscription.length * 1.5);
      if (maxTokens < 1024) maxTokens = 1024;

      // APIリクエストオプション - temperatureを0.3に設定
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
          temperature: 0.3,
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

      // 結果の検証（改善版）
      if (enhancedText.includes('ステップ1') || enhancedText.includes('Plan') || enhancedText.includes('Solve')) {
        Logger.log('思考プロセスが含まれているため、除去を試みます');
        // 「---」の後の部分だけを抽出する試み
        var parts = enhancedText.split('---');
        if (parts.length > 1) {
          enhancedText = parts[parts.length - 1].trim();
        } else {
          // 思考プロセスの後に続くテキストを抽出する別の試み
          var match = enhancedText.match(/(【営業担当者】[\s\S]*)/);
          if (match) {
            enhancedText = match[1].trim();
          } else {
            // それでも抽出できない場合は元のテキストを返す
            Logger.log('思考プロセスの除去に失敗しました。元のテキストを返します');
            return rawTranscription;
          }
        }
      }

      // 基本的な検証（極端に長い場合など）
      if (enhancedText.length > rawTranscription.length * 2) {
        Logger.log('整形結果が極端に長くなりました。元のテキストを返します');
        return rawTranscription;
      }

      // 最終チェック - 会話が含まれているか
      if (!enhancedText.includes('【営業担当者】') || !enhancedText.includes('【お客様】')) {
        Logger.log('話者ラベルが失われています。元のテキストを返します');
        return rawTranscription;
      }

      return enhancedText;
    } catch (error) {
      Logger.log('文字起こし自然化中にエラー: ' + error.toString());
      return rawTranscription; // エラーが発生した場合は元の文字起こし結果を返す
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

      // 1. 連続した繰り返し表現を削除（「お世話になっております。お世話になっております。」など）
      var repeatPattern = /(.{5,20})\s*\1/g;
      while (repeatPattern.test(content)) {
        content = content.replace(repeatPattern, "$1");
      }

      // 2. 不自然な文字列を適切な表現に置き換え（パターンマッチング）
      // 相槌や間投詞の修正
      content = content.replace(/お\s*だし\s*[ょ]\s*[ーう]/gi, "かしこまりました");
      content = content.replace(/おだし[ょ][ーう]/gi, "かしこまりました");
      content = content.replace(/だし[ょ][ーう]/g, "でしょう");
      content = content.replace(/お\s*だし/g, "はい");
      content = content.replace(/おだし/g, "はい");
      content = content.replace(/は\s*い\s*は\s*い/g, "はい");
      content = content.replace(/え\s*え\s*え\s*え/g, "ええ");
      content = content.replace(/あ\s*あ\s*あ\s*あ/g, "ああ");

      // 3. 不自然な数字の間の空白を削除
      content = content.replace(/([A-Za-z])\s+(\d)/g, "$1$2");

      // 4. 敬語表現の自然化
      if (content.split("お世話になっております").length > 2) {
        content = content.replace(/お世話になっております。?\s*/g, "");
        content = "お世話になっております。 " + content.trim();
      }

      // 5. 不自然な単語分割を修正
      content = content.replace(/([一-龯ぁ-ゔァ-ヴー])\s+([一-龯ぁ-ゔァ-ヴー])/g, "$1$2");

      // 6. 会話の自然化
      // 文末の「です」「ます」の連続を整理
      content = content.replace(/(です|ます)。\s*(です|ます)。/g, "$1。");

      // 不自然な「です」「ます」の連続を修正
      content = content.replace(/(です|ます)です/g, "$1");
      content = content.replace(/(です|ます)ます/g, "$1");

      // 7. 句読点の整理
      content = content.replace(/。\s*。/g, "。");
      content = content.replace(/、\s*、/g, "、");
      content = content.replace(/。\s*、/g, "。");
      content = content.replace(/、\s*。/g, "。");

      // 8. 余分な空白の削除
      content = content.replace(/\s+/g, " ").trim();

      // 9. 敬語表現の自然化（追加）
      // 「させていただく」の過剰使用を修正
      content = content.replace(/させていただいております/g, "しております");
      content = content.replace(/させていただきます/g, "します");
      content = content.replace(/させていただきました/g, "しました");

      // 10. 会話の流れの改善（追加）
      // 不自然な間投詞を修正
      content = content.replace(/ねそうな/g, "そうですね");
      content = content.replace(/でそうです/g, "そうですね");
      content = content.replace(/ね分かりました/g, "分かりました");
      content = content.replace(/ねありがとうございます/g, "ありがとうございます");

      // 11. 断り表現の自然化（追加）
      content = content.replace(/必要で\s*はないん\s*ですよ/g, "必要ないんですよ");
      content = content.replace(/必要で\s*はないの\s*ですよ/g, "必要ないんですよ");

      // 12. 相槌の自然化（追加）
      content = content.replace(/ありがとうございます(?!。)/g, "ありがとうございます。");
      content = content.replace(/分かりました(?!。)/g, "分かりました。");

      // 話者ラベルと処理済みコンテンツを結合
      processed.push(speakerLabel + " " + content);
    }

    return processed.join("\n\n");
  }

  // 公開メソッド
  return {
    transcribe: transcribe,
    enhanceTranscriptionWithOpenAI: enhanceTranscriptionWithOpenAI,
    formatSpeakerInfo: formatSpeakerInfo,
    postProcessTranscription: postProcessTranscription
  };
})();