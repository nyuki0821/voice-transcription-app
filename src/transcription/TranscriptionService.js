/**
 * 文字起こしサービスモジュール（Assembly AI APIとOpenAI GPT-4.1 mini対応版）
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
      Logger.log('文字起こし処理開始: ' + file.getName());
      // GPT-4o-mini-transcribeでの文字起こし
      const openaiTranscriptResult = transcribeWithGPT4oMini(file, openaiApiKey);
      Logger.log('GPT-4o-mini 文字起こし完了');

      try {
        // AssemblyAIでの話者分離処理
        const assemblyAIResult = transcribeWithAssemblyAI(file, assemblyAiApiKey);
        Logger.log('AssemblyAI 話者分離処理完了');

        // 結果のマージ
        const enhancedResult = enhanceDialogueWithGPT4Mini(openaiTranscriptResult, assemblyAIResult, openaiApiKey);
        Logger.log('GPT-4.1 miniによる会話洗練処理完了');

        return {
          text: enhancedResult.text,
          rawText: openaiTranscriptResult.text,
          utterances: assemblyAIResult.utterances || [],
          speakerInfo: InformationExtractor.extractSpeakerInfoFromConversation(assemblyAIResult.utterances || []),
          speakerRoles: InformationExtractor.identifySpeakerRoles(assemblyAIResult.utterances || []),
          fileName: file.getName()
        };
      } catch (assemblyError) {
        Logger.log('AssemblyAI処理中にエラー: ' + assemblyError.toString());

        // AssemblyAIでエラーが発生した場合、別の形式で試す
        Logger.log('別の形式でリクエストを試みます...');

        // 別の形式でリクエストを試す
        const altSuccess = tryAlternativeAssemblyAIRequest(file, assemblyAiApiKey);

        if (altSuccess) {
          Logger.log('代替リクエスト方式が成功しました。デバッグ用のログを確認してください。');
        } else {
          Logger.log('代替リクエスト方式も失敗しました。さらにV3 APIを試します...');

          // V3 APIも試す
          const v3Success = tryAssemblyAIV3Request(file, assemblyAiApiKey);

          if (v3Success) {
            Logger.log('V3 API方式が成功しました。デバッグ用のログを確認してください。');
          } else {
            Logger.log('すべての試行が失敗しました。GPT-4o-miniの結果のみを返します。');
          }
        }

        // GPT-4o-miniの結果だけで簡易フォーマット
        const simpleResult = {
          text: postProcessTranscription(openaiTranscriptResult.text),
          error: assemblyError.toString(),
          rawText: openaiTranscriptResult.text,
          utterances: [],
          speakerInfo: {},
          speakerRoles: {},
          fileName: file.getName()
        };

        return simpleResult;
      }
    } catch (error) {
      Logger.log('文字起こし処理中にエラー: ' + error.toString());
      throw error;
    }
  }

  /**
   * GPT-4o-mini Transcribeで文字起こしを行う
   * @param {File} file - 音声ファイル
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - 文字起こし結果
   */
  function transcribeWithGPT4oMini(file, apiKey) {
    Logger.log('GPT-4o-mini-transcribeでの文字起こし開始...');

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

      // word-levelタイムスタンプも追加（両方指定して確実に取得）
      let wordTimestampsPart = '--' + boundary + '\r\n';
      wordTimestampsPart += 'Content-Disposition: form-data; name="timestamp_granularities[]"\r\n\r\n';
      wordTimestampsPart += 'word' + '\r\n';

      // 最後の境界を追加
      let endPart = '--' + boundary + '--\r\n';

      // パート結合のためのバイト配列を作成
      const head = Utilities.newBlob(data).getBytes();
      const modelPartBytes = Utilities.newBlob(modelPart).getBytes();
      const langPartBytes = Utilities.newBlob(langPart).getBytes();
      const responsePartBytes = Utilities.newBlob(responsePart).getBytes();
      const timestampsPartBytes = Utilities.newBlob(timestampsPart).getBytes();
      const wordTimestampsPartBytes = Utilities.newBlob(wordTimestampsPart).getBytes();
      const tail = Utilities.newBlob(endPart).getBytes();

      // 全てのバイト配列を結合
      const payload = [].concat(
        head,
        fileBytes,
        modelPartBytes,
        langPartBytes,
        responsePartBytes,
        timestampsPartBytes,
        wordTimestampsPartBytes,
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
      Logger.log('GPT-4o-mini APIレスポンスコード: ' + responseCode);

      if (responseCode !== 200) {
        Logger.log('エラーレスポンス: ' + response.getContentText());
        throw new Error('GPT-4o-mini API呼び出しエラー: ' + responseCode);
      }

      // レスポンスをJSONとしてパース
      const responseJson = JSON.parse(response.getContentText());

      // シンプルに文字起こし結果を返す
      return {
        text: responseJson.text,
        language: responseJson.language,
        duration: responseJson.duration,
        fileName: fileName
      };

    } catch (error) {
      Logger.log('GPT-4o-mini Transcribe処理中にエラー: ' + error.toString());
      throw new Error('GPT-4o-mini Transcribe処理中にエラー: ' + error.toString());
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

      // 言語に応じた最適なオプションを設定
      let transcriptPayload = {
        audio_url: audioUrl,
        language_code: 'ja',
        speaker_labels: true,
        entity_detection: true  // エンティティ検出を有効化
      };

      // 言語に応じた機能の設定
      if (transcriptPayload.language_code === 'ja') {
        // 日本語では以下の機能が利用できないため除外
        // - 感情分析 (sentiment_analysis)
        // - 自動チャプター (auto_chapters)
        // - 自動ハイライト (auto_highlights)

        // SLAM-1モデルは日本語でも対応しているため設定
        transcriptPayload.language_model = 'assemblyai/slam-1';
      } else {
        // 英語など他の言語では高度な機能を有効化
        transcriptPayload.auto_chapters = true;
        transcriptPayload.sentiment_analysis = true;
        transcriptPayload.auto_highlights = true;
        transcriptPayload.language_model = 'assemblyai/slam-1';
      }

      // 文字起こしリクエストを送信
      const transcriptOptions = {
        method: 'post',
        headers: {
          'Authorization': apiKey,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(transcriptPayload),
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
      const transcriptId = transcriptResponseJson.id;

      // トランスクリプションIDの取得確認をログに出力
      Logger.log('AssemblyAI トランスクリプションID取得: ' + transcriptId);

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
      const pollingInterval = 5000; // 5秒ごとにポーリング
      const maxPollingAttempts = 120; // 最大ポーリング回数（約10分）
      let pollingAttempts = 0;

      while (status !== 'completed' && status !== 'error' && pollingAttempts < maxPollingAttempts) {
        // ポーリング間隔を待機
        Utilities.sleep(pollingInterval);

        try {
          const pollingResponse = UrlFetchApp.fetch(pollingUrl, pollingOptions);
          const responseCode = pollingResponse.getResponseCode();

          if (responseCode !== 200) {
            const errorText = pollingResponse.getContentText();
            Logger.log('AssemblyAI ポーリングエラー: ' + errorText);

            if (errorText.includes('not found') && pollingAttempts > 5) {
              throw new Error('AssemblyAI トランスクリプションIDが無効です: ' + transcriptId);
            } else if (pollingAttempts > 10) {
              throw new Error('AssemblyAI ポーリングエラー: レスポンスコード ' + responseCode);
            }

            Utilities.sleep(10000); // エラーの場合は長めに待機
            pollingAttempts++;
            continue;
          }

          pollingResponseJson = JSON.parse(pollingResponse.getContentText());
          status = pollingResponseJson.status;

          // 処理進捗のログ記録（5回ごと）
          if (pollingAttempts % 5 === 0) {
            Logger.log('AssemblyAI文字起こし進捗: ステータス=' + status + ', 試行回数=' + pollingAttempts + '/' + maxPollingAttempts);
          }
        } catch (pollingError) {
          Logger.log('AssemblyAI ポーリング中の例外: ' + pollingError.toString());
          Utilities.sleep(10000); // エラーの場合は長めに待機
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

      // エンティティ検出結果をログに記録
      if (pollingResponseJson.entities) {
        Logger.log('エンティティ検出結果: ' + JSON.stringify(pollingResponseJson.entities));
      }

      // 詳細なエンティティ情報があれば記録
      let entitiesOriginal = [];
      if (pollingResponseJson.results &&
        pollingResponseJson.results.entities &&
        Array.isArray(pollingResponseJson.results.entities)) {
        entitiesOriginal = pollingResponseJson.results.entities;
        Logger.log('詳細なエンティティ情報: ' + JSON.stringify(entitiesOriginal).substring(0, 500) + '...');
      }

      return {
        text: pollingResponseJson.text || '',
        utterances: pollingResponseJson.utterances || [],
        entities: pollingResponseJson.entities || {},  // エンティティ情報を追加
        entities_original: entitiesOriginal,           // 詳細なエンティティ情報を追加
        fileName: fileName
      };

    } catch (error) {
      Logger.log('AssemblyAI処理中にエラー: ' + error.toString());
      throw error;
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
   * OpenAI GPT-4.1 miniを使用して話者分離マージを洗練する
   * @param {Object} openaiTranscriptResult - GPT-4o-mini Transcribeの結果
   * @param {Object} assemblyAIResult - AssemblyAIの結果
   * @param {string} apiKey - OpenAI APIキー
   * @return {Object} - マージした結果
   */
  function enhanceDialogueWithGPT4Mini(openaiTranscriptResult, assemblyAIResult, apiKey) {
    Logger.log('GPT-4.1 miniによる会話洗練処理を開始...');

    try {
      // AssemblyAIの話者情報を抽出
      const utterances = assemblyAIResult.utterances || [];

      // 話者分離情報がない場合はシンプルなマージ結果を返す
      if (!utterances || utterances.length === 0) {
        return textBasedMerge(openaiTranscriptResult, assemblyAIResult);
      }

      // 現在のマージ結果を取得（下準備として実行）
      Logger.log('初期マージデータを準備中（下準備）...');
      const currentMergeResult = textBasedMerge(openaiTranscriptResult, assemblyAIResult);
      Logger.log('初期マージデータの準備完了、GPT-4.1 miniによる高度な処理を開始...');

      // 話者ごとのセリフを抽出
      const speakerRoles = InformationExtractor.identifySpeakerRoles(utterances);
      const speakerInfo = InformationExtractor.extractSpeakerInfoFromConversation(utterances);

      // 現在の話者情報を整形
      const speakerLabels = {};
      Object.keys(speakerRoles).forEach(speakerId => {
        const role = speakerRoles[speakerId];
        const info = speakerInfo[speakerId] || {};

        // 役割ラベル
        const roleLabel = role === 'sales' ? '営業担当者' : (role === 'customer' ? 'お客様' : '話者' + speakerId);

        // 詳細情報
        let detailLabel = '';
        if (info.company && info.name) {
          detailLabel = `${info.company} ${info.name}`;
        } else if (info.company) {
          detailLabel = info.company;
        } else if (info.name) {
          detailLabel = info.name;
        }

        speakerLabels[speakerId] = {
          role: roleLabel,
          detail: detailLabel
        };
      });

      // エンティティ情報をテキスト形式に変換
      let entitiesText = '';
      if (assemblyAIResult.entities && Object.keys(assemblyAIResult.entities).length > 0) {
        entitiesText = 'エンティティ検出情報:\n';

        // エンティティの種類と日本語名のマッピング
        const entityTypeMap = {
          'person_name': '人物名',
          'location': '場所',
          'organization': '組織',
          'date': '日付',
          'time': '時間',
          'money': '金額',
          'percent': 'パーセント',
          'event': 'イベント',
          'product': '製品',
          'phone_number': '電話番号',
          'email_address': 'メールアドレス',
          'url': 'URL',
          'age': '年齢',
          'occupation': '職業',
          'date_interval': '期間'
        };

        // タイプごとのエンティティカウントを表示
        for (const [type, count] of Object.entries(assemblyAIResult.entities)) {
          const typeJa = entityTypeMap[type] || type;
          entitiesText += `${typeJa}: ${count}件\n`;
        }

        // 詳細なエンティティ情報があれば追加
        if (Array.isArray(assemblyAIResult.entities_original)) {
          const entityDetails = {};

          // エンティティをタイプごとにグループ化
          assemblyAIResult.entities_original.forEach(entity => {
            const type = entity.entity_type;
            const typeJa = entityTypeMap[type] || type;

            if (!entityDetails[typeJa]) {
              entityDetails[typeJa] = new Set();
            }
            entityDetails[typeJa].add(entity.text);
          });

          // グループ化したエンティティを表示
          entitiesText += '\n詳細なエンティティ情報:\n';
          for (const [typeJa, values] of Object.entries(entityDetails)) {
            entitiesText += `${typeJa}: ${Array.from(values).join('、')}\n`;
          }
        }
      }

      // GPT-4.1 miniに送るプロンプトを作成
      const prompt = {
        model: "gpt-4.1-mini",
        messages: [
          {
            role: "system",
            content: `あなたはインサイドセールスの電話会話を分析・整理する専門家です。以下のステップで会話を最適化してください：

【思考ステップ】
1. 提供された情報（元の文字起こし、話者分離結果、検出された話者情報）を比較分析する
2. 会話の流れを把握し、どの発言がどの話者に属するか判断する
3. 自動音声、営業担当者、顧客など、話者の役割を特定する
4. 同じ話者の連続した発言を適切にまとめるべきか判断する
5. 不自然な文の区切りや誤った話者の切り替わりを特定する

【行動ステップ】
1. 各発言に適切な話者ラベルを付ける【会社・団体名 部署名 担当者名】または【会社・団体名 担当者名】の形式で
2. 同じ話者の意味が繋がる発言は適切にまとめる
3. 文脈から明らかに話者が切り替わる箇所では、話者を正しく分離する
4. 自然な会話の流れになるよう発言をグループ化する
5. 最後に、整理された会話全体を出力

【営業会社の候補リスト】
以下は営業会社の候補リストです。話者が営業側と判断される場合は、このリストから適切な会社名を選んでください：
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
・株式会社インフィニットマインド

【重要ルール】
- 話者ラベルは【会社・団体名 部署名 担当者名】または【会社・団体名 担当者名】の形式で記載する
- 話者の中で営業側と顧客側を明確に区別する
- 営業側は上記リストの会社から選び、顧客側は会話から判断して適切な会社・団体名を記載する
- 会社名と担当者名は会話内容から抽出すること
- AssemblyAIで文字起こしした結果が正しいとは限らないことに注意
- 飛躍した推論は避ける必要があるものの、論理的に導かれる話者の分離や文章の整形は行なって下さい
- 各発話の前に話者ラベルを付け、発話ごとに適切に改行してください
- 最終的な出力は読みやすく、自然な対話形式にしてください`
          },
          {
            role: "user",
            content: `元の文字起こし文章:
${openaiTranscriptResult.text}

現在の話者分離結果:
${currentMergeResult.text}

AssemblyAIが検出した話者情報:
${JSON.stringify(speakerLabels, null, 2)}

${entitiesText}

これらの情報を比較分析して、以下のプロセスで会話を整理してください：

1. まず、各発言者の特定：話者の特徴（言葉遣い、自己紹介、発言内容）から誰が話しているか判断してください
2. 次に、営業側と顧客側を区別：営業側は営業会社候補リストから選び、顧客側は会話から判断してください
3. 検出されたエンティティ情報（特に人物名、組織名、役職など）を活用して、話者の正確な所属や役割を特定してください
4. 連続した発言のグループ化：同じ話者の発言を自然にまとめてください
5. 最後に、整理された会話全体を出力してください (飛躍した推論は避ける必要があるものの、論理的に導かれる話者の分離や文章の整形は行なって下さい)

話者ラベルは【会社・団体名 部署名 担当者名】または【会社・団体名 担当者名】の形式で記載してください。
営業側は必ず営業会社候補リストから選んでください。顧客側は会話内容から最適な表記を使用してください。

例：
【株式会社ENERALL 佐藤】 （営業側の例）
【東京商事 総務企画課 田中様】 （顧客側の例）

発言が切れ目なく繋がるよう注意し、会話の流れを崩さないように整理してください。
最終的な出力は自然な対話形式で、各話者の発言が明確に区別できるようにしてください。

会話の最後に、検出されたエンティティ情報を要約して含めてください。主な人物、組織、日付、時間などを簡潔にまとめてください。`
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
        // エラー時は現在のマージ結果を返す
        return currentMergeResult;
      }

      const responseJson = JSON.parse(response.getContentText());
      const enhancedText = responseJson.choices[0].message.content;

      Logger.log('GPT-4.1 miniによる会話洗練処理が完了しました（最終結果として採用）');

      // 最終的なポスト処理を適用
      const finalText = postProcessTranscription(enhancedText);

      return {
        text: finalText,
        original: {
          openai: openaiTranscriptResult,
          assemblyAI: assemblyAIResult,
          basicMerge: currentMergeResult
        }
      };

    } catch (error) {
      Logger.log('GPT-4.1 miniによる会話洗練処理中にエラー: ' + error.toString());
      // エラー時は単純にテキストを整形して返す
      return {
        text: postProcessTranscription(openaiTranscriptResult.text),
        error: error.toString(),
        original: {
          openai: openaiTranscriptResult,
          assemblyAI: assemblyAIResult || { error: 'Processing failed' }
        }
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
        // 話者ラベルがない場合は、単純にテキストを段落ごとに分割
        // 文章ごとに分割
        const sentences = paragraph.split(/(?<=[。．！？])\s*/);
        // 3-4文ごとに段落を分ける
        let currentParagraph = '';
        for (let j = 0; j < sentences.length; j++) {
          currentParagraph += sentences[j];
          if ((j + 1) % 4 === 0 || j === sentences.length - 1) {
            processed.push(currentParagraph.trim());
            currentParagraph = '';
          }
        }
        continue;
      }

      var speakerLabel = speakerMatch[1];
      var content = paragraph.substring(speakerLabel.length).trim();

      // 1. 連続した繰り返し表現を削除
      var repeatPattern = /(.{5,20})\s*\1/g;
      while (repeatPattern.test(content)) {
        content = content.replace(repeatPattern, "$1");
      }

      // 2. 不自然な単語分割を修正
      content = content.replace(/([一-龯ぁ-ゔァ-ヴー])\s+([一-龯ぁ-ゔァ-ヴー])/g, "$1$2");

      // 3. 余分な空白の削除
      content = content.replace(/\s+/g, " ").trim();

      // 4. 句読点の整理
      content = content.replace(/。\s*。/g, "。");
      content = content.replace(/、\s*、/g, "、");
      content = content.replace(/。\s*、/g, "。");
      content = content.replace(/、\s*。/g, "。");

      // 話者ラベルと処理済みコンテンツを結合
      processed.push(speakerLabel + " " + content);
    }

    return processed.join("\n\n");
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
      const role = speakerRoles[speaker];
      if (role === 'sales') {
        return '【営業担当者】';
      } else if (role === 'customer') {
        return '【お客様】';
      } else {
        return '【話者' + speaker + '】';
      }
    }

    // デフォルトのフォーマット
    return '【話者' + speaker + '】';
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
   * セグメント情報がない場合に、テキストベースでマージする
   * @param {Object} openaiResult - GPT-4.1 mini Transcribeの結果
   * @param {Object} assemblyAIResult - AssemblyAIの結果
   * @return {Object} - マージした結果
   */
  function textBasedMerge(openaiResult, assemblyAIResult) {
    // GPT-4.1 miniに渡すための下準備としての処理かどうかを判定
    const isForGpt41Mini = (new Error().stack || '').includes('enhanceDialogueWithGPT4Mini');

    if (isForGpt41Mini) {
      Logger.log('GPT-4.1 mini用の下準備として初期マージデータを作成中...');
    } else {
      Logger.log('テキストベースのマージ処理を実行します...');
    }

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
      const speakerRoles = InformationExtractor.identifySpeakerRoles(utterances);
      const speakerInfo = InformationExtractor.extractSpeakerInfoFromConversation(utterances);

      // OpenAIのテキストを文ごとに分割
      const sentences = openaiResult.text.split(/(?<=[。．！？])\s*/);

      // 文章を話者のタイミングに基づいて分割
      const speakerSegments = splitTextBySpeakerTiming(sentences, utterances);

      // 話者情報を付与したテキストを生成
      let mergedText = '';

      for (let i = 0; i < speakerSegments.length; i++) {
        const segment = speakerSegments[i];
        // 話者IDから役割ラベルを生成
        const speakerId = segment.speaker;
        const role = speakerRoles[speakerId];
        const info = speakerInfo[speakerId] || {};

        // 役割ラベル
        const roleLabel = role === 'sales' ? '営業担当者' : (role === 'customer' ? 'お客様' : '話者' + speakerId);

        mergedText += `【${roleLabel}】 ${segment.text}\n\n`;
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
      // エラー時はOpenAIの結果を簡易フォーマットして返す
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

  // トランスクリプションサービスをエクスポート
  return {
    transcribe: function (file, assemblyAiApiKey, openaiApiKey) {
      try {
        Logger.log('文字起こし処理開始: ' + file.getName());
        return transcribe(file, assemblyAiApiKey, openaiApiKey);
      } catch (error) {
        Logger.log('文字起こし処理でエラー: ' + error.toString());
        throw error;
      }
    },
    transcribeWithGPT4oMini: transcribeWithGPT4oMini,
    transcribeWithAssemblyAI: transcribeWithAssemblyAI,
    enhanceDialogueWithGPT4Mini: enhanceDialogueWithGPT4Mini,
    getSpeakerCount: getSpeakerCount,
    formatSimpleText: formatSimpleText,
    postProcessTranscription: postProcessTranscription,
    formatSpeakerLabel: formatSpeakerLabel,
    splitTextBySpeakerTiming: splitTextBySpeakerTiming,
    tryAlternativeAssemblyAIRequest: tryAlternativeAssemblyAIRequest,
    tryAssemblyAIV3Request: tryAssemblyAIV3Request,
    textBasedMerge: textBasedMerge
  };
})();