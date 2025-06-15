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
      const enhancedResult = enhanceDialogueWithGPT4Mini({text: whisperResult.text}, whisperResult, openaiApiKey);
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
        throw new Error('ファイルサイズが大きすぎます: ' + (fileSize / 1024 / 1024).toFixed(2) + 'MB (10MB以下にしてください)');
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
      Logger.log('Whisper文字起こし中にエラー: ' + error.toString());
      throw error;
    }
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

【営業会社候補】
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
・株式会社インフィニットマインド

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
      const separatedText = responseJson.choices[0].message.content;

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

【営業会社の候補リスト】
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
      const enhancedText = responseJson.choices[0].message.content;

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
    formatSpeakerLabel: formatSpeakerLabel
  };
})();