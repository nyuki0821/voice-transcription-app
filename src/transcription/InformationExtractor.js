/**
 * 情報抽出モジュール（OpenAI GPT-4.1 専用版、日本時間対応）
 */
var InformationExtractor = (function () {
  /**
   * 文字起こしテキストから情報を抽出する
   * @param {string} text - 文字起こしテキスト
   * @param {Array} utterances - 話者分離情報の配列（互換性のために残すが使用しない）
   * @param {string} openaiApiKey - OpenAI APIキー（オプション、設定から取得可能）
   * @return {Object} - 抽出された情報
   */
  function extract(text, utterances, openaiApiKey) {
    if (!text) {
      throw new Error('テキストが指定されていません');
    }

    // APIキーが直接渡されない場合は設定から取得
    if (!openaiApiKey) {
      var settings = getSystemSettings();
      openaiApiKey = settings.OPENAI_API_KEY;
    }

    if (!openaiApiKey) {
      throw new Error('OpenAI APIキーが指定されていません');
    }

    try {
      // 短すぎるテキストの場合は早期リターン（ハルシネーション防止）
      if (text.length < 20) {
        Logger.log('テキストが短すぎるため解析を省略します: ' + text);
        return {
          salesCompany: "",
          salesPerson: "",
          customerCompany: "",
          customerName: "",
          callStatus: "不明",
          reasonForRefusal: "",
          reasonForAppointment: "",
          summary: text // 元のテキストをそのまま返す
        };
      }

      // OpenAI APIのエンドポイントURL
      var url = 'https://api.openai.com/v1/chat/completions';

      // システムプロンプトの簡略化版
      var systemPrompt = "あなたは営業電話の会話から重要な情報を抽出する専門家です。以下のステップで情報を正確に抽出してください。\n\n" +
        "【思考ステップ】\n" +
        "1. 会話全体を注意深く読み、話の流れと話者を特定する\n" +
        "2. 営業側と顧客側を明確に区別する\n" +
        "3. 会社名や担当者名が明示的に述べられている部分を特定する\n" +
        "4. 会話の結論（コンタクト、アポイント、断り、継続検討など）を判断する\n" +
        "5. 会話の主要なポイントを要約する\n\n" +
        "【行動ステップ】\n" +
        "1. 特定した情報をJSONフィールドに適切に割り当てる\n" +
        "2. 確証がない情報は空欄にする（推測しない）\n" +
        "3. call_statusは定義に従って最も適切な値を選択する\n" +
        "4. 思考プロセスは含めず、JSONオブジェクトのみを出力する\n\n" +
        "【JSON形式】\n" +
        "{\"sales_company\":\"\", \"sales_person\":\"\", \"customer_company\":\"\", \"customer_name\":\"\", " +
        "\"call_status\":\"\", \"reason_for_refusal\":\"\", \"reason_for_appointment\":\"\", \"summary\":\"\"}\n\n" +
        "【call_statusの定義】\n" +
        "・アポイント：顧客が次回の打ち合わせに明確に同意した場合のみ\n" +
        "・断り：顧客が明確に興味がない、必要ないと断った場合のみ\n" +
        "・継続検討：担当者不在、忙しい、検討する、持ち帰る、後日連絡する等の場合\n" +
        "・コンタクト：担当者本人と直接話ができた場合（最初に出た人が受付や他の人でなく担当者本人だった、または担当者に取り次いでもらえた場合）。但し、アポイントや断りには至らなかった場合のみ\n" +
        "・不明：会話から判断できない場合のみ\n\n" +
        "【重要なルール】\n" +
        "・reason_for_refusalは断りの場合のみ入力\n" +
        "・reason_for_appointmentはアポイントの場合のみ入力\n" +
        "・確実な情報のみを抽出し、不明な場合は空欄にする\n" +
        "・思考プロセスは出力に含めず、JSONオブジェクトのみを返すこと";

      // ユーザープロンプト（思考プロセスを促す）
      var userPrompt = "以下の営業電話の会話から重要な情報を抽出してください。\n\n" +
        "1. まず営業担当者と顧客を特定し、それぞれの会社名と名前を見つけてください\n" +
        "2. 次に会話の結果（アポイント、断り、継続検討、コンタクトのいずれか）を判断してください\n" +
        "3. 断りの場合はその理由、アポイントの場合はその詳細を特定してください\n" +
        "4. 最後に会話の要点を簡潔にまとめてください\n\n" +
        "会話内容：\n\n" + text;

      // 現在の日本時間をコンテキストとして追加
      var now = new Date();
      var jstOffset = 9 * 60; // 日本時間はUTC+9（分単位）
      var nowJST = new Date(now.getTime() + (jstOffset * 60000));
      var contextPrompt = "現在の日本時間は " +
        Utilities.formatDate(nowJST, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm') +
        " です。";

      // APIリクエストオプション
      var options = {
        method: 'post',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + openaiApiKey
        },
        payload: JSON.stringify({
          model: "gpt-4.1-mini",
          messages: [
            {
              role: "system",
              content: systemPrompt
            },
            {
              role: "user",
              content: contextPrompt + "\n\n" + userPrompt
            }
          ],
          temperature: 0,
          max_tokens: 1024,
          response_format: { "type": "json_object" }
        }),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('OpenAI APIからのレスポンスエラー: ' + responseCode);
      }

      // レスポンスをJSONとしてパース
      var responseJson = JSON.parse(response.getContentText());
      var extractedTextJson = responseJson.choices[0].message.content;
      var extractedInfo;

      try {
        extractedInfo = JSON.parse(extractedTextJson);
      } catch (parseError) {
        // JSONパースに失敗した場合の最低限の情報
        extractedInfo = {
          sales_company: "",
          sales_person: "",
          customer_company: "",
          customer_name: "",
          call_status: "不明",
          reason_for_refusal: "",
          reason_for_appointment: "",
          summary: "JSONの解析に失敗しました。"
        };
      }

      // 抽出されたデータの検証とデフォルト値の設定
      var validatedInfo = {
        salesCompany: extractedInfo.sales_company || "",
        salesPerson: extractedInfo.sales_person || "",
        customerCompany: extractedInfo.customer_company || "",
        customerName: extractedInfo.customer_name || "",
        callStatus: validateCallStatus(extractedInfo.call_status) || "不明",
        reasonForRefusal: extractedInfo.reason_for_refusal || "",
        reasonForAppointment: extractedInfo.reason_for_appointment || "",
        summary: extractedInfo.summary || text
      };

      // コールステータスに基づいた追加検証
      // 断り以外の場合はreason_for_refusalを空にする
      if (validatedInfo.callStatus !== "断り") {
        validatedInfo.reasonForRefusal = "";
      }

      // アポイント以外の場合はreason_for_appointmentを空にする
      if (validatedInfo.callStatus !== "アポイント") {
        validatedInfo.reasonForAppointment = "";
      }

      // 抽出情報のポスト処理
      var cleanInfo = postProcessExtractedInfo(validatedInfo);

      // 会話が短い場合の追加検証（ハルシネーション防止）
      if (text.length < 100) {
        // 特に短い会話では、不審に具体的な情報をエラーとして検出
        var suspiciousInfo = false;

        // 会話に含まれない会社名が抽出されていないか確認
        if (cleanInfo.salesCompany && text.indexOf(cleanInfo.salesCompany) === -1) {
          cleanInfo.salesCompany = "";
          suspiciousInfo = true;
        }

        // 会話に含まれない担当者名が抽出されていないか確認
        if (cleanInfo.salesPerson && text.indexOf(cleanInfo.salesPerson) === -1) {
          cleanInfo.salesPerson = "";
          suspiciousInfo = true;
        }

        if (suspiciousInfo) {
          Logger.log('短い会話から不審な情報抽出が検出されました。抽出結果をリセットします: ' + text);
          cleanInfo.callStatus = "不明";
          cleanInfo.summary = text; // 元のテキストをそのまま返す
        }
      }

      return cleanInfo;
    } catch (error) {
      throw new Error('AIによる情報抽出中にエラー: ' + error.toString());
    }
  }

  /**
   * call_statusの値を検証する
   * @param {string} status - 抽出されたcall_status
   * @return {string} - 検証済みのcall_status
   */
  function validateCallStatus(status) {
    if (!status) return "不明";

    var normalizedStatus = String(status).trim();

    // 許可されたステータス値
    var allowedStatuses = ["アポイント", "断り", "継続検討", "コンタクト", "不明"];

    if (allowedStatuses.indexOf(normalizedStatus) !== -1) {
      return normalizedStatus;
    }

    // 類似のステータスを標準化（正規表現を最適化）
    if (/アポ|訪問|面談|打ち合わせ|日程/.test(normalizedStatus)) {
      return "アポイント";
    } else if (/断り|拒否|お断り|不要|必要ない|興味がない/.test(normalizedStatus)) {
      return "断り";
    } else if (/検討|持ち帰|考え|相談|不在|忙しい|後日/.test(normalizedStatus)) {
      return "継続検討";
    } else if (/コンタクト|接触|担当者|本人|話した|話せた|取り次|代わった/.test(normalizedStatus)) {
      return "コンタクト";
    }

    // デフォルトでは継続検討として扱う
    return "継続検討";
  }

  /**
   * 話者の役割を識別する（営業側かお客様側か）
   * @param {Array} utterances - 話者分離情報の配列
   * @return {Object} - 話者IDと役割のマッピング
   */
  function identifySpeakerRoles(utterances) {
    if (!utterances || utterances.length === 0) {
      return {};
    }

    var speakerRoles = {};

    // 各話者の発言を集約
    var speakerTexts = {};
    var firstUtterances = {}; // 各話者の最初の発言を保存
    var speakerUtteranceCounts = {}; // 各話者の発言回数
    var speakerTotalLength = {}; // 各話者の総発言文字数

    for (var i = 0; i < utterances.length; i++) {
      var utterance = utterances[i];
      var speaker = utterance.speaker;
      var text = utterance.text || "";

      if (!speakerTexts[speaker]) {
        speakerTexts[speaker] = '';
        firstUtterances[speaker] = text;
        speakerUtteranceCounts[speaker] = 0;
        speakerTotalLength[speaker] = 0;
      }

      speakerTexts[speaker] += text + ' ';
      speakerUtteranceCounts[speaker]++;
      speakerTotalLength[speaker] += text.length;
    }

    // 営業担当者特有のパターン (より強力な判断材料)
    var salesPatterns = [
      // 提案・案内フレーズ
      /ご案内で.*ご連絡/i,
      /ご連絡いたしました/i,
      /お電話いたしました/i,
      /ご紹介させていただ/i,
      /ご提案させていただ/i,

      // 自社言及
      /弊社は/i,
      /弊社では/i,
      /私どもは/i,
      /私どもの/i,
      /グループの一員として/i,

      // サービス紹介
      /提供して(おり|い)/i,
      /お手伝いさせていただ/i,
      /ご提供して/i,
      /システムをご提供/i,

      // 特徴的なフレーズ
      /おめでとうございます/i,
      /業務をお手伝い/i,
      /認定.*おめでとう/i,

      // 営業担当者の典型的な表現
      /失礼いたします/i,
      /お忙しいところ恐れ入り/i,
      /お時間よろしいでしょうか/i,
      /ご興味(は|が)あります/i
    ];

    // お客様特有のパターン
    var customerPatterns = [
      // 対応・確認フレーズ
      /確認いたします/i,
      /少々お待ちください/i,
      /検討いたします/i,

      // 拒否・不要フレーズ
      /不要です/i,
      /必要ありません/i,
      /結構です/i,
      /検討して/i,

      // 会社内部情報
      /関連会社/i,
      /弊社でも同様のシステム/i,
      /同じようなシステム/i,
      /社内で/i,

      // 特徴的な返答
      /そうですか/i,
      /なるほど/i,
      /承知いたしました/i,
      /かしこまりました/i,

      // お客様の典型的な反応
      /今は考えてい(ない|ません)/i,
      /予算が(ない|厳しい)/i,
      /他社(で|と)契約/i,
      /担当者に(確認|連絡)/i
    ];

    // 自己紹介パターンの検出（会社名・個人名の抽出方法）
    var introductionPatterns = {
      // 自己紹介は役割判定に重要だが、特定の会社名に依存しない汎用的なパターン
      sales: [
        /私.*(と申します|です)/i,
        /私は.*の.*と申します/i,
        /の者です/i,
        /と申します.*ご案内/i
      ],
      customer: [
        /こちら.*担当の/i,
        /を担当して/i,
        /にお繋ぎし/i
      ]
    };

    // 各話者のスコアを計算
    var speakerScores = {};

    for (var speaker in speakerTexts) {
      speakerScores[speaker] = {
        salesScore: 0,
        customerScore: 0
      };

      var text = speakerTexts[speaker];
      var firstText = firstUtterances[speaker];

      // 営業担当者パターンのマッチをチェック
      for (var i = 0; i < salesPatterns.length; i++) {
        if (text.match(salesPatterns[i])) {
          speakerScores[speaker].salesScore += 2; // 重み付けスコア
        }
      }

      // お客様パターンのマッチをチェック
      for (var i = 0; i < customerPatterns.length; i++) {
        if (text.match(customerPatterns[i])) {
          speakerScores[speaker].customerScore += 2; // 重み付けスコア
        }
      }

      // 自己紹介パターンのチェック
      for (var i = 0; i < introductionPatterns.sales.length; i++) {
        if (text.match(introductionPatterns.sales[i])) {
          speakerScores[speaker].salesScore += 2; // 自己紹介は重要な指標
        }
      }

      for (var i = 0; i < introductionPatterns.customer.length; i++) {
        if (text.match(introductionPatterns.customer[i])) {
          speakerScores[speaker].customerScore += 2; // 自己紹介は重要な指標
        }
      }

      // 発言量の分析（営業担当者は通常説明が長い）
      var avgLength = speakerTotalLength[speaker] / speakerUtteranceCounts[speaker];
      if (avgLength > 50) { // 平均発言が長い場合は営業側の可能性が高い
        speakerScores[speaker].salesScore += 1;
      } else {
        speakerScores[speaker].customerScore += 1;
      }

      // 最初の発言が短い（10文字以下）場合は顧客の可能性が高い
      if (firstText.length <= 10) {
        speakerScores[speaker].customerScore += 1;
      }
    }

    // スコアに基づいて役割を決定
    for (var speaker in speakerScores) {
      var scores = speakerScores[speaker];

      if (scores.salesScore > scores.customerScore) {
        speakerRoles[speaker] = 'sales';
      } else if (scores.customerScore > scores.salesScore) {
        speakerRoles[speaker] = 'customer';
      } else {
        // スコアが同じ場合は、発言量・パターンなどから推測
        // 営業担当者は通常、最初に自社の説明などで長めの発言をする
        if (speakerTotalLength[speaker] / speakerUtteranceCounts[speaker] > 40) {
          speakerRoles[speaker] = 'sales';
        } else {
          speakerRoles[speaker] = 'customer';
        }
      }
    }

    // 2人の話者の場合、両方が同じ役割にならないよう調整
    var speakers = Object.keys(speakerRoles);
    if (speakers.length === 2 && speakerRoles[speakers[0]] === speakerRoles[speakers[1]]) {
      // スコアの差が大きい方を優先
      var scores0 = speakerScores[speakers[0]];
      var scores1 = speakerScores[speakers[1]];

      var diff0 = Math.abs(scores0.salesScore - scores0.customerScore);
      var diff1 = Math.abs(scores1.salesScore - scores1.customerScore);

      if (diff0 >= diff1) {
        // speaker0のスコア差が大きい場合はその判定を維持し、speaker1を反対の役割に
        speakerRoles[speakers[1]] = (speakerRoles[speakers[0]] === 'sales') ? 'customer' : 'sales';
      } else {
        // speaker1のスコア差が大きい場合はその判定を維持し、speaker0を反対の役割に
        speakerRoles[speakers[0]] = (speakerRoles[speakers[1]] === 'sales') ? 'customer' : 'sales';
      }
    }

    return speakerRoles;
  }

  /**
   * 会話内容から話者の会社名と担当者名を抽出する
   * @param {Array} utterances - 話者分離情報の配列
   * @return {Object} - 話者IDと会社名・担当者名のマッピング
   */
  function extractSpeakerInfoFromConversation(utterances) {
    if (!utterances || utterances.length === 0) {
      return {};
    }

    var speakerInfo = {};
    var companyPatterns = [
      /([^、。\s]+)株式会社/,
      /株式会社([^、。\s]+)/,
      /([^、。\s]+)の者です/,
      /([^、。\s]+)の担当/,
      /([^、。\s]+)から来ました/,
      /([^、。\s]+)の([^、。\s]+)と申します/,
      /([^、。\s]+)の([^、。\s]+)です/,
      // より一般的な会社名抽出パターン
      /私は([^、。\s]+)から/,
      /([^、。\s]+)の者で/,
      /私、([^、。\s]+)の/,
      /([^、。\s]+)(の|という|っていう|という会社|という企業|という団体)/
    ];

    var namePatterns = [
      /私は([^\s、。]+)と申します/,
      /([^\s、。]+)と申します/,
      /私、([^\s、。]+)と/,
      /([^\s、。]+)です/,
      /([^\s、。]+)と言います/,
      /([^\s、。]+)といいます/,
      // より一般的な担当者名抽出パターン
      /私、([^\s、。]+)(です|でございます)/,
      /担当の([^\s、。]+)(です|でございます)/,
      /担当([^\s、。]+)(です|でございます)/,
      /([^\s、。]+)(と申します|です)(が|けど|けれど)/
    ];

    // 各話者の最初の数発言を分析（最初の方が自己紹介が多い）
    var processedSpeakers = {};
    var analysisLimit = Math.min(10, utterances.length);

    for (var i = 0; i < analysisLimit; i++) {
      var utterance = utterances[i];
      var speaker = utterance.speaker;
      var text = utterance.text;

      // 既に会社名と担当者名の両方が特定できた話者はスキップ
      if (processedSpeakers[speaker] === true) {
        continue;
      }

      // 話者情報がまだ初期化されていない場合は初期化
      if (!speakerInfo[speaker]) {
        speakerInfo[speaker] = {
          company: '',
          name: ''
        };
        processedSpeakers[speaker] = false;
      }

      // 会社名の抽出（まだ抽出されていない場合のみ）
      if (!speakerInfo[speaker].company) {
        for (var j = 0; j < companyPatterns.length; j++) {
          var companyMatch = text.match(companyPatterns[j]);
          if (companyMatch && companyMatch[1]) {
            var company = companyMatch[1];
            // 余分な文字を削除
            if (company.endsWith('の')) {
              company = company.substring(0, company.length - 1);
            }
            speakerInfo[speaker].company = company;
            break;
          }
        }
      }

      // 担当者名の抽出（まだ抽出されていない場合のみ）
      if (!speakerInfo[speaker].name) {
        for (var j = 0; j < namePatterns.length; j++) {
          var nameMatch = text.match(namePatterns[j]);
          if (nameMatch && nameMatch[1]) {
            speakerInfo[speaker].name = nameMatch[1];
            break;
          }
        }
      }

      // 会社名と担当者名の両方が抽出できた場合は、この話者の処理を完了とマーク
      if (speakerInfo[speaker].company && speakerInfo[speaker].name) {
        processedSpeakers[speaker] = true;
      }
    }

    // 情報が抽出できなかった話者は削除
    for (var speaker in speakerInfo) {
      if (!speakerInfo[speaker].company && !speakerInfo[speaker].name) {
        delete speakerInfo[speaker];
      }
    }

    return speakerInfo;
  }

  /**
   * 抽出された情報のポスト処理を行う
   * @param {Object} info - 抽出された情報
   * @return {Object} - 処理後の情報
   */
  function postProcessExtractedInfo(info) {
    if (!info) return info;

    // 不自然な単語や表現を修正
    var cleanInfo = {
      salesCompany: cleanText(info.salesCompany),
      salesPerson: cleanText(info.salesPerson),
      customerCompany: cleanText(info.customerCompany),
      customerName: cleanText(info.customerName),
      callStatus: info.callStatus,
      reasonForRefusal: cleanText(info.reasonForRefusal),
      reasonForAppointment: cleanText(info.reasonForAppointment),
      summary: cleanText(info.summary)
    };

    return cleanInfo;
  }

  /**
   * テキストを自然な日本語に整形する
   * @param {string} text - 整形するテキスト
   * @return {string} - 整形後のテキスト
   */
  function cleanText(text) {
    if (!text) return "";

    var cleaned = text;

    // 不自然な数字の間の空白を削除
    cleaned = cleaned.replace(/([A-Za-z])\s+(\d)/g, "$1$2");

    // 連続する繰り返し表現を削除
    var repeatPattern = /(.{5,20})\s*\1/g;
    while (repeatPattern.test(cleaned)) {
      cleaned = cleaned.replace(repeatPattern, "$1");
    }

    // 不自然な単語分割を修正
    cleaned = cleaned.replace(/([一-龯ぁ-ゔァ-ヴー])\s+([一-龯ぁ-ゔァ-ヴー])/g, "$1$2");

    return cleaned;
  }

  // 公開メソッド
  return {
    extract: extract,
    identifySpeakerRoles: identifySpeakerRoles,
    extractSpeakerInfoFromConversation: extractSpeakerInfoFromConversation
  };
})();