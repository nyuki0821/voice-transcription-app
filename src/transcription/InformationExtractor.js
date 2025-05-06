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

      // システムプロンプトの作成（日本時間を明示、ハルシネーション防止強化）
      var systemPrompt = "あなたは顧客との営業電話の会話から重要な情報を正確に抽出するアシスタントです。" +
        "提供される会話文から以下の情報を抽出し、必ず正確なJSONフォーマットのみで出力してください。" +
        "情報が見つからない場合や確実でない場合は、絶対に推測せず、必ず空文字列を返してください。" +
        "会話が不完全または短い場合は、特に慎重に判断し、不明確な情報は全て空欄としてください。" +
        "会話内の時間情報はすべて日本時間（JST/UTC+9）として解釈してください。" +
        "回答は必ず有効なJSONオブジェクトとして解析できる形式でなければなりません。説明や前後の文章は一切不要です。" +
        "必ず以下のフィールドを含めてください：\n" +
        "- sales_company: 営業側の会社名 \n" +
        "- sales_person: 営業担当者の名前 \n" +
        "- customer_company: 顧客側の会社名 \n" +
        "- customer_name: 顧客担当者の名前 \n" +
        "- call_status: 通話結果（アポイント、断り、継続検討、不明のいずれか）\n" +
        "- reason_for_refusal: 断りの場合の理由 (~であるため。)\n" +
        "- reason_for_appointment: アポイントの場合の理由 (~であるため。)\n" +
        "- summary: 会話の要約\n" +
        "\n判断基準：\n" +
        "- アポイント：顧客が次回の打ち合わせや説明に明示的に同意した場合\n" +
        "- 断り：顧客が明確に興味がない、必要ない、遠慮したいなどと明示的に述べた場合\n" +
        "- 継続検討：顧客が検討する、持ち帰る、後日連絡する、責任者へ確認するなどと明示的に述べた場合\n" +
        "- 不明：上記に該当しない場合や、会話が不完全な場合\n" +
        "\n重要：会話が短い、不完全、または明確な情報がない場合は、絶対に情報を推測や創作せず、不明なフィールドは全て空欄にしてください。";

      // ユーザープロンプト（会話文）
      var userPrompt = "以下の会話から情報を抽出してJSONフォーマットで返してください。すべての日時は日本時間として解釈してください。" +
        "この会話が不完全または短い場合は、推測せず、分からない情報は全て空欄としてください：\n\n" + text;

      // 現在の日本時間をコンテキストとして追加
      var now = new Date();
      var jstOffset = 9 * 60; // 日本時間はUTC+9（分単位）
      var nowJST = new Date(now.getTime() + (jstOffset * 60000));
      var contextPrompt = "現在の日本時間は " +
        Utilities.formatDate(nowJST, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm') +
        " です。この情報を参考に、会話内の相対的な時間表現（「来週」「明日」など）を解釈してください。";

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
              content: contextPrompt + "\n\n" + userPrompt
            }
          ],
          temperature: 0.0, // ハルシネーション防止のため低温度に設定
          max_tokens: 1024,
          response_format: { "type": "json_object" } // JSON形式で返すように明示的に指定
        }),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('OpenAI APIからのレスポンスエラー: ' + responseCode + ', ' + response.getContentText());
      }

      // レスポンスをJSONとしてパース
      var responseJson = JSON.parse(response.getContentText());

      // OpenAI APIからのレスポンス構造に合わせて処理
      if (!responseJson.choices || responseJson.choices.length === 0) {
        throw new Error('OpenAI APIからの有効なレスポンスがありません');
      }

      var extractedTextJson = responseJson.choices[0].message.content;

      try {
        var extractedInfo = JSON.parse(extractedTextJson);
      } catch (parseError) {
        // JSONパースに失敗した場合は、最低限の情報を含む空のオブジェクトを返す
        extractedInfo = {
          sales_company: "",
          sales_person: "",
          customer_company: "",
          customer_name: "",
          call_status: "不明",
          reason_for_refusal: "",
          reason_for_appointment: "",
          summary: "JSONの解析に失敗しました。通話内容を手動で確認してください。"
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
        summary: extractedInfo.summary || text // 要約がない場合は元テキストをそのまま返す
      };

      // 会話が短い場合の追加検証（ハルシネーション防止）
      if (text.length < 100) {
        // 特に短い会話では、不審に具体的な情報をエラーとして検出
        var suspiciousInfo = false;

        // 会話に含まれない会社名が抽出されていないか確認
        if (validatedInfo.salesCompany && text.indexOf(validatedInfo.salesCompany) === -1) {
          validatedInfo.salesCompany = "";
          suspiciousInfo = true;
        }

        // 会話に含まれない担当者名が抽出されていないか確認
        if (validatedInfo.salesPerson && text.indexOf(validatedInfo.salesPerson) === -1) {
          validatedInfo.salesPerson = "";
          suspiciousInfo = true;
        }

        if (suspiciousInfo) {
          Logger.log('短い会話から不審な情報抽出が検出されました。抽出結果をリセットします: ' + text);
          validatedInfo.callStatus = "不明";
          validatedInfo.summary = text; // 元のテキストをそのまま返す
        }
      }

      return validatedInfo;
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
    var allowedStatuses = ["アポイント", "断り", "継続検討", "不明"];

    if (allowedStatuses.indexOf(normalizedStatus) !== -1) {
      return normalizedStatus;
    }

    // 類似のステータスを標準化
    if (/アポ|訪問|面談|打ち合わせ/.test(normalizedStatus)) {
      return "アポイント";
    } else if (/断り|拒否|お断り|見送り|不要|必要ない/.test(normalizedStatus)) {
      return "断り";
    } else if (/検討|持ち帰|考え|相談/.test(normalizedStatus)) {
      return "継続検討";
    }

    return "不明";
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
    var analysisLimit = Math.min(10, utterances.length); // 最初の10発言のみ分析

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

  // 公開メソッド
  return {
    extract: extract,
    identifySpeakerRoles: identifySpeakerRoles,
    extractSpeakerInfoFromConversation: extractSpeakerInfoFromConversation
  };
})();