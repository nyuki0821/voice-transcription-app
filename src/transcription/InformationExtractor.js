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
      // テキストが非常に短い場合も空文字列を出力
      if (text.length < 20) {
        Logger.log('テキストが短すぎるため解析を省略します: ' + text);
        return {
          sales_company: "",
          sales_person: "", // vlookup関数を使うため常に空文字列
          customer_company: "", // vlookup関数を使うため常に空文字列
          customer_name: "",
          call_status1: "非コンタクト", // 短すぎる会話は非コンタクトに変更
          call_status2: "",
          reason_for_refusal: "",
          reason_for_refusal_category: "",
          reason_for_appointment: "",
          reason_for_appointment_category: "",
          summary: text // 元のテキストをそのまま返す
        };
      }

      // 会話が短い、または中断されている場合の簡易判定
      // 会話の長さの制限を緩和し、より多くの会話で情報抽出を実行する
      if ((text.length < 100 && /また改め|また連絡|失礼します|失礼いたします|お時間.*ない|急いで|後で|担当者.*いない|取次.*できない/.test(text)) ||
        (text.length < 50)) {
        // 会話が非常に短い、または中断を示す表現がある場合

        // 対話形式の会話かどうかを確認（会話らしいパターンを検出）
        var hasConversationPattern = /【営業担当】.*【お客様】|【お客様】.*【営業担当】|［営業担当］.*［お客様］|［お客様］.*［営業担当］|\[営業担当\].*\[お客様\]|\[お客様\].*\[営業担当\]|営業担当.*お客様|お客様.*営業担当/.test(text);

        // 会話には見えても製品説明がない場合のみ非コンタクト判定
        if (!hasConversationPattern || (text.length < 80 && !/製品|サービス|ソリューション|事業|ご提案|ご案内|弊社|特徴|機能|導入|ご説明/.test(text))) {
          Logger.log('会話が非常に短いまたは対話形式ではないため非コンタクト判定: ' + text);
          return {
            sales_company: "",
            sales_person: "", // vlookup関数を使うため常に空文字列
            customer_company: "", // vlookup関数を使うため常に空文字列
            customer_name: "",
            call_status1: "非コンタクト",
            call_status2: "",
            reason_for_refusal: "",
            reason_for_refusal_category: "",
            reason_for_appointment: "",
            reason_for_appointment_category: "",
            summary: "会話が短いまたは中断されているため詳細な説明ができていません。"
          };
        }
      }

      // OpenAI APIのエンドポイントURL
      var url = 'https://api.openai.com/v1/chat/completions';

      // システムプロンプトの簡略化版
      var systemPrompt = "あなたは営業電話の会話から重要な情報を抽出する専門家です。以下のステップで情報を正確に抽出してください。\n\n" +
        "【思考ステップ】\n" +
        "1. 会話全体を注意深く読み、話の流れと話者を特定する\n" +
        "2. 営業側と顧客側を明確に区別する\n" +
        "3. 会社名や担当者名が明示的に述べられている部分を特定する\n" +
        "4. 会話の結論を以下の2つの観点で判断する：\n" +
        "   - まず、担当者と製品・サービスの説明ができたか（コンタクト）\n" +
        "   - 次に、コンタクトした場合の会話結果（アポイント/断り/継続）\n" +
        "5. 会話の主要なポイントを要約する\n\n" +
        "【行動ステップ】\n" +
        "1. 特定した情報をJSONフィールドに適切に割り当てる\n" +
        "2. 確証がない情報は空欄にする（推測しない）\n" +
        "3. call_status1とcall_status2は必ず定義に従って適切な値を選択する\n" +
        "4. 断り理由やアポイント理由については、その内容と適切なカテゴリを選定する\n" +
        "5. 思考プロセスは含めず、JSONオブジェクトのみを出力する\n\n" +
        "【JSON形式】\n" +
        "{\"sales_company\":\"\", \"sales_person\":\"\", \"customer_company\":\"\", \"customer_name\":\"\", " +
        "\"call_status1\":\"\", \"call_status2\":\"\", \"reason_for_refusal\":\"\", \"reason_for_refusal_category\":\"\", " +
        "\"reason_for_appointment\":\"\", \"reason_for_appointment_category\":\"\", \"summary\":\"\"}\n\n" +
        "【sales_companyの候補リスト】\n" +
        "営業会社名は以下のリストから最も適切なものを選んでください。会話中に名前が明示的に出ていない場合は空欄にしてください：\n" +
        "・株式会社ENERALL\n" +
        "・エムスリーヘルスデザイン株式会社\n" +
        "・株式会社TOKIUM\n" +
        "・株式会社グッドワークス\n" +
        "・テコム看護\n" +
        "・ハローワールド株式会社\n" +
        "・株式会社ワーサル\n" +
        "・株式会社NOTCH\n" +
        "・株式会社ジースタイラス\n" +
        "・株式会社佑人社\n\n" +
        "【call_statusの定義】\n" +
        "・call_status1：\n" +
        "  ・コンタクト: 担当者に関わらず、製品・サービスの詳細な説明ができた場合。単に「ご案内」と言っただけでは不十分で、製品・サービスの具体的な内容や特徴について説明できた場合のみ。\n" +
        "  ・非コンタクト: 自動音声のみや、受付で止まった場合。製品・サービスの説明ができなかった場合。担当者不在や取次ぎ失敗で詳細説明ができなかった場合も含む。会話が短く中断された場合も非コンタクト。\n" +
        "・call_status2（call_status1が「コンタクト」の場合のみ入力）：\n" +
        "  ・アポイント: 次回のアポイントを取得できた場合\n" +
        "  ・断り: 明確に興味がない、必要ないと断られた場合\n" +
        "  ・継続: 検討する、持ち帰る等で終話した場合\n" +
        "・不明: 文章が短すぎる等で判断できない場合は不明と入力\n\n" +
        "【reason_for_refusal_categoryの定義】\n" +
        "断りの理由を以下のカテゴリから選択する（断りの場合のみ入力）：\n" +
        "・忙しい\n" +
        "・断るよう言われている\n" +
        "・必要性を感じない・興味を感じない・予定がない\n" +
        "・導入済みで切替予定なし\n" +
        "・別会社・サービスで検討を進めている\n" +
        "・価格が高い\n" +
        "・その他\n\n" +
        "【reason_for_appointment_categoryの定義】\n" +
        "アポイントの理由を以下のカテゴリから選択する（アポイントの場合のみ入力）：\n" +
        "・必要性を感じていた\n" +
        "・価格に魅力を感じる\n" +
        "・サービスに興味がある\n" +
        "・とりあえず話を聞いてみる\n" +
        "・その他\n\n" +
        "【重要なルール】\n" +
        "・営業会社名（sales_company）は指定リストからのみ選択し、会話に明示的に出ていない場合は空欄\n" +
        "・reason_for_refusalは断りの場合のみ入力\n" +
        "・reason_for_refusal_categoryは断りの場合のみ入力\n" +
        "・reason_for_appointmentはアポイントの場合のみ入力\n" +
        "・reason_for_appointment_categoryはアポイントの場合のみ入力\n" +
        "・確実な情報のみを抽出し、不明な場合は空欄にする\n" +
        "・思考プロセスは出力に含めず、JSONオブジェクトのみを返すこと";

      // ユーザープロンプト（思考プロセスを促す）
      var userPrompt = "以下の営業電話の会話から重要な情報を抽出してください。\n\n" +
        "1. まず営業担当者と顧客を特定し、それぞれの会社名と名前を見つけてください\n" +
        "2. 次に会話の結果を2段階で判定してください：\n" +
        "   a) call_status1: 担当者と製品・サービスの説明ができたか（コンタクト/非コンタクト）\n" +
        "   b) call_status2: コンタクトの場合の成果（アポイント/断り/継続）\n" +
        "3. 断りの場合はその理由とカテゴリ、アポイントの場合はその理由とカテゴリを特定してください\n" +
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
          sales_person: "", // vlookup関数を使うため常に空文字列
          customer_company: "", // vlookup関数を使うため常に空文字列
          customer_name: "",
          call_status1: "不明",
          call_status2: "",
          reason_for_refusal: "",
          reason_for_refusal_category: "",
          reason_for_appointment: "",
          reason_for_appointment_category: "",
          summary: "JSONの解析に失敗しました。"
        };
      }

      // 抽出されたデータの検証とデフォルト値の設定
      var validatedInfo = {
        sales_company: validateSalesCompany(extractedInfo.sales_company) || "",
        sales_person: "", // vlookup関数を使うため常に空文字列
        customer_company: "", // vlookup関数を使うため常に空文字列
        customer_name: extractedInfo.customer_name || "",
        call_status1: validateCallStatus1(extractedInfo.call_status1) || "不明",
        call_status2: validateCallStatus2(extractedInfo.call_status2, extractedInfo.call_status1) || "",
        reason_for_refusal: extractedInfo.reason_for_refusal || "",
        reason_for_refusal_category: validateRefusalCategory(extractedInfo.reason_for_refusal_category) || "",
        reason_for_appointment: extractedInfo.reason_for_appointment || "",
        reason_for_appointment_category: validateAppointmentCategory(extractedInfo.reason_for_appointment_category) || "",
        summary: extractedInfo.summary || text
      };

      // コールステータスに基づいた追加検証
      // 断り以外の場合はreason_for_refusalとそのカテゴリを空にする
      if (validatedInfo.call_status2 !== "断り") {
        validatedInfo.reason_for_refusal = "";
        validatedInfo.reason_for_refusal_category = "";
      }

      // アポイント以外の場合はreason_for_appointmentとそのカテゴリを空にする
      if (validatedInfo.call_status2 !== "アポイント") {
        validatedInfo.reason_for_appointment = "";
        validatedInfo.reason_for_appointment_category = "";
      }

      // 非コンタクトの場合はcall_status2を空にする
      if (validatedInfo.call_status1 !== "コンタクト") {
        validatedInfo.call_status2 = "";
      }

      // 抽出情報のポスト処理
      var cleanInfo = postProcessExtractedInfo(validatedInfo);

      // デバッグ用：最終出力の内容をログに記録
      Logger.log('情報抽出結果 (JSON): ' + JSON.stringify(cleanInfo));

      // 会話が短い場合の追加検証（ハルシネーション防止）
      if (text.length < 80) {
        // 特に短い会話では、不審に具体的な情報をエラーとして検出
        var suspiciousInfo = false;

        // 会話に含まれない会社名が抽出されていないか確認
        if (cleanInfo.sales_company && text.indexOf(cleanInfo.sales_company) === -1 &&
          // 部分的な会社名も確認
          !text.includes('グッドワークス') &&
          !text.includes('ENERALL') &&
          !text.includes('トキウム') &&
          !text.includes('TOKIUM') &&
          !text.includes('ワーサル') &&
          !text.includes('NOTCH') &&
          !text.includes('ジースタイラス') &&
          !text.includes('佑人社') &&
          !text.includes('テコム') &&
          !text.includes('エムスリー')) {
          cleanInfo.sales_company = "";
          suspiciousInfo = true;
        }

        // 会話に含まれない担当者名が抽出されていないか確認
        if (cleanInfo.sales_person && text.indexOf(cleanInfo.sales_person) === -1) {
          cleanInfo.sales_person = "";
          suspiciousInfo = true;
        }

        if (suspiciousInfo) {
          Logger.log('短い会話から不審な情報抽出が検出されました。不審なフィールドをリセットします: ' + text);
          // AIが誤った情報を生成したと思われる場合は、その情報だけをリセット
          // ただし、会話の結果自体はそのまま保持
        }
      }

      return cleanInfo;
    } catch (error) {
      throw new Error('AIによる情報抽出中にエラー: ' + error.toString());
    }
  }

  /**
   * call_status1の値を検証する
   * @param {string} status - 抽出されたcall_status1
   * @return {string} - 検証済みのcall_status1
   */
  function validateCallStatus1(status) {
    if (!status) return "不明";

    var normalizedStatus = String(status).trim();

    // 許可されたステータス値（コンタクトと非コンタクトの二択のみ）
    var allowedStatuses = ["コンタクト", "非コンタクト", "不明"];

    if (allowedStatuses.indexOf(normalizedStatus) !== -1) {
      return normalizedStatus;
    }

    // call_status2に属するステータスがcall_status1に入ってしまった場合の対応
    if (/アポイント|断り|継続/.test(normalizedStatus)) {
      // デフォルトは「コンタクト」と見なす（アポ/断り/継続はコンタクト前提のため）
      return "コンタクト";
    }

    // 類似のステータスを標準化
    if (/製品.*詳細説明|サービス.*詳細説明|具体的.*説明|詳しく説明|プレゼン.*できた|十分.*紹介/.test(normalizedStatus)) {
      return "コンタクト";
    } else if (/自動音声|受付|説明.*できなかった|非コンタクト|取次.*失敗|担当者.*不在|担当者.*いない|また改め|短い会話|中断/.test(normalizedStatus)) {
      return "非コンタクト";
    }

    // デフォルトでは不明として扱う
    return "不明";
  }

  /**
   * call_status2の値を検証する
   * @param {string} status - 抽出されたcall_status2
   * @param {string} status1 - 抽出されたcall_status1
   * @return {string} - 検証済みのcall_status2
   */
  function validateCallStatus2(status, status1) {
    // call_status1が「コンタクト」でない場合は空文字を返す
    if (validateCallStatus1(status1) !== "コンタクト") {
      return "";
    }

    if (!status) return "継続"; // コンタクトの場合、デフォルトは継続

    var normalizedStatus = String(status).trim();

    // 許可されたステータス値
    var allowedStatuses = ["アポイント", "断り", "継続", "不明"];

    if (allowedStatuses.indexOf(normalizedStatus) !== -1) {
      return normalizedStatus;
    }

    // 類似のステータスを標準化
    if (/アポ|アポイント|訪問|面談|打ち合わせ|日程|伺|承諾/.test(normalizedStatus)) {
      return "アポイント";
    } else if (/断り|断わり|拒否|お断り|不要|必要ない|興味がない|結構/.test(normalizedStatus)) {
      return "断り";
    } else if (/検討|持ち帰|考え|相談|確認|後日|連絡|担当者|継続/.test(normalizedStatus)) {
      return "継続";
    }

    // デフォルトでは継続として扱う
    return "継続";
  }

  /**
   * reason_for_refusal_categoryの値を検証する
   * @param {string} category - 抽出された断り理由カテゴリ
   * @return {string} - 検証済みのカテゴリ
   */
  function validateRefusalCategory(category) {
    if (!category) return "";

    var normalizedCategory = String(category).trim();

    // 許可されたカテゴリ値
    var allowedCategories = [
      "忙しい",
      "断るよう言われている",
      "必要性を感じない・興味を感じない・予定がない",
      "導入済みで切替予定なし",
      "別会社・サービスで検討を進めている",
      "価格が高い",
      "その他"
    ];

    if (allowedCategories.indexOf(normalizedCategory) !== -1) {
      return normalizedCategory;
    }

    // 類似のカテゴリを標準化
    if (/忙しい|時間がない|多忙|スケジュール/.test(normalizedCategory)) {
      return "忙しい";
    } else if (/上司|上長|上の人|権限がない|指示|言われ/.test(normalizedCategory)) {
      return "断るよう言われている";
    } else if (/必要性|興味|ニーズ|予定|検討予定|予算|計画/.test(normalizedCategory)) {
      return "必要性を感じない・興味を感じない・予定がない";
    } else if (/導入済み|使っ|既に|十分|満足|切り替え/.test(normalizedCategory)) {
      return "導入済みで切替予定なし";
    } else if (/別|他社|競合|他のベンダー|すでに/.test(normalizedCategory)) {
      return "別会社・サービスで検討を進めている";
    } else if (/価格|コスト|高い|予算/.test(normalizedCategory)) {
      return "価格が高い";
    }

    // デフォルトではその他として扱う
    return "その他";
  }

  /**
   * reason_for_appointment_categoryの値を検証する
   * @param {string} category - 抽出されたアポイント理由カテゴリ
   * @return {string} - 検証済みのカテゴリ
   */
  function validateAppointmentCategory(category) {
    if (!category) return "";

    var normalizedCategory = String(category).trim();

    // 許可されたカテゴリ値
    var allowedCategories = [
      "必要性を感じていた",
      "価格に魅力を感じる",
      "サービスに興味がある",
      "とりあえず話を聞いてみる",
      "その他"
    ];

    if (allowedCategories.indexOf(normalizedCategory) !== -1) {
      return normalizedCategory;
    }

    // 類似のカテゴリを標準化
    if (/必要|課題|問題|解決|ニーズ/.test(normalizedCategory)) {
      return "必要性を感じていた";
    } else if (/価格|コスト|費用|安い|予算|節約/.test(normalizedCategory)) {
      return "価格に魅力を感じる";
    } else if (/興味|機能|特徴|メリット|魅力|ベネフィット/.test(normalizedCategory)) {
      return "サービスに興味がある";
    } else if (/とりあえず|話を聞く|情報収集|比較|検討|時間/.test(normalizedCategory)) {
      return "とりあえず話を聞いてみる";
    }

    // デフォルトではその他として扱う
    return "その他";
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
    var speakerLabels = {}; // 各話者のラベル（「営業担当」「お客様」など）

    for (var i = 0; i < utterances.length; i++) {
      var utterance = utterances[i];
      var speaker = utterance.speaker;
      var text = utterance.text || "";

      if (!speakerTexts[speaker]) {
        speakerTexts[speaker] = '';
        firstUtterances[speaker] = text;
        speakerUtteranceCounts[speaker] = 0;
        speakerTotalLength[speaker] = 0;
        speakerLabels[speaker] = '';
      }

      speakerTexts[speaker] += text + ' ';
      speakerUtteranceCounts[speaker]++;
      speakerTotalLength[speaker] += text.length;

      // 話者ラベルを検出（「営業担当」「お客様」などのマーカー）
      var labelMatch = text.match(/^[\[【［](.+?)[\]】］]/);
      if (labelMatch && !speakerLabels[speaker]) {
        speakerLabels[speaker] = labelMatch[1];
      }
    }

    // 許可リストの会社名（validateSalesCompanyから取得）
    var allowedCompanies = [
      "株式会社ENERALL",
      "エムスリーヘルスデザイン株式会社",
      "株式会社TOKIUM",
      "株式会社グッドワークス",
      "テコム看護",
      "ハローワールド株式会社",
      "株式会社ワーサル",
      "株式会社NOTCH",
      "株式会社ジースタイラス",
      "株式会社佑人社"
    ];

    // 会社名の特徴的な部分のリストを作成（「株式会社」を除いた部分など）
    var companyKeywords = allowedCompanies.map(function (company) {
      return company.replace(/株式会社|エムスリーヘルスデザイン|テコム看護|ハローワールド/g, "").trim();
    }).filter(function (keyword) {
      return keyword.length > 0;
    });

    // 短い会社名キーワードは除外（例: "の"など、誤検出しやすいもの）
    companyKeywords = companyKeywords.filter(function (keyword) {
      return keyword.length >= 2;
    });

    // 会社名の別表記も追加
    companyKeywords.push("ハローワールド");
    companyKeywords.push("エムスリー");
    companyKeywords.push("グッドワークス");
    companyKeywords.push("ジースタイラス");
    companyKeywords.push("ENERALL");
    companyKeywords.push("トキウム");
    companyKeywords.push("TOKIUM");
    companyKeywords.push("ワーサル");
    companyKeywords.push("NOTCH");
    companyKeywords.push("テコム");

    // 許可リスト会社名との関連性を各話者ごとに評価
    var companyAssociations = {};
    for (var speaker in speakerTexts) {
      companyAssociations[speaker] = {
        mentionedCompany: '',
        isSelfIntroduction: false,
        associationScore: 0
      };

      var text = speakerTexts[speaker];

      // 自己紹介パターンの検出
      var selfIntroPattern = /私は(.+?)(と申します|です|の者です|と言います|といいます|と申しますが|からです)/i;
      var selfIntroMatch = text.match(selfIntroPattern);

      // 会社名と自己紹介の関連を評価
      for (var i = 0; i < allowedCompanies.length; i++) {
        var company = allowedCompanies[i];
        if (text.indexOf(company) !== -1) {
          companyAssociations[speaker].mentionedCompany = company;
          companyAssociations[speaker].associationScore += 3;

          // 自己紹介との関連を評価
          if (selfIntroMatch && selfIntroMatch[1].indexOf(company.replace(/株式会社/g, "").trim()) !== -1) {
            companyAssociations[speaker].isSelfIntroduction = true;
            companyAssociations[speaker].associationScore += 5;
          }
        }
      }

      // 会社名キーワードとの部分一致も確認
      if (!companyAssociations[speaker].mentionedCompany) {
        for (var i = 0; i < companyKeywords.length; i++) {
          var keyword = companyKeywords[i];
          if (text.indexOf(keyword) !== -1) {
            // 自己紹介との関連を評価
            var selfIntroKeywordPattern = new RegExp(keyword + "(の|から|と申します|です|の者です)");
            if (selfIntroMatch && selfIntroMatch[1].match(selfIntroKeywordPattern)) {
              companyAssociations[speaker].isSelfIntroduction = true;
              companyAssociations[speaker].associationScore += 4;
              break;
            } else if (text.match(new RegExp(keyword + "(の|から).+?(と申します|です|の者です)"))) {
              companyAssociations[speaker].isSelfIntroduction = true;
              companyAssociations[speaker].associationScore += 3;
              break;
            } else {
              companyAssociations[speaker].associationScore += 1;
            }
          }
        }
      }
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
      /ご興味(は|が)あります/i,

      // 担当者在籍確認（営業側の典型的な質問）
      /いらっしゃいますでしょうか/i,
      /ご在席でしょうか/i,
      /お話できますでしょうか/i,
      /ご対応いただけますでしょうか/i,

      // 後日連絡系（営業側の特徴）
      /また(ご連絡|お電話|改めて)/i,
      /改めてご連絡/i,
      /後ほどお電話/i
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
      /担当者に(確認|連絡)/i,

      // 不在応答（お客様側の特徴）
      /席を外して/i,
      /不在です/i,
      /戻られる(の|ん)は/i,
      /時以降(なら|に|で|だと)/i
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
        customerScore: 0,
        confidence: 0 // 信頼度
      };

      var text = speakerTexts[speaker];
      var firstText = firstUtterances[speaker];
      var label = speakerLabels[speaker];

      // 話者ラベルがある場合は、それに基づいて初期スコアを付与
      if (label) {
        if (/営業|企業|会社|担当|オペレータ/.test(label)) {
          speakerScores[speaker].salesScore += 2;
          speakerScores[speaker].confidence += 1;
        } else if (/客|顧客|お客|ユーザ|利用者/.test(label)) {
          speakerScores[speaker].customerScore += 2;
          speakerScores[speaker].confidence += 1;
        }
      }

      // 許可リスト会社との関連性スコアを加算
      if (companyAssociations[speaker].associationScore > 0) {
        speakerScores[speaker].salesScore += companyAssociations[speaker].associationScore;

        // 自己紹介での許可リスト会社名の言及は非常に強い営業側の証拠
        if (companyAssociations[speaker].isSelfIntroduction) {
          speakerScores[speaker].salesScore += 5;
          speakerScores[speaker].confidence += 3;
        }
      }

      // 営業担当者パターンのマッチをチェック
      for (var i = 0; i < salesPatterns.length; i++) {
        if (text.match(salesPatterns[i])) {
          speakerScores[speaker].salesScore += 2; // 重み付けスコア
          speakerScores[speaker].confidence += 0.5;
        }
      }

      // お客様パターンのマッチをチェック
      for (var i = 0; i < customerPatterns.length; i++) {
        if (text.match(customerPatterns[i])) {
          speakerScores[speaker].customerScore += 2; // 重み付けスコア
          speakerScores[speaker].confidence += 0.5;
        }
      }

      // 自己紹介パターンのチェック
      for (var i = 0; i < introductionPatterns.sales.length; i++) {
        if (text.match(introductionPatterns.sales[i])) {
          speakerScores[speaker].salesScore += 2; // 自己紹介は重要な指標
          speakerScores[speaker].confidence += 0.5;
        }
      }

      for (var i = 0; i < introductionPatterns.customer.length; i++) {
        if (text.match(introductionPatterns.customer[i])) {
          speakerScores[speaker].customerScore += 2; // 自己紹介は重要な指標
          speakerScores[speaker].confidence += 0.5;
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

      // 質問の方向性の分析
      if (/いらっしゃいますか|おりますか|できますか|よろしいですか|\?/.test(text)) {
        // 質問が多い方が営業の可能性が高い
        speakerScores[speaker].salesScore += 1.5;
      }

      // 謝罪パターンの分析（営業担当は謝罪が多い）
      if ((text.match(/すみません|申し訳|失礼|ごめん/g) || []).length > 2) {
        speakerScores[speaker].salesScore += 1.5;
      }
    }

    // スコアに基づいて役割を決定
    for (var speaker in speakerScores) {
      var scores = speakerScores[speaker];

      // 許可リスト会社名との関連が非常に強い場合（自己紹介など）は、それを優先
      if (companyAssociations[speaker].isSelfIntroduction) {
        speakerRoles[speaker] = 'sales';
        continue;
      }

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

      // 信頼度も考慮
      var confidence0 = scores0.confidence;
      var confidence1 = scores1.confidence;

      if ((diff0 >= diff1 && confidence0 >= confidence1) || (confidence0 > confidence1 * 1.5)) {
        // speaker0のスコア差が大きいか信頼度が著しく高い場合はその判定を維持し、speaker1を反対の役割に
        speakerRoles[speakers[1]] = (speakerRoles[speakers[0]] === 'sales') ? 'customer' : 'sales';
      } else if ((diff1 > diff0 && confidence1 >= confidence0) || (confidence1 > confidence0 * 1.5)) {
        // speaker1のスコア差が大きいか信頼度が著しく高い場合はその判定を維持し、speaker0を反対の役割に
        speakerRoles[speakers[0]] = (speakerRoles[speakers[1]] === 'sales') ? 'customer' : 'sales';
      } else {
        // それでも判断が難しい場合は、営業パターンの一致度が高い方を営業担当に
        var salesPatternCount0 = 0;
        var salesPatternCount1 = 0;

        for (var i = 0; i < salesPatterns.length; i++) {
          if (speakerTexts[speakers[0]].match(salesPatterns[i])) salesPatternCount0++;
          if (speakerTexts[speakers[1]].match(salesPatterns[i])) salesPatternCount1++;
        }

        if (salesPatternCount0 > salesPatternCount1) {
          speakerRoles[speakers[0]] = 'sales';
          speakerRoles[speakers[1]] = 'customer';
        } else {
          speakerRoles[speakers[0]] = 'customer';
          speakerRoles[speakers[1]] = 'sales';
        }
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
      sales_company: cleanText(info.sales_company),
      sales_person: "", // vlookup関数を使うため常に空文字列を返す
      customer_company: "", // vlookup関数を使うため常に空文字列を返す
      customer_name: cleanText(info.customer_name),
      call_status1: info.call_status1,
      call_status2: info.call_status2,
      reason_for_refusal: cleanText(info.reason_for_refusal),
      reason_for_refusal_category: info.reason_for_refusal_category,
      reason_for_appointment: cleanText(info.reason_for_appointment),
      reason_for_appointment_category: info.reason_for_appointment_category,
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

  /**
   * sales_companyの値を検証する
   * @param {string} company - 抽出された会社名
   * @return {string} - 検証済みの会社名
   */
  function validateSalesCompany(company) {
    if (!company) return "";

    var normalizedCompany = String(company).trim();

    // 許可されたsales_companyのリスト
    var allowedCompanies = [
      "株式会社ENERALL",
      "エムスリーヘルスデザイン株式会社",
      "株式会社TOKIUM",
      "株式会社グッドワークス",
      "テコム看護",
      "ハローワールド株式会社",
      "株式会社ワーサル",
      "株式会社NOTCH",
      "株式会社ジースタイラス",
      "株式会社佑人社"
    ];

    // 完全一致をチェック
    if (allowedCompanies.indexOf(normalizedCompany) !== -1) {
      return normalizedCompany;
    }

    // 部分一致の会社名を確認
    for (var i = 0; i < allowedCompanies.length; i++) {
      // 会社名の特徴的な部分を抽出（例：「株式会社」を除いた部分）
      var companyKey = allowedCompanies[i].replace(/株式会社|エムスリーヘルスデザイン|テコム看護|ハローワールド/g, "").trim();

      // 特徴的な部分が含まれているかをチェック
      if (companyKey && normalizedCompany.indexOf(companyKey) !== -1) {
        return allowedCompanies[i];
      }

      // 「株式会社」の位置が異なる場合も確認
      if (allowedCompanies[i].indexOf("株式会社") === 0) {
        // 「株式会社XXX」の場合「XXX株式会社」でも一致
        var companyNameOnly = allowedCompanies[i].replace("株式会社", "").trim();
        if (normalizedCompany === companyNameOnly + "株式会社") {
          return allowedCompanies[i];
        }
      } else if (allowedCompanies[i].indexOf("株式会社") > 0) {
        // 「XXX株式会社」の場合「株式会社XXX」でも一致
        var companyNameOnly = allowedCompanies[i].replace("株式会社", "").trim();
        if (normalizedCompany === "株式会社" + companyNameOnly) {
          return allowedCompanies[i];
        }
      }
    }

    // どれにも一致しない場合は空文字を返す
    return "";
  }

  // 公開メソッド
  return {
    extract: extract,
    identifySpeakerRoles: identifySpeakerRoles,
    extractSpeakerInfoFromConversation: extractSpeakerInfoFromConversation
  };
})();