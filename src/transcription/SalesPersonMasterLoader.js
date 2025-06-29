/**
 * 担当者マスターデータローダー
 * スプレッドシートから担当者名のマスターデータを読み込む
 */
var SalesPersonMasterLoader = (function () {
  var cachedData = null;
  var cacheExpiry = null;
  var CACHE_DURATION = 60 * 60 * 1000; // 1時間のキャッシュ

  /**
   * スプレッドシートから担当者データを読み込む
   * @return {Object} 担当者データ
   */
  function loadSalesPersonData() {
    // キャッシュが有効な場合はキャッシュから返す
    if (cachedData && cacheExpiry && new Date() < cacheExpiry) {
      return cachedData;
    }

    try {
      // スクリプトプロパティからスプレッドシートIDを取得
      var scriptProperties = PropertiesService.getScriptProperties();
      var spreadsheetId = scriptProperties.getProperty('CLIENT_MASTER_SPREADSHEET_ID');

      if (!spreadsheetId) {
        throw new Error('CLIENT_MASTER_SPREADSHEET_IDがスクリプトプロパティに設定されていません');
      }

      // スプレッドシートを開く
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

      // 「担当者マスタ」シートを取得
      var sheet;
      try {
        sheet = spreadsheet.getSheetByName('担当者マスタ');
        if (!sheet) {
          throw new Error('「担当者マスタ」シートが見つかりません');
        }
      } catch (e) {
        Logger.log('担当者マスタシート取得エラー: ' + e.toString());
        return getDefaultSalesPersonData();
      }

      // データ範囲を取得（ヘッダー行を除く）
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();

      if (values.length < 2) {
        throw new Error('担当者マスタシートにデータがありません');
      }

      // ヘッダー行を取得してsales_employee_name列を探す
      var headers = values[0];
      var nameIndex = headers.indexOf('sales_employee_name');

      // sales_employee_name列が見つからない場合は、B列（インデックス1）を使用
      if (nameIndex === -1) {
        Logger.log('sales_employee_name列が見つからないため、B列（2列目）を使用します');
        nameIndex = 1; // B列のインデックス

        // B列が存在するかチェック
        if (headers.length < 2) {
          throw new Error('B列（sales_employee_name列）が存在しません');
        }
      }

      // 担当者名リストを作成
      var salesPersons = [];
      var personAliases = {}; // 担当者名のエイリアス（姓のみ、ひらがな、カタカナなど）

      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var personName = row[nameIndex];

        if (personName && personName.toString().trim()) {
          var normalizedName = personName.toString().trim();
          salesPersons.push(normalizedName);

          // 担当者名のバリエーションを生成
          var aliases = generatePersonAliases(normalizedName);
          personAliases[normalizedName] = aliases;
        }
      }

      // 重複を除去
      salesPersons = salesPersons.filter(function (person, index, self) {
        return self.indexOf(person) === index;
      });

      // キャッシュに保存
      cachedData = {
        salesPersons: salesPersons,
        personAliases: personAliases,
        lastUpdated: new Date()
      };
      cacheExpiry = new Date(Date.now() + CACHE_DURATION);

      Logger.log('担当者マスターデータを読み込みました: ' + salesPersons.length + '件');

      return cachedData;
    } catch (error) {
      Logger.log('担当者マスターデータの読み込みエラー: ' + error.toString());
      // エラーが発生した場合は、デフォルトのデータを返す
      return getDefaultSalesPersonData();
    }
  }

  /**
   * 担当者名のエイリアス（姓のみ、読み方など）を生成
   * @param {string} personName - 担当者名（フルネーム）
   * @return {Array} エイリアスのリスト
   */
  function generatePersonAliases(personName) {
    var aliases = [personName];

    // スペースで分割して姓と名を取得
    var parts = personName.split(/[\s　]+/);
    if (parts.length >= 2) {
      var lastName = parts[0];
      var firstName = parts[1];

      // 姓のみを追加
      aliases.push(lastName);

      // 姓名の間のスペースを全角・半角両方で対応
      aliases.push(lastName + ' ' + firstName);
      aliases.push(lastName + '　' + firstName);
    }

    // 特定の名前に対する読み方やバリエーション（例）
    var specialAliases = {
      '高野 仁': ['高野', 'たかの', 'タカノ', 'TAKANO', '高野さん'],
      '本間 隼': ['本間', 'ほんま', 'ホンマ', 'HONMA', '本間さん'],
      '下山 裕司': ['下山', 'しもやま', 'シモヤマ', 'SHIMOYAMA', '下山さん'],
      '桝田 明良': ['桝田', '枡田', 'ますだ', 'マスダ', 'MASUDA', '桝田さん', '枡田さん'],
      '中山 勇樹': ['中山', 'なかやま', 'ナカヤマ', 'NAKAYAMA', '中山さん'],
      '家永 麻里絵': ['家永', 'いえなが', 'イエナガ', 'IENAGA', '家永さん'],
      '齋藤 茜': ['齋藤', '斎藤', '斉藤', 'さいとう', 'サイトウ', 'SAITO', '齋藤さん', '斎藤さん'],
      '松岡 紘史': ['松岡', 'まつおか', 'マツオカ', 'MATSUOKA', '松岡さん'],
      '松田 隆治': ['松田', 'まつだ', 'マツダ', 'MATSUDA', '松田さん'],
      '横倉 佑樹': ['横倉', 'よこくら', 'ヨコクラ', 'YOKOKURA', '横倉さん'],
      '山内 藍': ['山内', 'やまうち', 'ヤマウチ', 'YAMAUCHI', '山内さん'],
      '原田 大地': ['原田', 'はらだ', 'ハラダ', 'HARADA', '原田さん'],
      '金野 祐樹': ['金野', 'きんの', 'こんの', 'キンノ', 'コンノ', 'KINNO', 'KONNO', '金野さん'],
      '荒木 美貴': ['荒木', 'あらき', 'アラキ', 'ARAKI', '荒木さん'],
      '竹内 裕里': ['竹内', 'たけうち', 'タケウチ', 'TAKEUCHI', '竹内さん'],
      '永尾 康平': ['永尾', 'ながお', 'ナガオ', 'NAGAO', '永尾さん'],
      '若山 大将': ['若山', 'わかやま', 'ワカヤマ', 'WAKAYAMA', '若山さん']
    };

    if (specialAliases[personName]) {
      aliases = aliases.concat(specialAliases[personName]);
    }

    // 重複を除去
    return aliases.filter(function (alias, index, self) {
      return self.indexOf(alias) === index;
    });
  }

  /**
   * デフォルトの担当者データを返す（フォールバック用）
   * @return {Object} デフォルトの担当者データ
   */
  function getDefaultSalesPersonData() {
    var defaultPersons = [
      "高野 仁",
      "本間 隼",
      "下山 裕司",
      "桝田 明良",
      "中山 勇樹",
      "家永 麻里絵",
      "齋藤 茜",
      "松岡 紘史"
    ];

    var personAliases = {};
    for (var i = 0; i < defaultPersons.length; i++) {
      personAliases[defaultPersons[i]] = generatePersonAliases(defaultPersons[i]);
    }

    return {
      salesPersons: defaultPersons,
      personAliases: personAliases,
      lastUpdated: new Date()
    };
  }

  /**
   * キャッシュをクリアする
   */
  function clearCache() {
    cachedData = null;
    cacheExpiry = null;
  }

  /**
   * 担当者名リストを取得
   * @return {Array} 担当者名のリスト
   */
  function getSalesPersons() {
    var data = loadSalesPersonData();
    return data.salesPersons;
  }

  /**
   * 担当者名とそのエイリアスを取得
   * @return {Object} 担当者名とエイリアスのマッピング
   */
  function getSalesPersonAliases() {
    var data = loadSalesPersonData();
    return data.personAliases;
  }

  /**
   * LLM用のプロンプトテキストを生成
   * @return {string} 担当者名リストのプロンプト
   */
  function getSalesPersonListPrompt() {
    var persons = getSalesPersons();
    var prompt = "\n\n【営業担当者名の正確な表記】\n以下は営業担当者の正しい氏名表記です。会話中で担当者名が言及される場合は、必ずこのリストの表記を使用してください：\n";

    for (var i = 0; i < persons.length; i++) {
      var person = persons[i];
      var aliases = getSalesPersonAliases()[person];

      // メインの担当者名
      prompt += "・" + person;

      // 主要なエイリアス（姓のみ、読み方）を括弧内に表示
      if (aliases && aliases.length > 1) {
        var displayAliases = aliases.slice(1, 4).filter(function (alias) {
          return alias !== person && alias.length > 0 && !alias.includes('さん');
        });
        if (displayAliases.length > 0) {
          prompt += "（" + displayAliases.join("、") + "）";
        }
      }

      prompt += "\n";
    }

    prompt += "\n※会話中で「たかの」「高野さん」などと言及された場合は「高野 仁」と正しく表記してください。\n";
    prompt += "※姓のみで呼ばれている場合でも、上記リストにある正式なフルネームで記載してください。\n";

    return prompt;
  }

  /**
   * 担当者名を正規化する（エイリアスから正式名称へ変換）
   * @param {string} name - 入力された名前（エイリアスの可能性あり）
   * @return {string} 正規化された担当者名（見つからない場合は元の名前を返す）
   */
  function normalizeSalesPersonName(name) {
    if (!name) return name;

    // 文字列以外の場合は文字列に変換
    var nameStr = typeof name === 'string' ? name : String(name);

    var data = loadSalesPersonData();
    // 複数スペースや全角スペースを正規化
    var normalizedName = nameStr.trim().replace(/[\s　]+/g, ' ');

    // 完全一致をチェック
    if (data.salesPersons.indexOf(normalizedName) !== -1) {
      return normalizedName;
    }

    // エイリアスをチェック（大文字小文字を区別しない完全一致）
    for (var fullName in data.personAliases) {
      if (data.personAliases.hasOwnProperty(fullName)) {
        var aliases = data.personAliases[fullName];
        for (var i = 0; i < aliases.length; i++) {
          if (aliases[i].toLowerCase() === normalizedName.toLowerCase()) {
            return fullName;
          }
        }
      }
    }

    // 見つからない場合は元の名前を返す
    return name;
  }

  // 公開メソッド
  return {
    loadSalesPersonData: loadSalesPersonData,
    getSalesPersons: getSalesPersons,
    getSalesPersonAliases: getSalesPersonAliases,
    getSalesPersonListPrompt: getSalesPersonListPrompt,
    normalizeSalesPersonName: normalizeSalesPersonName,
    clearCache: clearCache
  };
})(); 