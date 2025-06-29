/**
 * クライアントマスターデータローダー
 * スプレッドシートからクライアント名のマスターデータを読み込む
 */
var ClientMasterDataLoader = (function () {
  var cachedData = null;
  var cacheExpiry = null;
  var CACHE_DURATION = 60 * 60 * 1000; // 1時間のキャッシュ

  /**
   * スプレッドシートからクライアントデータを読み込む
   * @return {Object} クライアントデータ
   */
  function loadClientData() {
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

      // 「クライアントマスタ」シートを取得（存在しない場合はアクティブシートを使用）
      var sheet;
      try {
        sheet = spreadsheet.getSheetByName('クライアントマスタ');
        if (!sheet) {
          sheet = spreadsheet.getActiveSheet();
          Logger.log('「クライアントマスタ」シートが見つからないため、アクティブシートを使用します');
        }
      } catch (e) {
        sheet = spreadsheet.getActiveSheet();
        Logger.log('シート取得エラー、アクティブシートを使用します: ' + e.toString());
      }

      // データ範囲を取得（ヘッダー行を除く）
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();

      if (values.length < 2) {
        throw new Error('スプレッドシートにデータがありません');
      }

      // ヘッダー行を取得してcompany_name列を探す
      var headers = values[0];
      var companyNameIndex = headers.indexOf('company_name');

      // company_name列が見つからない場合は、B列（インデックス1）を使用
      if (companyNameIndex === -1) {
        Logger.log('company_name列が見つからないため、B列（2列目）を使用します');
        companyNameIndex = 1; // B列のインデックス

        // B列が存在するかチェック
        if (headers.length < 2) {
          throw new Error('B列（company_name列）が存在しません');
        }
      }

      // クライアント名リストを作成
      var companies = [];
      var companyAliases = {}; // 会社名のエイリアス（読み方、略称など）

      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var companyName = row[companyNameIndex];

        if (companyName && companyName.toString().trim()) {
          var normalizedName = companyName.toString().trim();
          companies.push(normalizedName);

          // 会社名のバリエーションを生成
          var aliases = generateCompanyAliases(normalizedName);
          companyAliases[normalizedName] = aliases;
        }
      }

      // 重複を除去
      companies = companies.filter(function (company, index, self) {
        return self.indexOf(company) === index;
      });

      // キャッシュに保存
      cachedData = {
        companies: companies,
        companyAliases: companyAliases,
        lastUpdated: new Date()
      };
      cacheExpiry = new Date(Date.now() + CACHE_DURATION);

      Logger.log('クライアントマスターデータを読み込みました: ' + companies.length + '件');

      return cachedData;
    } catch (error) {
      Logger.log('クライアントマスターデータの読み込みエラー: ' + error.toString());
      // エラーが発生した場合は、デフォルトのデータを返す
      return getDefaultClientData();
    }
  }

  /**
   * 会社名のエイリアス（読み方、略称など）を生成
   * @param {string} companyName - 会社名
   * @return {Array} エイリアスのリスト
   */
  function generateCompanyAliases(companyName) {
    var aliases = [companyName];

    // 株式会社を除いた形
    if (companyName.indexOf('株式会社') === 0) {
      aliases.push(companyName.substring(4));
    } else if (companyName.indexOf('株式会社') > 0) {
      aliases.push(companyName.replace('株式会社', ''));
    }

    // 特定の会社名に対する読み方やバリエーション
    var specialAliases = {
      '株式会社ENERALL': ['ENERALL', 'ENERRALL', 'エネラル', 'エネラール', 'エナラル'],
      '株式会社NOTCH': ['NOTCH', 'NOCH', 'ノッチ', 'ノーチ'],
      'エムスリーヘルスデザイン株式会社': ['エムスリーヘルスデザイン', 'エムスリーヘルス', 'M3ヘルス', 'エムスリー'],
      '株式会社佑人社': ['佑人社', '有人社', '勇人社', '優人社', 'ゆうじんしゃ', 'ユウジンシャ'],
      '株式会社TOKIUM': ['TOKIUM', 'トキウム'],
      'テコム看護': ['テコム'],
      '株式会社グッドワークス': ['グッドワークス'],
      'ハローワールド株式会社': ['ハローワールド'],
      '株式会社ワーサル': ['ワーサル'],
      '株式会社ジースタイラス': ['ジースタイラス'],
      '株式会社リディラバ': ['リディラバ'],
      '株式会社インフィニットマインド': ['インフィニットマインド']
    };

    if (specialAliases[companyName]) {
      aliases = aliases.concat(specialAliases[companyName]);
    }

    // 重複を除去
    return aliases.filter(function (alias, index, self) {
      return self.indexOf(alias) === index;
    });
  }

  /**
   * デフォルトのクライアントデータを返す（フォールバック用）
   * @return {Object} デフォルトのクライアントデータ
   */
  function getDefaultClientData() {
    var defaultCompanies = [
      "株式会社ENERALL",
      "エムスリーヘルスデザイン株式会社",
      "株式会社TOKIUM",
      "株式会社グッドワークス",
      "テコム看護",
      "ハローワールド株式会社",
      "株式会社ワーサル",
      "株式会社NOTCH",
      "株式会社ジースタイラス",
      "株式会社佑人社",
      "株式会社リディラバ",
      "株式会社インフィニットマインド"
    ];

    var companyAliases = {};
    for (var i = 0; i < defaultCompanies.length; i++) {
      companyAliases[defaultCompanies[i]] = generateCompanyAliases(defaultCompanies[i]);
    }

    return {
      companies: defaultCompanies,
      companyAliases: companyAliases,
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
   * クライアント名リストを取得
   * @return {Array} クライアント名のリスト
   */
  function getClientCompanies() {
    var data = loadClientData();
    return data.companies;
  }

  /**
   * クライアント名とそのエイリアスを取得
   * @return {Object} クライアント名とエイリアスのマッピング
   */
  function getClientCompanyAliases() {
    var data = loadClientData();
    return data.companyAliases;
  }

  /**
   * LLM用のプロンプトテキストを生成
   * @return {string} クライアント名リストのプロンプト
   */
  function getClientListPrompt() {
    var companies = getClientCompanies();
    var prompt = "営業会社名は以下のリストから最も適切なものを選んでください。会話は基本的にこのリストの会社のいずれかから行われています。営業担当者が自社名を名乗っている部分を特定し、該当する会社を選択してください。名乗りがない場合や自動音声のみの非常に短い会話でない限り、必ずいずれかの会社を選択してください：\n";

    for (var i = 0; i < companies.length; i++) {
      var company = companies[i];
      var aliases = getClientCompanyAliases()[company];

      // メインの会社名
      prompt += "・" + company;

      // 主要なエイリアスを括弧内に表示（最大3つ）
      if (aliases && aliases.length > 1) {
        var displayAliases = aliases.slice(1, 4).filter(function (alias) {
          return alias !== company && alias.length > 1;
        });
        if (displayAliases.length > 0) {
          prompt += "（" + displayAliases.join("、") + "）";
        }
      }

      prompt += "\n";
    }

    return prompt;
  }

  // 公開メソッド
  return {
    loadClientData: loadClientData,
    getClientCompanies: getClientCompanies,
    getClientCompanyAliases: getClientCompanyAliases,
    getClientListPrompt: getClientListPrompt,
    clearCache: clearCache
  };
})(); 