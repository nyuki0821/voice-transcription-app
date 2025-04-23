/**
 * sendDailySummary関数をテストする関数
 * エラーメッセージを詳細に確認し、修正を適用してテストします
 */
function testSendDailySummary() {
  try {
    Logger.log("sendDailySummaryのテスト開始");
    
    // 元の関数を呼び出し、結果を取得
    var result = sendDailySummary();
    
    // ログにテスト結果を出力
    Logger.log("テスト結果: " + result);
    
    return "テストが完了しました: " + result;
  } catch (error) {
    // エラー発生時の詳細なログ
    Logger.log("testSendDailySummaryでエラー: " + error.toString());
    
    if (error.stack) {
      Logger.log("エラースタック: " + error.stack);
    }
    
    return "テスト中にエラーが発生しました: " + error.toString();
  }
}

/**
 * 設定の読み込み状況を詳細に診断するテスト関数
 * スプレッドシートからの設定読み込みの各ステップを検証し、
 * 管理者メールアドレスが正しく設定されているか確認します
 */
function testSettingsLoading() {
  try {
    // スプレッドシートIDの確認
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      return 'SPREADSHEET_IDが設定されていません';
    }
    
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    
    // スプレッドシートを開く
    var spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      Logger.log('スプレッドシートが正常に開けました: ' + spreadsheet.getName());
    } catch (error) {
      return 'スプレッドシートを開けませんでした: ' + error.toString();
    }
    
    // システム設定シートを取得
    var settingsSheet;
    try {
      settingsSheet = spreadsheet.getSheetByName('システム設定');
      if (!settingsSheet) {
        return 'システム設定シートが見つかりません';
      }
      Logger.log('システム設定シートが見つかりました');
    } catch (error) {
      return 'システム設定シートの取得中にエラー: ' + error.toString();
    }
    
    // 設定データを読み込む
    var settingsData;
    try {
      settingsData = settingsSheet.getDataRange().getValues();
      Logger.log('設定データの行数: ' + settingsData.length);
    } catch (error) {
      return '設定データの読み込み中にエラー: ' + error.toString();
    }
    
    // 設定の一時保存用オブジェクト
    var settings = {};
    var adminEmailKey = null;
    var adminEmailValue = null;
    
    // 設定データの解析
    Logger.log('--- 設定キーと値の一覧 ---');
    for (var i = 1; i < settingsData.length; i++) {
      var key = String(settingsData[i][0]).trim();
      var value = settingsData[i][1];
      
      if (key) {
        settings[key] = value;
        Logger.log('キー: [' + key + '] 値: [' + value + ']');
        
        // 管理者メールアドレス関連のキーを検出
        if (key.toUpperCase().indexOf('ADMIN') !== -1 && key.toUpperCase().indexOf('EMAIL') !== -1) {
          adminEmailKey = key;
          adminEmailValue = value;
        }
      }
    }
    
    // 管理者メールの処理
    if (adminEmailKey) {
      Logger.log('--- 管理者メールアドレス情報 ---');
      Logger.log('キー: ' + adminEmailKey);
      Logger.log('値: ' + adminEmailValue);
      
      if (adminEmailValue) {
        // カンマ区切り処理をテスト
        var emails = String(adminEmailValue).split(',').map(function(email) {
          return email.trim();
        }).filter(function(email) {
          return email !== '';
        });
        
        Logger.log('分割後のメールアドレス数: ' + emails.length);
        for (var i = 0; i < emails.length; i++) {
          Logger.log('メールアドレス[' + i + ']: ' + emails[i]);
        }
        
        // 最終的に設定されるADMIN_EMAILSの検証
        var result = {
          ADMIN_EMAILS: emails
        };
        Logger.log('設定される管理者メールアドレス配列の長さ: ' + result.ADMIN_EMAILS.length);
        
        return '管理者メールアドレスの設定: キー=[' + adminEmailKey + '], 値=[' + adminEmailValue + '], 処理後の配列長=[' + emails.length + ']';
      } else {
        return '管理者メールアドレスキーは存在しますが、値が空です: ' + adminEmailKey;
      }
    } else {
      // 実際にloadSettings関数を呼び出して結果を確認
      var actualSettings = loadSettings();
      Logger.log('--- loadSettings()の実行結果 ---');
      Logger.log('ADMIN_EMAILS配列: ' + JSON.stringify(actualSettings.ADMIN_EMAILS));
      Logger.log('ADMIN_EMAILS配列の長さ: ' + (actualSettings.ADMIN_EMAILS ? actualSettings.ADMIN_EMAILS.length : 0));
      
      return '管理者メールアドレスに関連するキーが見つかりません。キー名が "ADMIN_EMAIL" または類似の名前になっているか確認してください。';
    }
  } catch (error) {
    Logger.log('テスト中にエラー: ' + error.toString());
    if (error.stack) {
      Logger.log('エラースタック: ' + error.stack);
    }
    return 'テスト実行中にエラーが発生しました: ' + error.toString();
  }
}


/**
 * 最もシンプルなテスト実装 - 直接メールアドレスを指定して機能をテスト
 */
function forceTestSendDailySummary() {
  try {
    Logger.log("強制的なsendDailySummaryテスト開始");
    
    // 強制的にテストするためのカスタム関数を作成
    function testSendSummary() {
      // 当日の日付を取得
      var today = new Date();
      var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
      
      // テスト用のダミー結果を作成
      var results = {
        success: 3,
        error: 1,
        details: [
          { fileName: "テストファイル1.mp3", status: "成功", processingTime: 120 },
          { fileName: "テストファイル2.mp3", status: "成功", processingTime: 150 },
          { fileName: "テストファイル3.mp3", status: "成功", processingTime: 180 },
          { fileName: "テストファイル4.mp3", status: "エラー", processingTime: 0 }
        ],
        date: todayStr,
        totalProcessingTime: 450,
        avgProcessingTime: 150
      };
      
      // テスト用の固定メールアドレスでメールを送信（必要に応じて変更）
      var testEmail = "jatpyuki0821@gmail.com";
      
      // 実際にメールを送信（テスト目的なので1通のみ）
      sendEnhancedDailySummary(testEmail, results, todayStr);
      
      return "テスト用メールを送信しました: " + testEmail;
    }
    
    // テスト実行
    var result = testSendSummary();
    
    // ログにテスト結果を出力
    Logger.log("テスト結果: " + result);
    
    return "テストが完了しました: " + result;
  } catch (error) {
    // エラー発生時の詳細なログ
    Logger.log("テスト中にエラー: " + error.toString());
    
    if (error.stack) {
      Logger.log("エラースタック: " + error.stack);
    }
    
    return "テスト中にエラーが発生しました: " + error.toString();
  }
}