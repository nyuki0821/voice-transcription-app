/**
 * プロジェクト全体で使用される定数定義
 */
var Constants = (function () {

  // ステータス定数
  var STATUS = {
    // 基本ステータス
    SUCCESS: 'SUCCESS',
    ERROR: 'ERROR',
    PENDING: 'PENDING',

    // 処理中ステータス
    PROCESSING: 'PROCESSING',
    RETRY: 'RETRY',
    FORCE_RETRY: 'FORCE_RETRY',
    RESET_PENDING: 'RESET_PENDING',

    // 検知・復旧ステータス
    ERROR_DETECTED: 'ERROR_DETECTED',
    INTERRUPTED: 'INTERRUPTED'
  };

  // リトライマーク定数
  var RETRY_MARKS = {
    RETRIED: '[RETRIED]',
    FORCE_RETRY: '[FORCE_RETRY]',
    RESET_RETRY: '[RESET_RETRY]',
    COPY_RECOVERED: '[COPY_RECOVERED]'
  };

  // エラーパターン定数
  var ERROR_PATTERNS = [
    { pattern: 'insufficient_quota', issue: 'OpenAI APIクォータ制限' },
    { pattern: 'You exceeded your current quota', issue: 'OpenAI APIクォータ制限' },
    { pattern: '【文字起こし失敗:', issue: '文字起こし処理失敗' },
    { pattern: 'GPT-4o-mini API呼び出しエラー', issue: 'GPT-4o-mini APIエラー' },
    { pattern: 'OpenAI APIからのレスポンスエラー', issue: 'OpenAI APIレスポンスエラー' },
    { pattern: 'エラー発生：', issue: '処理エラー' },
    { pattern: '情報抽出に失敗しました', issue: '情報抽出エラー' },
    { pattern: '不明（抽出エラー）', issue: '抽出エラー' },
    { pattern: 'JSONの解析に失敗しました', issue: 'JSON解析エラー' }
  ];

  // シート名定数
  var SHEET_NAMES = {
    RECORDINGS: 'recordings',
    CALL_RECORDS: 'call_records'
  };

  // カラム名定数
  var COLUMNS = {
    RECORDINGS: {
      ID: 'id',
      STATUS_TRANSCRIPTION: 'status_transcription',
      UPDATED_AT: 'updated_at',
      STATUS_FETCH: 'status_fetch'
    },
    CALL_RECORDS: {
      RECORD_ID: 'record_id',
      TRANSCRIPTION: 'transcription'
    }
  };

  // 通知時間定数
  var NOTIFICATION_HOURS = [9, 12, 19];

  // 処理設定定数
  var PROCESSING = {
    MAX_BATCH_SIZE_DEFAULT: 10,
    TIMEOUT_MINUTES: 30,
    RETRY_ATTEMPTS: 3
  };

  // ファイル関連定数
  var FILE = {
    AUDIO_MIME_TYPES: ['audio/', 'application/octet-stream'],
    AUDIO_EXTENSIONS: ['.mp3', '.wav', '.m4a', '.aac', '.ogg', '.flac', '.wma'],
    FILENAME_PATTERN: /zoom_call_\d+_([a-f0-9]+)\.mp3/i
  };

  // 営業会社リスト
  var SALES_COMPANIES = [
    'ENERALL',
    'エムスリーヘルスデザイン',
    'TOKIUM',
    'グッドワークス',
    'テコム看護',
    'ハローワールド',
    'ワーサル',
    'NOTCH',
    'ジースタイラス',
    '佑人社',
    'リディラバ',
    'インフィニットマインド'
  ];

  // メール件名テンプレート
  var EMAIL_SUBJECTS = {
    DAILY_SUMMARY: '顧客会話自動文字起こしシステム - {date} 日次処理結果サマリー',
    REALTIME_ERROR: '【緊急】顧客会話自動文字起こしシステム - 処理エラー発生',
    MISSED_ERROR_DETECTION: '顧客会話自動文字起こしシステム - 見逃しエラー検知・復旧完了'
  };

  // エラーメッセージテンプレート
  var ERROR_MESSAGES = {
    FILE_NOT_SPECIFIED: '移動対象ファイルが指定されていません',
    FOLDER_ID_NOT_SPECIFIED: '移動先フォルダIDが指定されていません',
    SHEET_NOT_FOUND: '{sheetName}シートが見つかりません',
    REQUIRED_COLUMNS_NOT_FOUND: '必要なカラムが見つかりません',
    RECORD_NOT_FOUND: 'レコードが見つかりません',
    METADATA_EXTRACTION_FAILED: 'メタデータの抽出に失敗しました'
  };

  /**
   * エラーメッセージテンプレートに値を埋め込む
   * @param {string} template - テンプレート文字列
   * @param {Object} values - 埋め込む値のオブジェクト
   * @return {string} - 値が埋め込まれた文字列
   */
  function formatMessage(template, values) {
    var result = template;
    for (var key in values) {
      if (values.hasOwnProperty(key)) {
        result = result.replace(new RegExp('{' + key + '}', 'g'), values[key]);
      }
    }
    return result;
  }

  /**
   * ファイルが音声ファイルかどうかを判定
   * @param {File} file - 判定対象ファイル
   * @return {boolean} - 音声ファイルの場合true
   */
  function isAudioFile(file) {
    if (!file) return false;

    var mimeType = file.getMimeType() || "";
    var fileName = file.getName().toLowerCase();

    // MIMEタイプチェック
    for (var i = 0; i < FILE.AUDIO_MIME_TYPES.length; i++) {
      if (mimeType.indexOf(FILE.AUDIO_MIME_TYPES[i]) === 0) {
        return true;
      }
    }

    // 拡張子チェック
    for (var j = 0; j < FILE.AUDIO_EXTENSIONS.length; j++) {
      if (fileName.indexOf(FILE.AUDIO_EXTENSIONS[j]) !== -1) {
        return true;
      }
    }

    return false;
  }

  /**
   * ファイル名からレコードIDを抽出
   * @param {string} fileName - ファイル名
   * @return {string|null} - 抽出されたレコードID、見つからない場合はnull
   */
  function extractRecordIdFromFileName(fileName) {
    if (!fileName) return null;

    var match = fileName.match(FILE.FILENAME_PATTERN);
    return match && match[1] ? match[1] : null;
  }

  // 公開オブジェクト
  return {
    STATUS: STATUS,
    RETRY_MARKS: RETRY_MARKS,
    ERROR_PATTERNS: ERROR_PATTERNS,
    SHEET_NAMES: SHEET_NAMES,
    COLUMNS: COLUMNS,
    NOTIFICATION_HOURS: NOTIFICATION_HOURS,
    PROCESSING: PROCESSING,
    FILE: FILE,
    SALES_COMPANIES: SALES_COMPANIES,
    EMAIL_SUBJECTS: EMAIL_SUBJECTS,
    ERROR_MESSAGES: ERROR_MESSAGES,

    // ユーティリティ関数
    formatMessage: formatMessage,
    isAudioFile: isAudioFile,
    extractRecordIdFromFileName: extractRecordIdFromFileName
  };
})(); 