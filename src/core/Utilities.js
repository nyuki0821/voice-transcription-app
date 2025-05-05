/**
 * ユーティリティモジュール
 */
var CustomUtilities = (function() {
  /**
   * UUIDを生成する
   * @return {string} - 生成されたUUID
   */
  function generateUuid() {
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
    return uuid;
  }
  
  /**
   * 日時をフォーマットする
   * @param {Date} date - 日時
   * @param {string} format - フォーマット
   * @return {string} - フォーマットされた日時文字列
   */
  function formatDate(date, format) {
    return Utilities.formatDate(date, 'Asia/Tokyo', format);
  }
  
  /**
   * 指定時間スリープする
   * @param {number} milliseconds - スリープ時間（ミリ秒）
   */
  function sleep(milliseconds) {
    Utilities.sleep(milliseconds);
  }
  
  // 公開メソッド
  return {
    generateUuid: generateUuid,
    formatDate: formatDate,
    sleep: sleep
  };
})();