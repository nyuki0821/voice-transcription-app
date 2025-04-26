/**
 * RetentionCleaner
 * Google Drive に保存している Zoom Phone 録音ファイルを保持期間を過ぎたら自動削除するユーティリティ
 *
 * スクリプトプロパティ
 *   RETENTION_DAYS    : 保持日数。未設定時は 90 日。
 *   SOURCE_FOLDER_ID  : 対象フォルダの ID
 */
function purgeOldRecordings() {
  try {
    // 保持日数を取得（デフォルト 90 日）
    var days = Number(PropertiesService.getScriptProperties().getProperty('RETENTION_DAYS') || 90);
    if (isNaN(days) || days <= 0) days = 90;

    // 対象フォルダ
    var folderId = PropertiesService.getScriptProperties().getProperty('SOURCE_FOLDER_ID');
    if (!folderId) throw new Error('SOURCE_FOLDER_ID が未設定です');

    var threshold = Date.now() - days * 24 * 60 * 60 * 1000; // ミリ秒

    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var trashed = 0;
    while (files.hasNext()) {
      var f = files.next();
      try {
        if (f.getDateCreated().getTime() < threshold) {
          f.setTrashed(true);
          trashed++;
        }
      } catch (innerErr) {
        // 例外が発生しても続行
        Logger.log('[RetentionCleaner] Skip file: ' + innerErr);
      }
    }
    Logger.log('[RetentionCleaner] Purged files: ' + trashed);
    return 'Purged files: ' + trashed;
  } catch (err) {
    Logger.log('[RetentionCleaner] Error: ' + err);
    return 'Error: ' + err;
  }
} 