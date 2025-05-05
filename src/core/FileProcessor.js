/**
 * ファイル処理モジュール（共有ドライブ対応版）
 */
var FileProcessor = (function() {
  /**
   * 未処理のファイルを取得する
   * @param {string} sourceFolderId - 処理対象フォルダID
   * @param {number} maxFiles - 最大取得ファイル数
   * @return {Array} - 未処理ファイルの配列
   */
  function getUnprocessedFiles(sourceFolderId, maxFiles) {
    if (!sourceFolderId) {
      throw new Error('処理対象フォルダIDが指定されていません');
    }
    
    try {
      var sourceFolder = DriveApp.getFolderById(sourceFolderId);
      var files = sourceFolder.getFiles();
      var fileArray = [];
      var count = 0;
      
      while (files.hasNext() && count < maxFiles) {
        var file = files.next();
        var mimeType = file.getMimeType() || ""; 
        
        // 音声ファイルのみを対象とする
        if (mimeType.indexOf('audio/') === 0 || 
            mimeType === 'application/octet-stream') { // 一部の音声ファイルはoctet-streamとして認識される場合がある
          fileArray.push(file);
          count++;
        }
      }
      
      return fileArray;
    } catch (error) {
      throw new Error('未処理ファイルの取得中にエラー: ' + error.toString());
    }
  }
  
  /**
   * ファイルを指定フォルダに移動する（共有ドライブ対応版）
   * @param {File} file - 移動対象ファイル
   * @param {string} targetFolderId - 移動先フォルダID
   */
  function moveFileToFolder(file, targetFolderId) {
    if (!file) {
      throw new Error('移動対象ファイルが指定されていません');
    }
    
    if (!targetFolderId) {
      throw new Error('移動先フォルダIDが指定されていません');
    }
    
    try {
      var targetFolder = DriveApp.getFolderById(targetFolderId);
      
      // 共有ドライブ対応：moveTo()メソッドを使用
      file.moveTo(targetFolder);
    } catch (error) {
      throw new Error('ファイルの移動中にエラー: ' + error.toString());
    }
  }
  
  /**
   * ファイルを指定フォルダにコピーする（代替手段としてのコピー処理）
   * @param {File} file - コピー対象ファイル
   * @param {string} targetFolderId - コピー先フォルダID
   * @return {File} - コピーされたファイル
   */
  function copyFileToFolder(file, targetFolderId) {
    if (!file) {
      throw new Error('コピー対象ファイルが指定されていません');
    }
    
    if (!targetFolderId) {
      throw new Error('コピー先フォルダIDが指定されていません');
    }
    
    try {
      var targetFolder = DriveApp.getFolderById(targetFolderId);
      var fileName = file.getName();
      var fileBlob = file.getBlob();
      
      // ファイルをコピー先フォルダにコピー
      var copiedFile = targetFolder.createFile(fileBlob);
      copiedFile.setName(fileName);
      
      return copiedFile;
    } catch (error) {
      throw new Error('ファイルのコピー中にエラー: ' + error.toString());
    }
  }
  
  // 公開メソッド
  return {
    getUnprocessedFiles: getUnprocessedFiles,
    moveFileToFolder: moveFileToFolder,
    copyFileToFolder: copyFileToFolder
  };
})();