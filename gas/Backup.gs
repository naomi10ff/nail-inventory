/**
 * スプレッドシート全体の自動バックアップ。
 * Googleスプレッドシートの変更履歴だけでは、ファイル自体の削除や大きな事故に対する
 * 備えにならないため、日付付きのコピーをGoogleドライブの別フォルダに定期保存する。
 */

var BACKUP_FOLDER_NAME = '在庫管理バックアップ';
var BACKUP_RETENTION_DAYS = 30;

/**
 * 現在のスプレッドシートを、日付付きの名前でバックアップ用フォルダにコピーする。
 * BACKUP_RETENTION_DAYSより古いバックアップは自動的に削除する(ドライブの容量対策)。
 * 通常は setupDailyBackupTrigger() で設定した日次トリガーから自動的に呼ばれるが、
 * Apps Scriptエディタから手動実行して今すぐ1回バックアップを取ることもできる。
 */
function backupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = getOrCreateBackupFolder_();
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  var name = ss.getName() + '_バックアップ_' + dateStr;
  var file = DriveApp.getFileById(ss.getId());
  file.makeCopy(name, folder);

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_RETENTION_DAYS);
  var files = folder.getFiles();
  var deletedCount = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (f.getDateCreated() < cutoff) {
      f.setTrashed(true);
      deletedCount++;
    }
  }
  Logger.log('バックアップを作成しました: ' + name + '(' + BACKUP_RETENTION_DAYS + '日より古いバックアップ ' + deletedCount + '件を削除)');
}

function getOrCreateBackupFolder_() {
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * 毎日自動でバックアップを取るトリガーを設定する。Apps Scriptエディタから手動で1回だけ
 * 実行する(初回実行時にGoogleドライブへのアクセス許可を求められるので許可する)。
 * 既に同じトリガーが設定済みの場合は何もしない(再実行しても安全)。
 */
function setupDailyBackupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'backupSpreadsheet') {
      Logger.log('既に自動バックアップのトリガーが設定されています。何もしませんでした。');
      return;
    }
  }
  ScriptApp.newTrigger('backupSpreadsheet')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('毎日深夜3時(日本時間)に自動でバックアップするよう設定しました。');
}
