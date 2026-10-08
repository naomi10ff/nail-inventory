/**
 * スプレッドシート全体の自動バックアップ。
 * Googleスプレッドシートの変更履歴だけでは、ファイル自体の削除や大きな事故に対する
 * 備えにならないため、日付付きのコピーをGoogleドライブの別フォルダに定期保存する。
 *
 * このプロジェクトはドライブの権限を drive.file(このスクリプトが作ったファイルだけ扱える)
 * に絞っている。以前は DriveApp(getFileById/getFoldersByName など、ドライブ全体の権限が
 * 必要)を使っていたため、毎回権限エラーで一度もバックアップが作られていなかった。
 * そのため、元ファイルをドライブ上でコピーするのではなく、SpreadsheetApp で新しい
 * スプレッドシートを作って全シートの中身を写し、フォルダへの移動・古いバックアップの削除は
 * Drive 拡張サービス(v3)で「このスクリプトが作ったファイル」に対してだけ行う。
 */

var BACKUP_FOLDER_NAME = '在庫管理バックアップ';
var BACKUP_RETENTION_DAYS = 30;
var BACKUP_FOLDER_ID_PROP = 'BACKUP_FOLDER_ID';

/**
 * 現在のスプレッドシートを、日付付きの名前でバックアップ用フォルダに保存する。
 * BACKUP_RETENTION_DAYSより古いバックアップは自動的に削除する(ドライブの容量対策)。
 * 通常は setupDailyBackupTrigger() で設定した日次トリガーから自動的に呼ばれるが、
 * Apps Scriptエディタから手動実行して今すぐ1回バックアップを取ることもできる。
 */
function backupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folderId = getOrCreateBackupFolderId_();
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  var name = ss.getName() + '_バックアップ_' + dateStr;

  var backup = SpreadsheetApp.create(name);
  // 新規作成時に最初からある空のシートは、元のシート名と重ならない仮の名前にしておき、最後に消す
  var placeholder = backup.getSheets()[0].setName('__backup_placeholder__');
  ss.getSheets().forEach(function (sheet) {
    var copy = sheet.copyTo(backup).setName(sheet.getName());
    // 数式が他のシートを参照していると、コピー順によっては参照が壊れるため、値で上書きして確実に残す
    var range = sheet.getDataRange();
    copy.getRange(1, 1, range.getNumRows(), range.getNumColumns()).setValues(range.getValues());
  });
  backup.deleteSheet(placeholder);

  // 作成直後はマイドライブ直下にあるので、バックアップ用フォルダへ移す
  var parents = Drive.Files.get(backup.getId(), { fields: 'parents' }).parents || [];
  Drive.Files.update({}, backup.getId(), null, { addParents: folderId, removeParents: parents.join(',') });

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_RETENTION_DAYS);
  var oldFiles = Drive.Files.list({
    q: "'" + folderId + "' in parents and trashed = false and createdTime < '" + cutoff.toISOString() + "'",
    fields: 'files(id)',
    pageSize: 100
  }).files || [];
  oldFiles.forEach(function (f) {
    Drive.Files.update({ trashed: true }, f.id);
  });
  Logger.log('バックアップを作成しました: ' + name + '(' + BACKUP_RETENTION_DAYS + '日より古いバックアップ ' + oldFiles.length + '件を削除)');
}

/**
 * バックアップ用フォルダのIDを返す。drive.file の権限では名前でフォルダを探せないため、
 * 初回に作ったフォルダのIDをスクリプトプロパティに覚えておき、消されていたら作り直す。
 */
function getOrCreateBackupFolderId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(BACKUP_FOLDER_ID_PROP);
  if (id) {
    try {
      if (!Drive.Files.get(id, { fields: 'trashed' }).trashed) return id;
    } catch (e) {
      // フォルダが削除されている・見つからない場合は作り直す
    }
  }
  var folder = Drive.Files.create({ name: BACKUP_FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder' });
  props.setProperty(BACKUP_FOLDER_ID_PROP, folder.id);
  return folder.id;
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
