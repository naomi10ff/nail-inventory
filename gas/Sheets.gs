/**
 * シート名・店舗一覧の定数と、初回セットアップ用スクリプト。
 * このファイルの setupSpreadsheet() を Apps Script エディタから一度だけ手動実行する。
 */

var SHEET_PRODUCTS = '商品マスタ';
var SHEET_STAFF = 'スタッフマスタ';
var SHEET_ACCOUNTS = 'アカウント';
var SHEET_LOG = '取引ログ';
var SHEET_SUMMARY = '現在庫サマリ';
var SHEET_SESSIONS = 'セッション';
var SHEET_BRANDS = 'ブランドマスタ';

var STORES = ['生駒店', '西大寺宝来店', '木津店'];

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // カラーNOは末尾に追加(既存行の列インデックスに影響を与えないため)
  setupSheet_(ss, SHEET_PRODUCTS, ['コード', '商品名', 'ブランド', 'カテゴリ', 'メーカー', '単位', '備考', '登録日', 'カラーNO']);
  setupSheet_(ss, SHEET_STAFF, ['店舗名', 'スタッフ名', '有効']);
  setupSheet_(ss, SHEET_ACCOUNTS, ['店舗名', 'ユーザー名', 'パスワードハッシュ', 'ソルト', '権限']);
  setupSheet_(ss, SHEET_LOG, ['タイムスタンプ', '店舗', 'スタッフ名', 'コード', '商品名', 'ブランド', '種別', '数量', 'メモ']);
  setupSheet_(ss, SHEET_SUMMARY, ['店舗', 'ブランド', '商品名', 'コード', '現在庫', '直近棚卸数', '欠品', '更新日時']);
  setupSheet_(ss, SHEET_SESSIONS, ['トークン', 'ユーザー名', '店舗', '権限', '発行日時']);
  setupSheet_(ss, SHEET_BRANDS, ['ブランド名']);

  seedAccounts_(ss);
}

function setupSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/** 既にアカウントが登録済みなら上書きしない(パスワード変更後の再実行で消えないように)。 */
function seedAccounts_(ss) {
  var sheet = ss.getSheetByName(SHEET_ACCOUNTS);
  if (sheet.getLastRow() > 1) return;

  var rows = [];
  STORES.forEach(function (store, i) {
    var salt = generateSalt_();
    var defaultPassword = 'change-me-' + (i + 1);
    rows.push([store, 'store' + (i + 1), hashPassword_(defaultPassword, salt), salt, 'store']);
  });
  var hqSalt = generateSalt_();
  rows.push(['本社', 'hq', hashPassword_('change-me-hq', hqSalt), hqSalt, 'hq']);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  Logger.log('初期アカウントを作成しました。必ずログイン後にパスワードを変更してください。');
  STORES.forEach(function (store, i) {
    Logger.log(store + ' : store' + (i + 1) + ' / change-me-' + (i + 1));
  });
  Logger.log('本社 : hq / change-me-hq');
}
