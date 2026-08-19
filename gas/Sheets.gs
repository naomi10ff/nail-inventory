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
var SHEET_APPROVALS = '棚卸承認';
var SHEET_CATEGORIES = 'カテゴリマスタ';

var STORES = ['生駒店', '西大寺宝来店', '木津店'];

// 商品マスタ・ブランドマスタの列定義(店舗ごとに独立管理するため「店舗」列を先頭に持つ)
// 「コード」はスキャン用のバーコード(または自動発行QR)、「品番」は発注時にメーカー/
// 問屋に伝える商品コードで、両者は別物(同じ商品でもバーコードが読めない・付いていない
// ことがあるが、品番は仕入れ先のカタログ上で常に存在する)。
var PRODUCTS_HEADERS = ['店舗', 'コード', '商品名', 'ブランド', 'カテゴリ', 'メーカー', '単位', '備考', '登録日', 'カラーNO', '品番'];
var BRANDS_HEADERS = ['店舗', 'ブランド名'];
var CATEGORIES_HEADERS = ['カテゴリー名'];
// カテゴリーはブランドと違い店舗ごとではなく全店舗共通の分類なので、店舗列を持たない。
var DEFAULT_CATEGORIES = ['ベース/トップ', 'カラー', 'その他', '物販'];

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSheet_(ss, SHEET_PRODUCTS, PRODUCTS_HEADERS);
  setupSheet_(ss, SHEET_STAFF, ['店舗名', 'スタッフ名', '有効']);
  setupSheet_(ss, SHEET_ACCOUNTS, ['店舗名', 'ユーザー名', 'パスワードハッシュ', 'ソルト', '権限']);
  setupSheet_(ss, SHEET_LOG, ['タイムスタンプ', '店舗', 'スタッフ名', 'コード', '商品名', 'ブランド', '種別', '数量', 'メモ']);
  setupSheet_(ss, SHEET_SUMMARY, ['店舗', 'ブランド', '商品名', 'コード', '現在庫', '直近棚卸数', '欠品', '更新日時']);
  setupSheet_(ss, SHEET_SESSIONS, ['トークン', 'ユーザー名', '店舗', '権限', '発行日時']);
  setupSheet_(ss, SHEET_BRANDS, BRANDS_HEADERS);
  setupSheet_(ss, SHEET_APPROVALS, ['店舗', '年月', '承認者', '承認日時']);
  setupSheet_(ss, SHEET_CATEGORIES, CATEGORIES_HEADERS);
  seedCategories_(ss);
  addItemNumberColumn_(ss);

  seedAccounts_(ss);
}

/**
 * 既に運用中の商品マスタ(「品番」列が無い状態)に、非破壊で列を追加する。
 * 既存データは一切変更せず、ヘッダーに「品番」が無ければ最後尾に列を追加するだけ。
 */
function addItemNumberColumn_(ss) {
  var sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheet || sheet.getLastRow() === 0) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('品番') !== -1) return;
  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue('品番').setFontWeight('bold');
}

/** カテゴリマスタが空(新規シート)なら、既定のカテゴリーを入れておく。 */
function seedCategories_(ss) {
  var sheet = ss.getSheetByName(SHEET_CATEGORIES);
  if (sheet.getLastRow() > 1) return;
  var rows = DEFAULT_CATEGORIES.map(function (name) { return [name]; });
  sheet.getRange(2, 1, rows.length, 1).setValues(rows);
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

/**
 * 商品マスタを「店舗ごとの独立管理」に切り替えるための一回限りの移行スクリプト。
 * 商品マスタ・ブランドマスタ・取引ログ・現在庫サマリの中身(テストデータ)を消去し、
 * 新しい列構成(店舗列つき)のヘッダーを書き直す。Apps Scriptエディタから手動で1回だけ実行する。
 * アカウント・スタッフマスタ・セッションは消さない。
 */
function migrateToPerStoreMasters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  resetSheetContent_(ss, SHEET_PRODUCTS, PRODUCTS_HEADERS);
  resetSheetContent_(ss, SHEET_BRANDS, BRANDS_HEADERS);
  resetSheetContent_(ss, SHEET_LOG, ['タイムスタンプ', '店舗', 'スタッフ名', 'コード', '商品名', 'ブランド', '種別', '数量', 'メモ']);
  resetSheetContent_(ss, SHEET_SUMMARY, ['店舗', 'ブランド', '商品名', 'コード', '現在庫', '直近棚卸数', '欠品', '更新日時']);
  Logger.log('商品マスタ・ブランドマスタ・取引ログ・現在庫サマリをリセットしました(店舗ごとの管理に対応した新しい列構成)。');
}

function resetSheetContent_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
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

/**
 * 木津店の既存スプレッドシート在庫データ(品名・ブランドのみ、実物のバーコードは
 * 含まない)を一括登録する一回限りの移行スクリプト。コードは空欄で自動発行される
 * (QR用の仮コード。あとで実物のバーコードが分かったら、商品マスタの編集画面で
 * コードを差し替える)。同じ店舗+ブランド+品名の組み合わせが既に存在する行は
 * スキップするので、再実行しても重複登録にはならない。
 * Apps Scriptエディタから手動で1回実行する。
 */
function bulkImportKizuProducts() {
  var store = '木津店';
  var items = [
    { brand: 'CHRISTRIO', name: 'CHRISTRIO クリアジェル LED/UV' },
    { brand: 'TRINA', name: 'ノンワイプトップジェルライナー' },
    { brand: 'TRINA', name: 'TRINA ノンワイプマットトップジェル' },
    { brand: 'para gel', name: 'para gel スーパースカルプジェル' },
    { brand: 'para gel', name: 'para gel クリアジェルEX' },
    { brand: 'para gel', name: 'para gel スカルプジェル クリア' },
    { brand: 'PREGEL', name: 'PREGEL マットコートR' },
    { brand: 'PREGEL', name: 'PREGEL ノンワイプクリア キャンジェル ノンヒート' },
    { brand: 'TOY\'s ×INITY', name: 'TOY\'s × INITY ノンワイプマットコート' },
    { brand: 'TOY\'s ×INITY', name: 'TOY\'s × INITY ノンワイプツヤトップコート' },
    { brand: 'MELTY GEL', name: 'MELTY GEL クリアジェル' },
    { brand: 'ageha', name: 'グラデーションクリアジェル' },
    { brand: 'ageha', name: 'ageha チャームオンノンワイプジェル' },
    { brand: 'ネイルパフェ', name: 'ネイルパフェ ビジューノンワイプトップ 10g' },
    { brand: 'ICEGEL', name: '1511 Dove' },
    { brand: 'ICEGEL', name: '1512 Cocoa' },
    { brand: 'ICEGEL', name: '1513 Puce' },
    { brand: 'ICEGEL', name: '1515 Flint' },
    { brand: 'ICEGEL', name: '1516 Wine' },
    { brand: 'ICEGEL', name: '1517 Berry' },
    { brand: 'ICEGEL', name: '1518 Ferra' },
    { brand: 'ICEGEL', name: 'ICE GEL A BLACK フレスコジェル S69 ホワイト' },
    { brand: 'Naility', name: 'Naility! ソリッドジェルクリア' },
    { brand: 'Naility', name: 'ソリッドジェル ホワイト' },
    { brand: 'TOY\'s ×INITY', name: 'ノンワイプアートクリアロー' },
    { brand: 'TOY\'s ×INITY', name: 'ノンワイプアートクリアハイ' },
    { brand: 'VETRO', name: 'アートクリア ミズアメ' },
    { brand: 'VETRO', name: 'アートクリア ハチミツ' },
    { brand: 'VETRO', name: 'アートクリア ワタガシ' },
    { brand: 'T-GEL', name: 'ホワイトアートジェル' },
    { brand: 'VETRO', name: '02 シャクヤク' },
    { brand: 'VETRO', name: '06 シャンパンピンク' },
    { brand: 'VETRO', name: '07 キルタンサス' },
    { brand: 'VETRO', name: '08 ロイヤルミルクティー' },
    { brand: 'VETRO', name: '19 ナイト' },
    { brand: 'VETRO', name: '22 ブラック' },
    { brand: 'VETRO', name: '23 マットホワイト' },
    { brand: 'VETRO', name: '26 オレンジ' },
    { brand: 'VETRO', name: '34 キャメロングリーン' },
    { brand: 'VETRO', name: '40 クラシックピンク' },
    { brand: 'VETRO', name: '52 ピーチヌード' },
    { brand: 'VETRO', name: '55 ホワイト' },
    { brand: 'VETRO', name: '56 シフォンローブ' },
    { brand: 'VETRO', name: '59 パーティライフ' },
    { brand: 'VETRO', name: '63 フラワーブーケ' },
    { brand: 'VETRO', name: '67 リップス' },
    { brand: 'VETRO', name: '70 ガールズトーク' },
    { brand: 'VETRO', name: '71 マゼンタ' },
    { brand: 'VETRO', name: '80 キャサリンズフェイバリット' },
    { brand: 'VETRO', name: '81 デボラグレージュ' },
    { brand: 'VETRO', name: '85 チャップリンスティック' },
    { brand: 'VETRO', name: '86 ジェームズグリーン' },
    { brand: 'VETRO', name: '88 リタレッド' },
    { brand: 'VETRO', name: '91 オードリーパール' },
    { brand: 'VETRO', name: '101 シャーベットブルー' },
    { brand: 'VETRO', name: '109 ウィステリア' },
    { brand: 'VETRO', name: '127 シーグリーン' },
    { brand: 'VETRO', name: '130 ビリヤードグリーン' },
    { brand: 'VETRO', name: '132 チョコレート' },
    { brand: 'VETRO', name: '133 ガーネット' },
    { brand: 'VETRO', name: '134 マホガニー' },
    { brand: 'VETRO', name: '135 プルシャンブルー' },
    { brand: 'VETRO', name: '142 エナメルサンセット' },
    { brand: 'VETRO', name: '188 スモークテラコッタ' },
    { brand: 'VETRO', name: '197 スモークグリーン' },
    { brand: 'VETRO', name: '199 ガール' },
    { brand: 'VETRO', name: '228 ミステリアスターコイズ' },
    { brand: 'VETRO', name: '231 グレイッシュグレージュ' },
    { brand: 'VETRO', name: '232 グレイッシュローズ' },
    { brand: 'VETRO', name: '233 グレイッシュラベンダー' },
    { brand: 'VETRO', name: '234 グレイッシュブルー' },
    { brand: 'VETRO', name: '259 ジュエルガーネット' },
    { brand: 'VETRO', name: '260 ジュエルトパーズ' },
    { brand: 'VETRO', name: '261 ジュエルアンバー' },
    { brand: 'VETRO', name: '266 ドレスショコラ' },
    { brand: 'VETRO', name: '267 ドレスノワール' },
    { brand: 'VETRO', name: '268 コバルトブルー' },
    { brand: 'VETRO', name: '269 シアンブルー' },
    { brand: 'VETRO', name: '272 ゴールドリーフ' },
    { brand: 'VETRO', name: '277 ポッパーピンク' },
    { brand: 'VETRO', name: '278 ポッパーマゼンタ' },
    { brand: 'VETRO', name: '284 スタジオNo.284' },
    { brand: 'VETRO', name: '285 スタジオNo.285' },
    { brand: 'VETRO', name: '289 ピグメントブラック' },
    { brand: 'VETRO', name: '290 ピグメントグリーン' },
    { brand: 'VETRO', name: '293 ピグメントオレンジ' },
    { brand: 'VETRO', name: '294 ピグメントイエロー' },
    { brand: 'VETRO', name: '296 カモフラカーキ' },
    { brand: 'VETRO', name: '305 エレクトリックレッド' },
    { brand: 'VETRO', name: '307 エレクトリックオレンジ' },
    { brand: 'VETRO', name: '310 ウォームオレンジ' },
    { brand: 'VETRO', name: '322 ラゲージ' },
    { brand: 'VETRO', name: '329 カシス・チョコラータ' },
    { brand: 'VETRO', name: '331 ディープスパ' },
    { brand: 'VETRO', name: '332 マンダリンオイル' },
    { brand: 'VETRO', name: '338 ストリッパー' },
    { brand: 'VETRO', name: '339 フェロモン' },
    { brand: 'VETRO', name: '343 イットガール' },
    { brand: 'VETRO', name: '344 オーロラシャワー' },
    { brand: 'VETRO', name: '346 アポロ' },
    { brand: 'VETRO', name: '352 クラッカーリッチ' },
    { brand: 'VETRO', name: '363 シーサイドテラス' },
    { brand: 'VETRO', name: '367 リネンソファ' },
    { brand: 'VETRO', name: '368 シオン' },
    { brand: 'VETRO', name: '372 チトセ' },
    { brand: 'VETRO', name: '375 レインミスト' },
    { brand: 'VETRO', name: '380 グラム' },
    { brand: 'VETRO', name: '381 サーフミント' },
    { brand: 'VETRO', name: '382 ライムリッチ' },
    { brand: 'VETRO', name: '386 ピオニー' },
    { brand: 'VETRO', name: '387 デイジー' },
    { brand: 'VETRO', name: '391 マネキンヌード' },
    { brand: 'VETRO', name: '394 グレンブルー' },
    { brand: 'VETRO', name: '396 ハネズ' },
    { brand: 'VETRO', name: '397 チョウシュンイロ' },
    { brand: 'VETRO', name: '398 サクラネズ' },
    { brand: 'VETRO', name: '401 シラチャ' },
    { brand: 'VETRO', name: '405 ビーラブド' },
    { brand: 'VETRO', name: '407 ナチュラル' },
    { brand: 'VETRO', name: '408 ビューティー' },
    { brand: 'VETRO', name: '409 ビクトリア' },
    { brand: 'VETRO', name: '411 シンプル' },
    { brand: 'VETRO', name: '425 ラプソディー' },
    { brand: 'VETRO', name: '434 ジュミニ' },
    { brand: 'VETRO', name: '456 カモミールオイル' },
    { brand: 'VETRO', name: '457 ハスカップオイル' },
    { brand: 'VETRO', name: '468 モモトセソウ' },
    { brand: 'VETRO', name: '482 カジュアルブルー' },
    { brand: 'VETRO', name: '515 アラベスク' },
    { brand: 'VETRO', name: '516 エレジー' },
    { brand: 'VETRO', name: '517 ワルツ' },
    { brand: 'VETRO', name: '518 カノン' },
    { brand: 'VETRO', name: '519 ソナタ' },
    { brand: 'VETRO', name: '520 ノクターン' },
    { brand: 'VETRO', name: '552' },
    { brand: 'VETRO', name: '553' },
    { brand: 'VETRO', name: '554' },
    { brand: 'VETRO', name: '975 トールホワイト' },
    { brand: 'VETRO', name: '976 メタシルバー' },
    { brand: 'VETRO', name: '977 メタゴールド' },
    { brand: 'VETRO', name: 'G65 プライスレス' },
    { brand: 'VETRO', name: 'G66 ドレスアップ' },
    { brand: 'VETRO', name: 'G75 オーガンジードレス' },
    { brand: 'VETRO', name: 'G77 シャンデリア' },
    { brand: 'VETRO', name: 'G119 ゴールデンロッド' },
    { brand: 'VETRO', name: 'G120 ライトシルバー' },
    { brand: 'VETRO×Bella nail', name: '002 シルバーシャインリーフ' },
    { brand: 'VETRO×Bella nail', name: '003 サンリーフ' },
    { brand: 'VETRO×Bella nail', name: '023 オレンジネオン' },
    { brand: 'VETRO×Bella nail', name: '034 アーティチョーク' },
    { brand: 'VETRO×Bella nail', name: '037 ホワイトジャック' },
    { brand: 'VETRO×Bella nail', name: '063 リフレクトイエロー' },
    { brand: 'VETRO×Bella nail', name: '064 リフレクトブルー' },
    { brand: 'VETRO×Bella nail', name: '067 リフレクトパープル' },
    { brand: 'VETRO×takiko', name: '2104 ロードライトガーネット' },
    { brand: 'VETRO×takiko', name: '2105 バイオレットサファイア' },
    { brand: 'VETRO×takiko', name: '2117 イエローカルサイト' },
    { brand: 'VETRO×takiko', name: '2120 ピンクスピネル' },
    { brand: 'VETRO×takiko', name: '2125 ソーダ' },
    { brand: 'BELLA FORMA', name: '033 サーモンピンク' },
    { brand: 'BELLA FORMA', name: '086 ピンクオークル' },
    { brand: 'BELLA FORMA', name: '153 ブルーキュラソー' },
    { brand: 'BELLA FORMA', name: '155 プリンストン' },
    { brand: 'BELLA FORMA', name: '225 ジュエリーワイン' },
    { brand: 'BELLA FORMA', name: '226 ジュエリーピンク' },
    { brand: 'BELLA FORMA', name: '277' },
    { brand: 'BELLA FORMA', name: '237 ジュエリーパープル' },
    { brand: 'BELLA FORMA', name: '286 ハンナリフジ' },
    { brand: 'BELLA FORMA', name: '289 レイヤードテラコッタ' },
    { brand: 'BELLA FORMA', name: '308 ココステラ' },
    { brand: 'BELLA FORMA', name: '309 ココグラビティ' },
    { brand: 'BELLA FORMA', name: '310 ココサテライト' },
    { brand: 'BELLA FORMA', name: '311 ココアース' },
    { brand: 'BELLA FORMA', name: '001 グラホワイト' },
    { brand: 'BELLA FORMA', name: '038 ピーチ' },
    { brand: 'BELLA FORMA', name: '085 ピンクオークル' },
    { brand: 'BELLA FORMA', name: '106 シルクシャンタン' },
    { brand: 'BELLA FORMA', name: '136 タンコウ' },
    { brand: 'BELLA FORMA', name: '138 ソライロ' },
    { brand: 'BELLA FORMA', name: '139 ウスナデシコ' },
    { brand: 'BELLA FORMA', name: '151 レモネードビタミン' },
    { brand: 'BELLA FORMA', name: '154 バイオレットフィズ' },
    { brand: 'BELLA FORMA', name: '157 オレンジピール' },
    { brand: 'BELLA FORMA', name: '236 ココビーナス' },
    { brand: 'BELLA FORMA', name: '241 ミエル' },
    { brand: 'BELLA FORMA', name: '259 ブルゴーニュ' },
    { brand: 'BELLA FORMA', name: '293 レイヤードアッシュベージュ' },
    { brand: 'BELLA FORMA', name: '000 アートホワイト' },
    { brand: 'mikinail', name: 'テンダー' },
    { brand: 'T-GEL', name: 'Sseries B006' },
    { brand: 'T-GEL', name: 'Sseries P007' },
    { brand: 'T-GEL', name: 'Sseries P009' },
    { brand: 'T-GEL', name: 'Sseries P015' },
    { brand: 'T-GEL', name: 'D011 チェリーブロッサム' },
    { brand: 'T-GEL', name: 'D015 アンティークパール' },
    { brand: 'T-GEL', name: 'D016 ホワイトアラバスター' },
    { brand: 'T-GEL', name: 'D020 シャンパンゴールド' },
    { brand: 'T-GEL', name: 'D025 サテンピンク' },
    { brand: 'T-GEL', name: 'D029 グレージュ' },
    { brand: 'T-GEL', name: 'D030 ミルキーグレージュ' },
    { brand: 'T-GEL', name: 'D059 エナメルブルー' },
    { brand: 'T-GEL', name: 'D063 アイボリー' },
    { brand: 'T-GEL', name: 'D065 シルクホワイト' },
    { brand: 'T-GEL', name: 'D067 ミディアムピンク' },
    { brand: 'T-GEL', name: 'D072 アッシュカーキ' },
    { brand: 'T-GEL', name: 'D098 キャメル' },
    { brand: 'T-GEL', name: 'D107 レッドボルドー' },
    { brand: 'T-GEL', name: 'D110 ピンクベージュ' },
    { brand: 'T-GEL', name: 'D134 シャンパンホワイト' },
    { brand: 'T-GEL', name: 'D135 レインボーホワイト' },
    { brand: 'T-GEL', name: 'D171 ココアブラウン' },
    { brand: 'T-GEL', name: 'D173 ブルーグレー' },
    { brand: 'T-GEL', name: 'D176 スケルトンダークレッド' },
    { brand: 'T-GEL', name: 'D177 スケルトンブルー' },
    { brand: 'T-GEL', name: 'D179 ヌーディアイボリー' },
    { brand: 'T-GEL', name: 'D194 パーフェクトレッド' },
    { brand: 'T-GEL', name: 'D197 グラデーションピンク' },
    { brand: 'T-GEL', name: 'D200 オリーブマスタード' },
    { brand: 'T-GEL', name: 'D204 ボルドー' },
    { brand: 'T-GEL', name: 'D210 ペールブルー' },
    { brand: 'T-GEL', name: 'D221 モカグレージュ' },
    { brand: 'T-GEL', name: 'D222 ディープブラウン' },
    { brand: 'T-GEL', name: 'D229 シフォンパープル' },
    { brand: 'T-GEL', name: 'D230 シフォンピンク' },
    { brand: 'T-GEL', name: 'D235 クリアダークオレンジ' },
    { brand: 'T-GEL', name: 'D236 クリアブルーグリーン' },
    { brand: 'T-GEL', name: 'D240 シフォンキャメル' },
    { brand: 'T-GEL', name: 'D245 シアーピンクベージュ' },
    { brand: 'T-GEL', name: 'D246 スキニーピンクベージュ' },
    { brand: 'T-GEL', name: 'D248 スキンベージュ' },
    { brand: 'T-GEL', name: 'D249 スキニーグレージュ' },
    { brand: 'T-GEL', name: 'D250 スモーキーピンク' },
    { brand: 'T-GEL', name: 'D251 ミスティローズ' },
    { brand: 'T-GEL', name: 'D253 ミストシャインベージュ' },
    { brand: 'T-GEL', name: 'D254 ミストシャインピンク' },
    { brand: 'T-GEL', name: 'D255 ミストシャインローズ' },
    { brand: 'T-GEL', name: 'D256 ミストシャイングレージュ' },
    { brand: 'T-GEL', name: 'D263 スモーキーレモン' },
    { brand: 'T-GEL', name: 'MG002 マグネット アッシュゴールド' },
    { brand: 'T-GEL', name: 'MG003 マグネット ブラウン' },
    { brand: 'mao gel', name: 'PEEK A BOO 01' },
    { brand: 'mao gel', name: 'PEEK A BOO 02' },
    { brand: 'mao gel', name: 'PEEK A BOO 03' },
    { brand: 'mao gel', name: 'PEEK A BOO 04' },
    { brand: 'mao gel', name: '601' },
    { brand: 'mao gel', name: '603' },
    { brand: 'mao gel', name: '604' },
    { brand: 'mao gel', name: '605 jastwo' },
    { brand: 'mao gel', name: '609 cheek' },
    { brand: 'mao gel', name: '615' },
    { brand: 'TRINA', name: 'CLー13' },
    { brand: 'TRINA', name: 'CLー17' },
    { brand: 'enoi', name: 'mg118' },
    { brand: 'enoi', name: 'mg128' },
    { brand: 'enoi', name: 'mg129' },
    { brand: 'enoi', name: 'mg131' },
    { brand: 'enoi', name: 'mg132' },
    { brand: 'enoi', name: 'mg136' },
    { brand: 'enoi', name: 'mg138' },
    { brand: 'enoi', name: 'mg139' },
    { brand: 'enoi', name: 'mg144' },
    { brand: 'enoi', name: 'mg148' },
    { brand: 'enoi', name: 'mg160' },
    { brand: 'enoi', name: 'mg161' },
    { brand: 'enoi', name: 'mg162' },
    { brand: 'enoi', name: 'mg164' },
    { brand: 'enoi', name: 'mg168' },
    { brand: 'enoi', name: 'mg170' },
    { brand: 'enoi', name: 'mg173' },
    { brand: 'enoi', name: 'mg179' },
    { brand: 'enoi', name: 'mg185' },
    { brand: 'enoi', name: 'gt03' },
    { brand: 'enoi', name: 'gt05' },
    { brand: 'enoi', name: 'gt08' },
    { brand: 'enoi', name: 'gt13' },
    { brand: 'enoi', name: 'f25' },
    { brand: 'Prem Doll', name: 'DOLL-B58 いじわるチェシャ猫' },
    { brand: 'Prem Doll', name: 'DOLL-B60 なみだの池' },
    { brand: 'Prem Doll', name: 'DOLL-B64 ブラックダイヤ' },
    { brand: 'Prem Doll', name: 'DOLL-B67 モスクの神秘' },
    { brand: 'Prem Doll', name: 'DOLL-721 ファーボルドー' },
    { brand: 'Prem Doll', name: 'DOLL-722 ファーキャラメル' },
    { brand: 'Prem Doll Muse', name: 'スパークルシルバー PDM-G413' },
    { brand: 'Prem Doll Muse', name: 'ピンクシャンパン PDM-G428' },
    { brand: 'Prem Doll Muse', name: 'ボジョレーワイン PDM-M463' },
    { brand: 'Prem Doll Muse', name: 'フルーツワイン PDU-S591' },
    { brand: 'PREGEL Muse', name: 'M095 ほうじ茶ラテ' },
    { brand: 'PREGEL Muse', name: 'S210 クリアハニー' },
    { brand: 'PREGEL Muse', name: 'S211 クリアブラウン' },
    { brand: 'PRESTO', name: 'PRESTO カラージェル 259' },
    { brand: 'PRESTO', name: 'PRESTO カラージェル 261' },
    { brand: 'PRESTO', name: 'PRESTO カラージェル 265' },
    { brand: 'PRESTO', name: 'PRESTO カラージェル 325' },
    { brand: 'PRESTO', name: 'PRESTO カラージェル 328' },
    { brand: 'ageha', name: '133 アラン G・MIX' },
    { brand: 'ageha', name: '253 ローズブラウン' },
    { brand: 'ageha', name: '256 フォグブルー' },
    { brand: 'ageha', name: 'ホワイトグラデーション' },
    { brand: 'TOY\'s×INITY', name: '極みべっ甲 T-KBE' },
    { brand: 'TOY\'s×INITY', name: 'T-SGP01 brown' },
    { brand: 'TOY\'s×INITY', name: 'T-SGP03 green' },
    { brand: 'TOY\'s×INITY', name: 'T-SGJ01 gold' },
    { brand: 'Choco Chip Gel', name: 'CCー01 Vanila' },
    { brand: 'Choco Chip Gel', name: 'CC-03 Blood Orange' },
    { brand: 'Choco Chip Gel', name: 'CC-06 Cafe Au lait' },
    { brand: 'lem.', name: 's425 アンズ LM-S425' },
    { brand: 'STORYJEL365', name: 'CL11' },
    { brand: 'VETRO', name: '094 Golden Star' },
    { brand: 'VETRO', name: '095 Rosy Glimmer' },
    { brand: 'VETRO', name: '096 Frosted Lavender' },
    { brand: 'VETRO', name: '097 CAST LIGHT' },
    { brand: 'VETRO', name: '098 LAB LIGHT' },
    { brand: 'VETRO', name: '099 URU:white' },
    { brand: 'VETRO', name: '100 URU:paleblue' },
    { brand: 'Bella nail', name: '2901 ブリエシルバー' },
    { brand: 'Bella nail', name: '2902 ブリエゴールド' },
    { brand: 'Bella nail', name: '2903 ブリエグリーン' },
    { brand: 'Bella nail', name: '2904 ブリエブルー' },
    { brand: 'Bella nail', name: '2905 ブリエピンク' },
    { brand: 'Bella nail', name: '2906 ブリエオレンジ' },
    { brand: 'Bella nail', name: '2907 ブリエイエロー' },
    { brand: 'Bella nail', name: '2910 ブリエブラック' },
    { brand: 'Bella nail', name: '2911 ブリエフューシャ' },
    { brand: 'Bella nail', name: '025 メロンパンチ' },
    { brand: 'Bella nail', name: '026 モアーイエロー' },
    { brand: 'Bella nail', name: '039 スキャンダルフラッシュ' },
    { brand: 'Bella nail', name: 'BL5001 スパークリングサンド' },
    { brand: 'Bella nail', name: 'BL5002  フェアリーブリンク' },
    { brand: 'Bella nail', name: 'BL5003 エンチャントヴァイオレット' },
    { brand: 'Bella nail', name: 'BL5004 マジカルブルー' },
    { brand: 'BellaForma', name: '009 Dazzle pink' },
    { brand: 'BellaForma', name: '010 Dazzle skin' },
    { brand: 'BellaForma', name: '011 Dazzle rose' },
    { brand: 'BellaForma', name: '013 Dazzle red' },
    { brand: 'BellaForma', name: '014 Dazzle frost' },
    { brand: 'GELGRAPH', name: '225G' },
    { brand: 'ICEGEL', name: '1206' },
    { brand: 'ICEGEL', name: '1467 Raspberry' },
    { brand: 'ICEGEL', name: '1468 Pinki' },
    { brand: 'ICEGEL', name: '1469 Rose' },
    { brand: 'ICEGEL', name: '1470 Mist' },
    { brand: 'ICEGEL', name: '1471 Steel Blue' },
    { brand: 'ICEGEL', name: '1472 Camo' },
    { brand: 'ICEGEL', name: '1541 Coral' },
    { brand: 'ICEGEL', name: '1542 Bianca' },
    { brand: 'ICEGEL', name: '1543 Mabel' },
    { brand: 'ICEGEL', name: '1544 Orchid' },
    { brand: 'ICEGEL', name: '1545 Onyx' },
    { brand: 'my&bee', name: 'REー002G' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 01' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 02' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 03' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 04' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 05' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 06' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 16' },
    { brand: 'CLETO', name: 'CLETO 【選カ】ヘイローマグ 06' },
    { brand: 'CLETO', name: 'グローマグ07' },
    { brand: 'CLETO', name: 'グローマグ08' },
    { brand: 'Irie', name: 'CF-02 ココア' },
    { brand: 'Irie', name: 'CF-05 チョコレート' },
    { brand: 'Irie', name: 'CF-06 ダーク' },
    { brand: 'Irie', name: 'アンリ IR-RFM01' },
    { brand: 'Irie', name: 'MN-03 ドリーム' },
    { brand: 'Irie', name: 'MN-04 スリープ' },
    { brand: 'M・GEL', name: 'C791 Wine Boa' },
    { brand: 'M・GEL', name: '792' },
    { brand: 'M・GEL', name: 'C793 Pastel Mix' },
    { brand: 'TRINA', name: 'TRINA ノンワイプマットトップジェル ライナー' },
    { brand: 'emena', name: 'フラッシュジェル0013' },
    { brand: 'emena', name: 'フラッシュジェル1011' },
    { brand: 'emena', name: 'フラッシュジェル1012' },
    { brand: 'emena', name: 'フラッシュジェル1013' },
    { brand: 'emena', name: 'フラッシュジェル1023' },
    { brand: 'emena', name: 'フラッシュジェル1025' },
    { brand: 'emena', name: 'フラッシュジェル1036' },
    { brand: 'emena', name: 'マグネティジェル0551' },
    { brand: 'emena', name: 'マグネティジェル0552' },
    { brand: 'emena', name: 'マグネティジェル0553' },
    { brand: 'emena', name: 'マグネティジェル0554' },
    { brand: 'emena', name: 'マグネティフラッシュジェル01' },
    { brand: 'emena', name: 'マグネティフラッシュジェル02' },
    { brand: 'emena', name: 'マグネティフラッシュジェル03' },
    { brand: 'emena', name: 'マグネティフラッシュジェル04' },
    { brand: 'emena', name: 'マグネティフラッシュジェル05' },
    { brand: 'emena', name: 'マグネティフラッシュジェル06' },
    { brand: 'TOY\'s×INITY', name: 'MM01' },
    { brand: 'TOY\'s×INITY', name: 'MM02' },
    { brand: 'TOY\'s×INITY', name: 'MM03' },
    { brand: 'TOY\'s×INITY', name: 'MM04' },
    { brand: 'TOY\'s×INITY', name: 'MM05' },
    { brand: 'TOY\'s×INITY', name: 'MM06' },
    { brand: 'TOY\'s×INITY', name: 'BE04' },
    { brand: 'TOY\'s×INITY', name: 'BE05' },
    { brand: 'TOY\'s×INITY', name: 'BE06' },
    { brand: 'TOY\'s×INITY', name: 'SM02' },
    { brand: 'TOY\'s×INITY', name: 'SM03' },
    { brand: 'TOY\'s×INITY', name: 'WA02' },
    { brand: 'TOY\'s×INITY', name: 'LM06' },
    { brand: 'TOY\'s×INITY', name: 'T-PF01 ウルトラ' },
    { brand: 'TOY\'s×INITY', name: 'T-PF05 エレクトリック' },
    { brand: 'TOY\'s×INITY', name: 'T-PF06 バーレスク' },
    { brand: 'TOY\'s×INITY', name: 'T-PF07 ハーレム' },
    { brand: 'TOY\'s×INITY', name: 'T-PF17 オクタゴン' },
    { brand: 'TOY\'s×INITY', name: 'T-PF21 ホープ' },
    { brand: 'INITY', name: 'THM-01' },
    { brand: 'INITY', name: 'THM-02' },
    { brand: 'INITY', name: 'THM-03' },
    { brand: 'INITY', name: 'THM-04' },
    { brand: 'INITY', name: 'THM-05' },
    { brand: 'INITY', name: 'THM-06' },
    { brand: 'INITY', name: 'THM-07' },
    { brand: 'INITY', name: 'UHM-01' },
    { brand: 'INITY', name: 'UHM-05' },
    { brand: 'INITY', name: 'UHM-06' },
    { brand: 'QUE', name: '002 タフィーピンク' },
    { brand: 'QUE', name: '003 ゼニスブルー' },
    { brand: 'QUE', name: '004 サーモンピンク' },
    { brand: 'QUE', name: '005 リーフグリーン' },
    { brand: 'QUE', name: '006 ラベンダーパープル' },
    { brand: 'QUE', name: '007 スカイブルー' },
    { brand: 'QUE', name: '008 ウルトラピンク' },
    { brand: 'NOVEL', name: 'MG02' },
    { brand: 'NOVEL', name: 'MG03' },
    { brand: 'マイビー', name: 'HO-003G' },
    { brand: 'マイビー', name: 'HO-004G' },
    { brand: 'マイビー', name: 'HO-005G' },
    { brand: 'マイビー', name: 'HO-008G' },
    { brand: 'マイビー', name: 'GS-008G' },
    { brand: 'マイビー', name: 'GS-011G' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル モアナジャスミン' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル アップルブリーゼ' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル シャンパンリリー' },
    { brand: 'EORA', name: 'ハンドクリーム ハニージンジャー' },
    { brand: 'EORA', name: 'ハンドクリーム ベルガモット' },
    { brand: 'EORA', name: 'ハンド＆ボディークリーム' },
    { brand: 'PPREANFA', name: 'ccキューティクルオイル' },
    { brand: 'CUCCIO', name: 'キューティクルオイル ザクロ' }
  ];

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store) {
      existing[data[i][3] + '||' + data[i][2]] = true;
    }
  }

  var rows = [];
  var skipped = 0;
  items.forEach(function (item) {
    var key = item.brand + '||' + item.name;
    if (existing[key]) { skipped++; return; }
    existing[key] = true;
    ensureBrandExists_(store, item.brand);
    var code = 'QR-' + Utilities.getUuid().split('-')[0].toUpperCase();
    rows.push([store, code, item.name, item.brand, '', '', '本', '', new Date(), '']);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('木津店: ' + rows.length + '件登録、' + skipped + '件は既存のためスキップ');
}

/**
 * 生駒店の展開準備: 事前に受け取ったスプレッドシート(棚卸表)からブランド・品名
 * (分かる範囲で品番も)だけを一括登録する一時的なスクリプト。木津店のときと同様、
 * 実在庫数やバーコードはここでは扱わない。バーコードは実施日に店舗で1点ずつ
 * スキャンして、ここで登録した既存の品名に紐づける想定。
 * 木津店のbulkImportKizuProducts()と同じ考え方: 一度実行したら不要になる関数。
 */
function bulkImportIkomaProducts() {
  var store = '生駒店';
  var items = [
    { brand: '', name: 'クリストリオ' },
    { brand: '', name: 'スネークベース' },
    { brand: '', name: 'スネークベースストロング' },
    { brand: 'PREGEL', name: 'ノンワイプクリア キャンジェル ノンヒート', partNo: '27759' },
    { brand: 'para gel', name: 'クリアジェルEX', partNo: '99187' },
    { brand: 'KOKOIST', name: 'ノンワイプアクセサリーボンド', partNo: '122263' },
    { brand: 'T-GEL', name: 'D001 クリーミーピンク', partNo: '114891' },
    { brand: 'T-GEL', name: 'Sseries B007', partNo: '123000' },
    { brand: 'T-GEL', name: 'D008 アブソリュートレッド', partNo: '114896' },
    { brand: 'T-GEL', name: 'D016 ホワイトアラバスター 10ml', partNo: '130958' },
    { brand: 'T-GEL', name: 'D029 グレージュ', partNo: '114909' },
    { brand: 'T-GEL', name: 'D032 モーブプラム', partNo: '114912' },
    { brand: 'T-GEL', name: 'D036 ローズボルドー', partNo: '114914' },
    { brand: 'T-GEL', name: 'D042 ディープターコイズ', partNo: '114918' },
    { brand: 'T-GEL', name: 'D045 ネオンオレンジ', partNo: '114920' },
    { brand: 'T-GEL', name: 'D046 ルミナスレッド', partNo: '114921' },
    { brand: 'T-GEL', name: 'D057 スカイブルー', partNo: '114925' },
    { brand: 'T-GEL', name: 'D059 エナメルブルー', partNo: '114926' },
    { brand: 'T-GEL', name: 'D062 フレンチホワイト', partNo: '114928' },
    { brand: 'T-GEL', name: 'D063 アイボリー', partNo: '114929' },
    { brand: 'T-GEL', name: 'D067 ミディアムピンク', partNo: '114932' },
    { brand: 'T-GEL', name: 'D076 マスタード', partNo: '114938' },
    { brand: 'T-GEL', name: 'D086 スモーキーデニム', partNo: '114940' },
    { brand: 'T-GEL', name: 'D096 レッドブラウン', partNo: '114944' },
    { brand: 'T-GEL', name: 'D097　ネイビー' },
    { brand: 'T-GEL', name: 'D103 スケルトンブラック', partNo: '114948' },
    { brand: 'T-GEL', name: 'D105 ブラック', partNo: '114949' },
    { brand: 'T-GEL', name: 'D106 グレー', partNo: '114950' },
    { brand: 'T-GEL', name: 'D107 レッドボルドー', partNo: '114951' },
    { brand: 'T-GEL', name: 'D110 ピンクベージュ', partNo: '114952' },
    { brand: 'T-GEL', name: 'D112 コーラルベージュ', partNo: '114954' },
    { brand: 'T-GEL', name: 'D143 イエロー', partNo: '114958' },
    { brand: 'T-GEL', name: 'D148 グリーン', partNo: '114961' },
    { brand: 'T-GEL', name: 'D157 ミディアムグレー', partNo: '114963' },
    { brand: 'T-GEL', name: 'D159 レインボーシャイン', partNo: '114964' },
    { brand: 'T-GEL', name: 'D175 スケルトンダークブラウン', partNo: '114977' },
    { brand: 'T-GEL', name: 'D182 ヌーディピンク', partNo: '114984' },
    { brand: 'T-GEL', name: 'D189 パステルパープル', partNo: '114990' },
    { brand: 'T-GEL', name: 'D191 ローズマゼンダ', partNo: '114991' },
    { brand: 'T-GEL', name: 'D197 グラデーションピンク', partNo: '114994' },
    { brand: 'T-GEL', name: 'D204 ボルドー', partNo: '114999' },
    { brand: 'T-GEL', name: 'D212 ハニーラベンダー', partNo: '115006' },
    { brand: 'T-GEL', name: 'D221 モカグレージュ', partNo: '115012' },
    { brand: 'T-GEL', name: 'D238 シフォングレー', partNo: '115028' },
    { brand: 'T-GEL', name: 'D239 シフォンアイボリー', partNo: '115029' },
    { brand: 'T-GEL', name: 'D244 シアーベージュ', partNo: '115034' },
    { brand: 'T-GEL', name: 'D245 シアーピンクベージュ', partNo: '115035' },
    { brand: 'T-GEL', name: 'D247 ベージュ', partNo: '115037' },
    { brand: 'T-GEL', name: 'D250 スモーキーピンク', partNo: '117431' },
    { brand: 'T-GEL', name: 'D254 ミストシャインピーチ', partNo: '117438' },
    { brand: 'T-GEL', name: 'D260 テラコッタ', partNo: '128990' },
    { brand: 'T-GEL', name: 'D261 スモーキーオリーブ', partNo: '128991' },
    { brand: 'T-GEL', name: 'ブラックアートジェル', partNo: '131741' },
    { brand: 'T-GEL', name: 'M001メタリックシルバー' },
    { brand: 'T-GEL', name: 'M002メタリックゴールド' },
    { brand: 'emena', name: 'Flash 1013' },
    { brand: 'emena', name: 'Flash gel 1016', partNo: '140314' },
    { brand: 'emena', name: 'Flash gel 1017', partNo: '141788' },
    { brand: 'emena', name: 'Flash gel 1018', partNo: '141789' },
    { brand: 'emena', name: 'Flash gel 1019', partNo: '141790' },
    { brand: 'emena', name: 'Flash gel 1020', partNo: '141791' },
    { brand: 'emena', name: 'Flash gel 1021', partNo: '141793' },
    { brand: 'emena', name: 'Flash 1022' },
    { brand: 'emena', name: 'Flash 1025' },
    { brand: 'emena', name: 'Magnety Flash gel 1317', partNo: '148672' },
    { brand: 'emena', name: 'Magnety Flash gel 1318' },
    { brand: 'emena', name: 'Magnety gel 0504', partNo: '133897' },
    { brand: 'emena', name: 'Magnety Flash gel 1324', partNo: '151187' },
    { brand: 'emena', name: 'Magnety Flash gel 1307' },
    { brand: 'emena', name: 'アートインク E-AIQ5' },
    { brand: 'emena', name: 'Terracotta gel 0702', partNo: '133988' },
    { brand: 'my&bee', name: 'MT-002 ノンワイプメタ ゴールド', partNo: '132479' },
    { brand: 'my&bee', name: 'EN-005S ベア', partNo: '130471' },
    { brand: 'my&bee', name: 'SE-004S グレージュ', partNo: '132483' },
    { brand: 'my&bee', name: 'SE-005S スモーキーカシス', partNo: '132484' },
    { brand: 'my&bee', name: 'RE-005G', partNo: '148950' },
    { brand: 'my&bee', name: 'BT-009MS エスプレッソ', partNo: '130487' },
    { brand: 'my&bee', name: 'ほんちゃんマグ001', partNo: '147470' },
    { brand: 'my&bee', name: 'ほんちゃんマグ003', partNo: '147472' },
    { brand: '', name: 'ほんちゃんマグ004', partNo: '147473' },
    { brand: 'my&bee', name: 'ほんちゃんマグ006', partNo: '147475' },
    { brand: 'my&bee', name: 'ほんちゃんマグ008', partNo: '147477' },
    { brand: 'my&bee', name: 'ぴゅあマグ005', partNo: '147442' },
    { brand: 'my&bee', name: 'ちゅるフラッシュマグ001' },
    { brand: 'my&bee', name: 'ちゅるフラッシュマグ005', partNo: '147454' },
    { brand: 'my&bee', name: 'ちゅるフラッシュマグ010', partNo: '147467' },
    { brand: '', name: 'ちゅるフラッシュマグ012', partNo: '147469' },
    { brand: 'my&bee', name: 'ハニーフラッシュマグ004', partNo: '152271' },
    { brand: 'my&bee', name: 'GS-012G', partNo: '152416' },
    { brand: 'INITY', name: 'THM01 くずきり 10ml', partNo: '141209' },
    { brand: 'INITY', name: 'THM02 さくらもち', partNo: '141210' },
    { brand: 'INITY', name: 'THM-03いちごみつ' },
    { brand: 'INITY', name: 'THM-04もなか' },
    { brand: 'INITY', name: 'THM-05きんつば' },
    { brand: 'INITY', name: 'THM-06ようかん' },
    { brand: 'INITY', name: 'THM-07こはくとう' },
    { brand: 'TOY\'s×INITY', name: 'T-PF18 コーチェラ', partNo: '141291' },
    { brand: 'TOY\'s×INITY', name: 'T-PF19 ラッキー', partNo: '141292' },
    { brand: 'TOY\'s×INITY', name: 'T-PF20 ワンダーランド', partNo: '141293' },
    { brand: 'TOY\'s×INITY', name: 'T-PF21 ホープ', partNo: '141294' },
    { brand: 'TOY\'s×INITY', name: 'T-PF22 グリーンルーム', partNo: '141295' },
    { brand: 'TOY\'s×INITY', name: 'ADVN04 アンバーホワイト', partNo: '144683' },
    { brand: 'CLETO', name: 'ブラック 12μm', partNo: '144603' },
    { brand: 'CLETO', name: 'ホワイト 4μm', partNo: '144600' },
    { brand: 'CLETO', name: 'ピンクベージュ 12μm', partNo: '144606' },
    { brand: 'CLETO', name: 'ブラウン 12μm', partNo: '144609' },
    { brand: 'CLETO', name: 'CLETO クロムマグ 04', partNo: '139204' },
    { brand: 'CLETO', name: 'CLETO グローマグ 10', partNo: '147244' },
    { brand: 'CLETO', name: 'CLETO インク 06', partNo: '136672' },
    { brand: 'KOKOIST', name: 'D-3 Purple×Magenta', partNo: '130894' },
    { brand: 'KOKOIST', name: '#E-238S', partNo: '114072' },
    { brand: 'KOKOIST', name: 'KOKOIST ブリーディングインク ホワイト', partNo: '110777' },
    { brand: 'KOKOIST', name: 'F06 アレキサンドライトフラッシュ', partNo: '142431' },
    { brand: 'KOKOIST', name: '#E-223 ミッドナイトチョコレート', partNo: '110773' },
    { brand: 'KOKOIST', name: '#E-327 バタークリーム', partNo: '144561' },
    { brand: 'ICE GEL', name: 'クリスタルジェル 1461 スワン', partNo: '133520' },
    { brand: 'ICEGEL', name: 'ジェル 1623 ピュアホワイト', partNo: '144309' },
    { brand: 'ICEGEL', name: 'S172 ホワイト', partNo: '138720' },
    { brand: 'ohana', name: 'K02 パウダースノー', partNo: '146317' },
    { brand: 'ohana', name: 'M01 ホワイト', partNo: '146355' },
    { brand: 'STORY JEL365', name: 'GG42 パパラチア', partNo: '148400' },
    { brand: 'STORY JEL365', name: 'STORY JEL365 スターダスト', partNo: '115561' },
    { brand: 'D.nail', name: 'クリアカラー C07 クリアブルー', partNo: '141925' },
    { brand: 'D.nail', name: 'クリアカラー C08 クリアバイオレット', partNo: '141926' },
    { brand: 'ageha', name: '511 ハニーシロップ', partNo: '19912' },
    { brand: 'ageha', name: '5-14 マスカレードオレンジ', partNo: '135128' },
    { brand: 'PREGEL', name: 'DOLL-629 ボルドー', partNo: '37384' },
    { brand: 'PREGEL', name: 'DOLL-706 白銀姫', partNo: '25105' },
    { brand: 'enoi', name: 'enoi galass magnet 113', partNo: '146103' },
    { brand: 'enoi', name: 'diamond flash  f27', partNo: '146146' },
    { brand: 'enoi', name: '（ダイヤモンドフラッシュ） f49', partNo: '153069' },
    { brand: 'enoi', name: '（ミルクマグネット） mg138', partNo: '147041' },
    { brand: 'Miss Mirage', name: 'Miss Mirage カラープラスタージェル ホワイト', partNo: '132530' },
    { brand: 'Miss Mirage', name: 'GH30s エトワ マロン', partNo: '124291' },
    { brand: 'LEAFGEL', name: '485 コルクベージュ', partNo: '118098' },
    { brand: 'TRINA', name: 'CL-6 コッツウォルズコッパー', partNo: '116637' },
    { brand: 'LUMIERE avenir', name: '001SS月の砂漠', partNo: '140463' },
    { brand: 'Putiel', name: '610 シルバー', partNo: '87960' },
    { brand: 'Lily gel', name: '#FR01 レースピンク', partNo: '137069' },
    { brand: 'CLETO', name: 'クリアマグ85μm' },
    { brand: 'CLETO', name: 'クロムマグ08' },
    { brand: 'ネイル工房', name: '5micron magnet 12' },
    { brand: 'FROM THE NAIL', name: 'グレー容器' },
    { brand: 'FROM THE NAIL', name: 'ピンク容器' },
    { brand: 'CLODI', name: '♥グレー容器' },
    { brand: 'CLODI', name: '★ピンク容器' },
    { brand: 'KOKOIST', name: 'ウルトラレインボーノンワイプトップ' },
    { brand: 'KOKOIST', name: 'インク　Black' },
    { brand: 'KOKOIST', name: 'インク　Blue' },
    { brand: 'KOKOIST', name: 'インク　Green' },
    { brand: 'KOKOIST', name: 'インク　purple' },
    { brand: 'KOKOIST', name: 'インク　Red' },
    { brand: 'D.nail', name: '転写ジェル' },
    { brand: 'Mable　Liquid', name: 'MA-04' },
    { brand: 'Mable　Liquid', name: 'MA-05' },
    { brand: 'Mable　Liquid', name: 'MA-09' },
    { brand: 'Mable　Liquid', name: 'MA-10' },
    { brand: 'MGEL', name: 'C816' },
    { brand: 'TOY\'s×INITY', name: '粘土ジェル' },
    { brand: 'ICEGEL', name: 'A Black Glass Parts GEL' },
    { brand: 'CHRISTRIO ✖', name: 'CLEAR GEL　✖' },
    { brand: 'TRINA', name: 'ノンワイプトップジェルライナー' },
    { brand: 'MELTY GEL', name: 'CLEAR GEL' },
    { brand: 'GEL　GRAPH', name: 'Multi-3D gel' },
    { brand: 'PREGEL', name: 'ベースホワイト PG-CE100' },
    { brand: 'PREGEL', name: 'フォレストフロア PG-CE1009' },
    { brand: 'PREGEL', name: 'ストロベリードロップs800' },
    { brand: 'PREGEL', name: 'ブルーベリードロップs808' },
    { brand: 'PREGEL', name: 'ネイビー PG-CE259' },
    { brand: 'PREGEL', name: 'コットンベージュ　PGU-G1023' },
    { brand: 'Palms Graceful', name: 'シャドウブルー177' },
    { brand: 'LIly gel', name: 'SG02' },
    { brand: '&ii', name: 'skin series ⅠⅠⅠ029' },
    { brand: '&ii', name: 'nude Flash series  ⅠⅠⅠ064' },
    { brand: '25gel', name: '038g' },
    { brand: '25gel', name: 'FLASH201' },
    { brand: '25gel', name: 'FLASH196' },
    { brand: '25gel', name: 'MAG&FLASH188' },
    { brand: '25gel', name: '143G' },
    { brand: 'Riccagel', name: '003MS', partNo: '150287' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル　アップルプリーゼ' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル　シャンパンリリー' },
    { brand: 'SHAREYDVA', name: 'キューティクルオイル　モアナジャスミン' },
    { brand: 'EORA', name: 'モイストリンクルクリーム　50g' },
    { brand: 'EORA', name: 'ハンドクリームRO　30g' },
    { brand: 'AROMARY', name: 'アロマリーアロマ　モイスチュアライザー' },
    { brand: 'Riccagel', name: '002M' },
    { brand: 'Myit\'s', name: 'MC003' },
    { brand: 'D.nail', name: 'マットコート' },
    { brand: 'PREGEL', name: 'M110' },
    { brand: 'loade', name: 'FM001' },
    { brand: 'loade', name: 'FM002' },
    { brand: 'Putiel', name: '207' },
    { brand: 'Putiel', name: '240' },
    { brand: 'MIss Mirage', name: 'TM1S' },
    { brand: 'MIss Mirage', name: 'R43S' },
    { brand: 'INITY', name: 'MK02M' },
    { brand: 'INITY', name: 'MK03M' },
    { brand: 'INITY', name: 'MK04M' },
    { brand: 'enoi', name: 'milk#T01' },
    { brand: 'ICEGEL', name: '1156' },
    { brand: 'ICEGEL', name: '1150' },
    { brand: 'ICEGEL', name: '1163' },
    { brand: 'ICEGEL', name: '1249' },
    { brand: 'ICEGEL', name: '1228' },
    { brand: 'ICEGEL', name: '1147' },
    { brand: 'Lilygel', name: 'SG05' },
    { brand: 'T-GEL', name: 'D065' },
    { brand: 'T-GEL', name: 'D-253' },
    { brand: 'TOY\'s×INITY', name: 'ビタベース' },
    { brand: 'emena', name: 'magnetygel0560' },
    { brand: 'CLETO', name: 'スキンマグ03' },
    { brand: 'CLETO', name: 'カームフラッシュマグ03」' },
    { brand: 'CLETO', name: 'カームフラッシュマグ06' },
    { brand: 'MYSTIC JO', name: 'SLI-04　インク' },
    { brand: 'MYSTIC JO', name: 'SLI-05インク' },
    { brand: 'TRINA', name: 'ドリップインクLUXE SILVER15' },
    { brand: 'リーフジェル', name: '124' },
    { brand: 'プリジェル✖', name: '✖' },
    { brand: 'MYSTIC JO', name: '606' },
    { brand: 'TRINA', name: 'PJ3（ラメ）' },
    { brand: 'TOY\'S×INITY', name: 'ノンワイプハケ跡' },
    { brand: 'プリジェル', name: 'COLOR　Z040G' },
    { brand: 'プリジェル', name: 'COLOR　Z077G' },
    { brand: 'プリジェル', name: 'COLOR　Z082g' },
    { brand: 'プリジェル', name: 'COLOR　Z084g' },
    { brand: 'KOKOIST', name: 'e-109' },
    { brand: 'プリジェル', name: 'ホワイトモカポルカM234' },
    { brand: 'enoi', name: 'MG131' },
    { brand: 'TRINA', name: 'SH-37' },
    { brand: 'TRINA', name: 'NC-5' },
    { brand: 'TRINA', name: 'NC-4' },
    { brand: 'TRINA', name: 'SKN-9' },
    { brand: 'TRINA', name: 'SKN-24' },
    { brand: 'MYSTIC JO', name: '614' },
    { brand: 'AnnGel', name: 'ノンワイプマットコート' },
    { brand: 'KOKOIST', name: 'ウルトラグロッシーノンワイプ' },
    { brand: 'ネイルタウン', name: 'マグ01' },
    { brand: 'ネイルタウン', name: 'マグ11' },
    { brand: 'SPELL', name: 'マグN。401' },
    { brand: 'MYSTIC JO✖', name: '✖' },
    { brand: 'ageha', name: '04-06' },
    { brand: 'ageha', name: '02-01' },
    { brand: 'ageha', name: '02-02' },
    { brand: 'ageha', name: '02-03' },
    { brand: 'Riccagel', name: '062M' },
    { brand: 'T-GEL', name: 'D030' },
    { brand: 'D.nail', name: '2ライナーホワイト' },
    { brand: 'D.nail', name: '29' },
    { brand: 'More Couture', name: '505' },
    { brand: 'PREGEL', name: 'アートホイルジェル' },
    { brand: 'INITY', name: 'ノンワイプビジュージェル' },
    { brand: 'Lem', name: 'ms223だいだい' },
    { brand: 'ageha', name: '1-27' },
    { brand: 'ageha', name: '1-32' },
    { brand: 'ageha', name: '404' },
    { brand: 'enoi', name: 'gt17' },
    { brand: 'Pregel', name: 'ml04ライナーBlack' },
    { brand: 'TRINA', name: 'SKN-21' },
    { brand: 'ICEGEL', name: 's208　粘土' },
    { brand: 'naility', name: 'クリア粘土' },
    { brand: 'my&bee', name: 'TF10　フラッシュマグ' },
    { brand: 'emena', name: 'magnetyflsh1341' },
    { brand: 'Barbie', name: 'B-FG14' },
    { brand: 'TOY\'S×INITY', name: 't-pf44フラッシュ' },
    { brand: 'KOKOIST', name: 'bc-32フラッシュ' },
    { brand: 'ニューリージュエル', name: 'jf01フラッシュ' },
    { brand: 'ニューリージュエル', name: 'jf02' },
    { brand: 'enoi', name: 'f41' },
    { brand: 'TOY\'S×INITY', name: 'マグJM01' },
    { brand: 'enoi', name: 'melo #11' },
    { brand: 'lem', name: 'n819' },
    { brand: 'PREGEL', name: '019N' },
    { brand: 'my&bee', name: 'KF-08' },
    { brand: 'my&bee', name: 'KF-03' },
    { brand: 'putiel', name: '323' },
    { brand: 'TOY\'S×INITY', name: 'kiwami crack gel' },
    { brand: 'pompomgel', name: '6' },
    { brand: 'pompomgel', name: '19' },
    { brand: 'pompomgel', name: '31' },
    { brand: 'd.nail', name: '1' },
    { brand: 'ＪAM', name: 'jー05' },
    { brand: 'ＪAM', name: 't-04' },
    { brand: 'ＪAM', name: 't-07' },
    { brand: 'ＪAM', name: 'l-05' },
    { brand: 'ICEGEL', name: 'マーブルポーション' },
    { brand: 'CLODI', name: 'ピンク♥マグ' },
    { brand: 'CLODI', name: 'ベージュ　花マークマグ' },
    { brand: 'my&bee', name: 'パウダーフリーコート' },
    { brand: 'GRACE', name: 'dazzling03' },
    { brand: 'GRACE', name: 'dazzling04' },
    { brand: 'GRACE', name: 'dazzling05' },
    { brand: 'GRACE', name: 'Nuance A09' },
    { brand: 'GRACE', name: 'Clear Neon H05' },
    { brand: 'CLETO', name: 'bubble mag 04' },
    { brand: 'CLETO', name: 'bubble mag 12' },
    { brand: 'STORY JEL365', name: 'AQ02マグ' },
    { brand: 'mon PeTiT LAPin', name: '211s ウォーターオレンジ' },
    { brand: 'mon PeTiT LAPin', name: '212Ｓ　ウォーターパイン' },
    { brand: 'mon PeTiT LAPin', name: '214s　ウォーターロビンスエッグ' },
    { brand: 'mon PeTiT LAPin', name: '217S　ウォータードラゴンフルーツ' },
    { brand: 'enoi', name: 'melo  #11　フラッシュ' },
    { brand: 'MON PETIT LAPIN', name: 'FLASHGEL　07' },
    { brand: '＆ii', name: 'gloss series  III004' },
    { brand: '＆ii', name: 'bubble   III206' },
    { brand: '＆ii', name: 'Mermaid    III170' },
    { brand: '＆ii', name: 'Mermaid    III176' },
  ];

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store) {
      existing[data[i][3] + '||' + data[i][2]] = true;
    }
  }

  var rows = [];
  var skipped = 0;
  items.forEach(function (item) {
    var key = item.brand + '||' + item.name;
    if (existing[key]) { skipped++; return; }
    existing[key] = true;
    ensureBrandExists_(store, item.brand);
    var code = 'QR-' + Utilities.getUuid().split('-')[0].toUpperCase();
    rows.push([store, code, item.name, item.brand, '', '', '本', '', new Date(), '', item.partNo || '']);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('生駒店: ' + rows.length + '件登録、' + skipped + '件は既存のためスキップ');
}
