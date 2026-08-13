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
var PRODUCTS_HEADERS = ['店舗', 'コード', '商品名', 'ブランド', 'カテゴリ', 'メーカー', '単位', '備考', '登録日', 'カラーNO'];
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

  seedAccounts_(ss);
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
