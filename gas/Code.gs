/**
 * Web App エントリポイントと在庫管理のビジネスロジック。
 * フロントエンド(web/store, web/hq)からは doPost へ { action, payload } 形式のJSONを
 * text/plain として送信する(preflight回避のため。詳細はREADME参照)。
 */

function doGet(e) {
  return jsonOutput_({ ok: true, message: 'Nail Salon Inventory API is running' });
}

function doPost(e) {
  var action, payload;
  try {
    var body = JSON.parse(e.postData.contents);
    action = body.action;
    payload = body.payload || {};
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'リクエストの形式が不正です' });
  }

  try {
    var result = routeAction_(action, payload);
    return jsonOutput_({ ok: true, data: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function routeAction_(action, p) {
  switch (action) {
    case 'login':
      return login_(p.username, p.password);
    case 'getStaffList':
      return getStaffList_(validateToken_(p.token), p.store);
    case 'addStaff':
      return addStaff_(validateToken_(p.token), p);
    case 'setStaffActive':
      return setStaffActive_(validateToken_(p.token), p);
    case 'lookupProduct':
      return lookupProductForClient_(validateToken_(p.token), p);
    case 'registerProduct':
      return registerProduct_(validateToken_(p.token), p);
    case 'recordIncoming':
      return recordIncoming_(validateToken_(p.token), p);
    case 'recordDisposal':
      return recordDisposal_(validateToken_(p.token), p);
    case 'submitStocktake':
      return submitStocktake_(validateToken_(p.token), p);
    case 'getInventorySummary':
      return getInventorySummary_(validateToken_(p.token), p.store, p.month);
    case 'lookupCurrentStock':
      return lookupCurrentStock_(validateToken_(p.token), p);
    case 'getOutOfStock':
      return getOutOfStock_(validateToken_(p.token), p.store);
    case 'deleteLogEntry':
      return deleteLogEntry_(validateToken_(p.token), p.rowIndex);
    case 'clearStoreLog':
      return clearStoreLog_(validateToken_(p.token), p);
    case 'listStores':
      return { stores: STORES };
    case 'listProducts':
      return listProducts_(validateToken_(p.token), p.store);
    case 'deleteProduct':
      return deleteProduct_(validateToken_(p.token), p);
    case 'getLogEntries':
      return getLogEntries_(validateToken_(p.token), p.store, p.limit, p.type, p.staffName);
    case 'createAccount':
      return createAccount_(validateToken_(p.token), p);
    case 'resetPassword':
      return resetPassword_(validateToken_(p.token), p);
    case 'updateProduct':
      return updateProduct_(validateToken_(p.token), p);
    case 'getBrandList':
      return getBrandList_(validateToken_(p.token), p.store);
    case 'addBrand':
      return addBrand_(validateToken_(p.token), p);
    case 'mergeBrand':
      return mergeBrand_(validateToken_(p.token), p);
    case 'getCategoryList':
      return getCategoryList_(validateToken_(p.token));
    case 'resetProductStock':
      return resetProductStock_(validateToken_(p.token), p);
    case 'adjustProductStock':
      return adjustProductStock_(validateToken_(p.token), p);
    case 'resetStoreInventory':
      return resetStoreInventory_(validateToken_(p.token), p);
    case 'getStocktakeReview':
      return getStocktakeReview_(validateToken_(p.token), p.store, p.month);
    case 'approveStocktake':
      return approveStocktake_(validateToken_(p.token), p);
    case 'rejectStocktake':
      return rejectStocktake_(validateToken_(p.token), p);
    case 'getStocktakeStatus':
      return getStocktakeStatusForClient_(validateToken_(p.token), p.store, p.month);
    default:
      throw new Error('不明なactionです: ' + action);
  }
}

// ---- スタッフマスタ ----

function getStaffList_(session, storeParam) {
  var store = session.role === 'hq' ? storeParam || null : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  var sheet = getSheet_(SHEET_STAFF);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var s = data[i][0], name = data[i][1], active = data[i][2];
    if (active === false) continue;
    if (store && s !== store) continue;
    list.push({ store: s, name: name });
  }
  return { staff: list };
}

function addStaff_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);
  if (!p.name) throw new Error('スタッフ名を入力してください');
  var sheet = getSheet_(SHEET_STAFF);
  sheet.appendRow([store, p.name, true]);
  return { store: store, name: p.name };
}

function setStaffActive_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);
  var sheet = getSheet_(SHEET_STAFF);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store && data[i][1] === p.name) {
      sheet.getRange(i + 1, 3).setValue(!!p.active);
      return { store: store, name: p.name, active: !!p.active };
    }
  }
  throw new Error('スタッフが見つかりません');
}

// ---- 商品マスタ ----
// 商品は「店舗」×「コード」で一意。店舗ごとに独立したマスタを持ち、同じバーコードでも
// 店舗が違えば別の商品として登録できる。

/** 内部用の純粋な検索。store・codeを直接指定する(セッションの解釈は呼び出し側で行う)。 */
function lookupProduct_(store, code) {
  if (!store || !code) return null;
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store && String(data[i][1]) === String(code)) {
      return {
        store: data[i][0],
        code: data[i][1],
        name: data[i][2],
        brand: data[i][3],
        category: data[i][4],
        maker: data[i][5],
        unit: data[i][6],
        colorNo: data[i][9],
        itemNumber: data[i][10]
      };
    }
  }
  return null;
}

/** クライアントの「lookupProduct」action用。storeはセッションから解決する。 */
function lookupProductForClient_(session, p) {
  var store = session.role === 'hq' ? (p.store || session.store) : session.store;
  return lookupProduct_(store, p.code);
}

/**
 * 新規商品登録。本社アカウントのみ呼べる(商品マスタは本社が店舗を選んで管理する)。
 * バーコードがある商品はそのコードで、ない商品は code を空にして呼ぶと自前QRコードを
 * 発行する。既存商品の上書きはできない(店舗+コードが重複していればエラーになる)。
 * 既存商品の修正は updateProduct_ を使う。ブランドは新しい名前ならその店舗の
 * ブランドマスタに自動登録される。
 */
function registerProduct_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.name) throw new Error('品名を入力してください');
  var store = p.store;
  if (!store) throw new Error('店舗を指定してください');

  var code = p.code;
  if (!code) {
    code = 'QR-' + Utilities.getUuid().split('-')[0].toUpperCase();
  } else if (lookupProduct_(store, code)) {
    throw new Error('このコードは既に登録されています');
  }

  ensureBrandExists_(store, p.brand);
  ensureCategoryExists_(p.category);

  var sheet = getSheet_(SHEET_PRODUCTS);
  sheet.appendRow([
    store, code, p.name, p.brand || '', p.category || '', p.maker || '',
    p.unit || '本', p.memo || '', new Date(), p.colorNo || '', p.itemNumber || ''
  ]);
  return { store: store, code: code, name: p.name, brand: p.brand || '', colorNo: p.colorNo || '' };
}

/**
 * 既存商品の品名・ブランド・カテゴリー・カラーNO・品番・メモの修正。本社アカウントのみ。
 * p.newCode を指定すると、コード自体も変更できる(バーコードなしで自動発行した
 * プレースホルダーのコードを、後から見つかった実物のバーコードに差し替える用途)。
 * 注意: コードを変更すると、それ以前にそのコードで記録された取引ログ(棚卸・入荷・廃棄)
 * は新しいコードには紐づかず、現在庫の計算上は新しいコードの履歴がない状態から
 * 始まる。まだ取引ログが無い(登録直後の)商品のコード変更を想定している。
 */
function updateProduct_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store || !p.code) throw new Error('店舗とコードを指定してください');

  ensureBrandExists_(p.store, p.brand);
  ensureCategoryExists_(p.category);

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === p.store && String(data[i][1]) === String(p.code)) {
      var row = i + 1;
      var newCode = p.newCode ? String(p.newCode).trim() : '';
      if (newCode && newCode !== String(p.code)) {
        if (lookupProduct_(p.store, newCode)) {
          throw new Error('このコードは既に登録されています');
        }
        sheet.getRange(row, 2).setValue(newCode);
      }
      sheet.getRange(row, 3).setValue(p.name || '');
      sheet.getRange(row, 4).setValue(p.brand || '');
      sheet.getRange(row, 5).setValue(p.category || '');
      sheet.getRange(row, 8).setValue(p.memo || '');
      sheet.getRange(row, 10).setValue(p.colorNo || '');
      sheet.getRange(row, 11).setValue(p.itemNumber || '');
      refreshSummary_(); // ブランド名等の変更をサマリ表示にも反映する
      return { store: p.store, code: newCode || p.code };
    }
  }
  throw new Error('商品が見つかりません');
}

/** 指定した店舗の商品マスタ一覧。本社は店舗を指定して閲覧する。 */
function listProducts_(session, storeParam) {
  var store = session.role === 'hq' ? storeParam : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  if (!store) throw new Error('店舗を指定してください');

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== store) continue;
    list.push({
      code: data[i][1], name: data[i][2], brand: data[i][3],
      category: data[i][4], maker: data[i][5], unit: data[i][6],
      memo: data[i][7], colorNo: data[i][9], itemNumber: data[i][10]
    });
  }
  return { products: list };
}

/** 既存商品の削除。本社アカウントのみ。 */
function deleteProduct_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store || !p.code) throw new Error('店舗とコードを指定してください');
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === p.store && String(data[i][1]) === String(p.code)) {
      sheet.deleteRow(i + 1);
      return { deleted: p.code };
    }
  }
  throw new Error('商品が見つかりません');
}

// ---- ブランドマスタ ----
// 店舗ごとに独立したブランド一覧。商品登録・編集時に新しいブランド名が入力されると
// 自動でその店舗のブランドマスタに追加され、以後は入力候補として表示される。
// 本社の「ブランド管理」画面からも、店舗を指定して手動で追加できる。

function getBrandNamesForStore_(store) {
  var sheet = getSheet_(SHEET_BRANDS);
  var data = sheet.getDataRange().getValues();
  var brands = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store && data[i][1]) brands.push(data[i][1]);
  }
  // 選択肢を辞書順にしておくと、プルダウンで頭文字を入力したときの候補ジャンプが機能しやすい
  // ブランド名が数字だけだとGoogleスプレッドシート側で数値型として返ってくることがあり、
  // その場合localeCompareが無いため呼び出し前に必ず文字列化する
  brands.sort(function (a, b) { return String(a).localeCompare(String(b), 'ja'); });
  return brands;
}

function getBrandList_(session, storeParam) {
  var store = session.role === 'hq' ? storeParam : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  if (!store) throw new Error('店舗を指定してください');
  return { brands: getBrandNamesForStore_(store) };
}

/** 新しいブランド名なら自動でそのブランドマスタに追加する(登録・編集の裏側で使う)。 */
function ensureBrandExists_(store, name) {
  if (!store || !name) return;
  if (getBrandNamesForStore_(store).indexOf(name) !== -1) return;
  getSheet_(SHEET_BRANDS).appendRow([store, name]);
}

/** 本社が手動でブランドを追加登録する。 */
function addBrand_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store) throw new Error('店舗を指定してください');
  if (!p.name) throw new Error('ブランド名を入力してください');
  if (getBrandNamesForStore_(p.store).indexOf(p.name) !== -1) {
    throw new Error('このブランド名は既に登録されています');
  }
  getSheet_(SHEET_BRANDS).appendRow([p.store, p.name]);
  return { store: p.store, name: p.name };
}

/**
 * 表記ゆれ(大文字/小文字、全角/半角、スペースの有無など)で複数の書き方になっている
 * ブランド名を1つに統一する。p.from(統一前の表記の配列)がついた商品マスタの行をすべて
 * p.to(統一後の表記)に書き換え、統一前の表記をブランドマスタから削除する。
 */
function mergeBrand_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store) throw new Error('店舗を指定してください');
  if (!p.to) throw new Error('統一後のブランド名を指定してください');
  var from = p.from || [];
  if (!from.length) throw new Error('統一するブランド名を指定してください');

  ensureBrandExists_(p.store, p.to);

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var updatedCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === p.store && from.indexOf(data[i][3]) !== -1) {
      sheet.getRange(i + 1, 4).setValue(p.to);
      updatedCount++;
    }
  }

  var brandSheet = getSheet_(SHEET_BRANDS);
  var brandData = brandSheet.getDataRange().getValues();
  for (var j = brandData.length - 1; j >= 1; j--) {
    if (brandData[j][0] === p.store && from.indexOf(brandData[j][1]) !== -1) {
      brandSheet.deleteRow(j + 1);
    }
  }

  refreshSummary_(); // まとめて直す操作なので、参考シートも合わせて更新しておく
  return { store: p.store, to: p.to, updatedCount: updatedCount };
}

// ---- カテゴリマスタ ----
// ブランドと違い、全店舗共通の1つの分類リスト。商品登録・編集時に新しいカテゴリー名が
// 入力されると自動でカテゴリマスタに追加され、以後は全店舗の入力候補として表示される。

function getCategoryNames_() {
  var sheet = getSheet_(SHEET_CATEGORIES);
  var data = sheet.getDataRange().getValues();
  var categories = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) categories.push(data[i][0]);
  }
  return categories;
}

function getCategoryList_(session) {
  return { categories: getCategoryNames_() };
}

/** 新しいカテゴリー名なら自動でカテゴリマスタに追加する(登録・編集の裏側で使う)。 */
function ensureCategoryExists_(name) {
  if (!name) return;
  if (getCategoryNames_().indexOf(name) !== -1) return;
  getSheet_(SHEET_CATEGORIES).appendRow([name]);
}

// ---- 取引ログ(入荷・廃棄・棚卸) ----
// 入荷登録・破棄登録はいずれも本社アカウントが店舗を選んで行う(本社が全店舗の納品・廃棄を
// 一元管理する運用のため)。店舗アカウントは棚卸のみ自分で記録する。

/**
 * timestampを省略すると現在時刻を使う。棚卸のように1回の送信で複数商品を記録する
 * 場合は、呼び出し側で1つのDateを作って全商品分に同じ値を渡すことで、「どの記録が
 * 同じ送信に属するか」を後から突き合わせられるようにしている。
 */
function appendLog_(store, staffName, product, type, quantity, memo, timestamp) {
  var sheet = getSheet_(SHEET_LOG);
  var ts = timestamp || new Date();
  sheet.appendRow([ts, store, staffName || '', product.code, product.name, product.brand, type, quantity, memo || '']);
  revokeApprovalIfStale_(store, ts);
}

/**
 * その店舗・その月がすでに本社承認済みの場合、新しい取引ログ(棚卸・入荷・廃棄)が
 * 入った時点で承認を取り消し、未承認に戻す。承認後にデータが変わったのに承認済みの
 * 表示が残ってしまうのを防ぐため。送信自体は制限しない(承認有無に関わらず記録できる)。
 */
function revokeApprovalIfStale_(store, timestamp) {
  var month = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy-MM');
  var sheet = getSheet_(SHEET_APPROVALS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store && data[i][1] === month) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

/**
 * 本社が店舗を選んで納品(入荷)を登録する。
 * refreshSummary_()(取引ログ全体を読み直して「現在庫サマリ」シートを書き直す処理)は
 * ここでは呼ばない。入荷登録はスキャンごとに何度も実行される操作で、取引ログが増えるほど
 * refreshSummary_()が遅くなり体感速度に直結するため。「現在庫サマリ」はアプリの画面表示
 * では使っておらず(画面は毎回取引ログから直接集計している)、スプレッドシートを直接見る
 * ための参考シートなので、最新化したい場合はApps Scriptエディタから
 * refreshCurrentStockSummary() を手動実行する。
 */
function recordIncoming_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store) throw new Error('店舗を指定してください');
  var product = lookupProduct_(p.store, p.code);
  if (!product) throw new Error('商品が見つかりません。先に商品マスタへ登録してください');
  var quantity = Number(p.quantity) || 1;
  appendLog_(p.store, session.username, product, '入荷', quantity, p.memo);
  return { product: product, quantity: quantity };
}

/** 本社が店舗を選んで廃棄(劣化・不良などによる在庫減)を登録する。refreshSummary_()を呼ばない理由はrecordIncoming_と同じ。 */
function recordDisposal_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store) throw new Error('店舗を指定してください');
  var product = lookupProduct_(p.store, p.code);
  if (!product) throw new Error('商品が見つかりません');
  var quantity = Number(p.quantity) || 1;
  appendLog_(p.store, session.username, product, '廃棄', quantity, p.memo);
  return { product: product, quantity: quantity };
}

/**
 * items: [{ code, count }, ...] 棚卸で連続スキャンした結果をまとめて送信する。
 * mode: 'add'なら今月すでに記録されている数量に加算、'overwrite'(省略時のデフォルト)なら
 * 今回数えた数でそのまま置き換える。承認済みの月への送信はロックされ、エラーになる
 * (差し戻しを受けるまで再送信できない)。
 */
function submitStocktake_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);

  var month = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  // 承認状態(SHEET_APPROVALS)は行数が少なく軽いのでここで毎回確認する。取引ログ全体を
  // 読む重い処理(加算モードのベース値計算)は、必要な場合のみ・1回だけに絞る(下記)。
  if (getApprovalStatus_(store, month).status === '承認済み') {
    throw new Error('今月は本社承認済みのため棚卸はできません。修正が必要な場合は本社に差し戻しを依頼してください');
  }
  var mode = p.mode === 'add' ? 'add' : 'overwrite';
  // 加算モードのときだけ、今月のこの店舗の棚卸ログを1回だけ読み込んで商品ごとの最新数量を
  // マップ化する。商品ごとに毎回シートを読み直すと件数分だけ遅くなるため。
  var thisMonthQtyByCode = mode === 'add' ? getThisMonthStocktakeMap_(store, month) : {};

  var items = p.items || [];
  // 商品マスタを1回だけ読み込んでマップ化する。棚卸は1回の送信で多数の商品(場合によっては
  // 数百件)を扱うため、商品ごとにlookupProduct_で毎回シートを読み直すと件数分だけ遅くなる。
  var productMap = getProductMap_(store);
  var timestamp = new Date(); // 送信内の全商品に同じ時刻を持たせ、1回の送信として後から突き合わせられるようにする
  var recorded = [];
  var unknownCodes = [];
  items.forEach(function (item) {
    var product = productMap[String(item.code)];
    if (!product) {
      unknownCodes.push(item.code);
      return;
    }
    var qty = mode === 'add' && hasKey_(thisMonthQtyByCode, item.code)
      ? thisMonthQtyByCode[item.code] + item.count
      : item.count;
    appendLog_(store, p.staffName, product, '棚卸', qty, '', timestamp);
    recorded.push({ code: item.code, name: product.name, brand: product.brand, count: qty });
  });
  // refreshSummary_()(現在庫サマリの再構築)や、この送信結果を反映したgetInventorySummary_の
  // 再計算はここでは呼ばない。どちらも取引ログ・商品マスタ全体を読み直す重い処理で、画面側は
  // 呼び出し結果のsummaryを使っておらず(recorded/unknownCodesのみ使用)、完全に無駄になって
  // いたため削除した。最新の在庫は各画面が表示時に毎回計算し直す。

  return {
    recorded: recorded,
    unknownCodes: unknownCodes
  };
}

/** 商品マスタを1回のシート読み込みで「コード→商品」のマップにする(棚卸の一括処理用)。 */
function getProductMap_(store) {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== store) continue;
    map[String(data[i][1])] = {
      store: data[i][0],
      code: data[i][1],
      name: data[i][2],
      brand: data[i][3],
      category: data[i][4],
      maker: data[i][5],
      unit: data[i][6],
      colorNo: data[i][9],
      itemNumber: data[i][10]
    };
  }
  return map;
}

/** その店舗・今月の棚卸ログから、商品コードごとの最新数量のマップを1回のシート読み込みで作る。 */
function getThisMonthStocktakeMap_(store, month) {
  var curCutoff = monthEndCutoff_(month);
  var prevCutoff = monthEndCutoff_(previousMonthString_(month));
  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] !== store || row[6] !== '棚卸') continue;
    var ts = new Date(row[0]).getTime();
    if (ts < prevCutoff.getTime() || ts >= curCutoff.getTime()) continue;
    map[row[3]] = Number(row[7]) || 0; // ログは追記順なので、後の行が常に最新の値になる
  }
  return map;
}

function deleteLogEntry_(session, rowIndex) {
  requireRole_(session, ['hq']);
  var sheet = getSheet_(SHEET_LOG);
  rowIndex = Number(rowIndex);
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    throw new Error('無効な行番号です');
  }
  sheet.deleteRow(rowIndex);
  refreshSummary_();
  return { deleted: rowIndex };
}

/**
 * 指定した店舗の取引ログを全部削除する(取り消し不可)。テスト運用で入った
 * データを、本番運用の前にまとめて消すためのもの。全店舗を対象にした一括削除は
 * 事故防止のためできない(店舗を1つ指定する必要がある)。取引ログを消すと、その
 * 店舗のデータを根拠にした「承認済み」表示も意味を持たなくなるため、承認記録も
 * 合わせて削除する。
 */
function clearStoreLog_(session, p) {
  requireRole_(session, ['hq']);
  verifyOwnPassword_(session, p.password);
  if (!p.store) throw new Error('店舗を指定してください');
  // typeを指定すると、その種別(棚卸/入荷/廃棄/調整)のログだけを削除する。
  // 例えばテストの棚卸データだけを消して、入荷・廃棄の実データは残したい場合に使う。
  var type = p.type || null;

  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var deletedCount = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === p.store && (!type || data[i][6] === type)) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  // 棚卸データを消す場合(typeが棚卸、または全種別クリア)は、その店舗の承認状態も
  // 意味を持たなくなるため合わせて削除し、未実施からやり直せるようにする。
  // 入荷・廃棄だけを消す場合は棚卸の承認状態には影響しないため触らない。
  if (!type || type === '棚卸') {
    var approvalSheet = getSheet_(SHEET_APPROVALS);
    var approvalData = approvalSheet.getDataRange().getValues();
    for (var j = approvalData.length - 1; j >= 1; j--) {
      if (approvalData[j][0] === p.store) {
        approvalSheet.deleteRow(j + 1);
      }
    }
  }

  refreshSummary_();
  return { store: p.store, deletedCount: deletedCount };
}

function getLogEntries_(session, storeParam, limit, type, staffName) {
  var store = session.role === 'hq' ? storeParam || null : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    if (store && row[1] !== store) continue;
    if (type && row[6] !== type) continue;
    if (staffName && String(row[2]).indexOf(staffName) === -1) continue;
    rows.push({
      rowIndex: i + 1,
      timestamp: row[0],
      store: row[1],
      staffName: row[2],
      code: row[3],
      name: row[4],
      brand: row[5],
      type: row[6],
      quantity: row[7],
      memo: row[8]
    });
    if (limit && rows.length >= limit) break;
  }
  return { logs: rows };
}

// ---- 在庫集計 ----
// 直近の棚卸カウントをその時点の実在庫として採用し、それ以降の入荷(+)・廃棄(-)を
// 加減算する。取引ログは appendRow のみで追加されるため常に時系列順であることを前提にしている。
// cutoffDate を渡すと、その日時より前の取引ログだけで集計する(=過去のある時点の在庫を再現する)。

function computeAllSummary_() {
  return computeAllSummaryAsOf_(null);
}

function computeAllSummaryAsOf_(cutoffDate) {
  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (cutoffDate && new Date(row[0]).getTime() >= cutoffDate.getTime()) continue;

    var store = row[1], code = row[3], name = row[4], brand = row[5], type = row[6], qty = Number(row[7]) || 0;
    var key = store + '||' + code;
    if (!map[key]) {
      map[key] = { store: store, code: code, name: name, brand: brand, current: 0, lastStocktakeCount: null };
    }
    var entry = map[key];
    entry.name = name;
    entry.brand = brand;
    if (type === '棚卸') {
      entry.current = qty;
      entry.lastStocktakeCount = qty;
    } else if (type === '入荷') {
      entry.current += qty;
    } else if (type === '廃棄') {
      entry.current -= qty;
    } else if (type === '調整') {
      entry.current = qty;
    }
  }

  return Object.keys(map).map(function (k) {
    return map[k];
  });
}

/** "YYYY-MM" 形式の月文字列から、その月の翌月1日0時(=集計の上限、この日時より前だけを含める)を作る。 */
function monthEndCutoff_(month) {
  if (!month) return null;
  var parts = month.split('-');
  var year = Number(parts[0]);
  var monthIndex1based = Number(parts[1]);
  if (!year || !monthIndex1based) throw new Error('月の指定が不正です');
  return new Date(year, monthIndex1based, 1); // Dateの月は0始まりなので、これで「指定月の翌月1日」になる
}

/** 「店舗||コード」→カラーNOの対応表。在庫一覧の検索用にカラーNOも一緒に返すために使う。 */
function getProductColorNoMap_() {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[data[i][0] + '||' + data[i][1]] = data[i][9] || '';
  }
  return map;
}

/** month(YYYY-MM)を指定すると、その月末時点の在庫スナップショットを返す。省略時は現在の在庫。 */
function getInventorySummary_(session, storeParam, month) {
  var store = session.role === 'hq' ? storeParam || null : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }

  var all = computeAllSummaryAsOf_(monthEndCutoff_(month));
  if (store) all = all.filter(function (e) { return e.store === store; });

  var colorNoMap = getProductColorNoMap_();
  var items = all.map(function (e) {
    return {
      store: e.store,
      code: e.code,
      name: e.name,
      brand: e.brand,
      colorNo: colorNoMap[e.store + '||' + e.code] || '',
      currentStock: e.current,
      outOfStock: e.current <= 0,
      lastStocktakeCount: e.lastStocktakeCount
    };
  });

  items.sort(function (a, b) {
    if (a.store !== b.store) return a.store < b.store ? -1 : 1;
    // 品名が数字だけの商品(例: VETRO「552」)はGoogleスプレッドシート側で数値型として
    // 返ってくることがあり、その場合localeCompareが無いため必ず文字列化してから比較する
    if (a.brand !== b.brand) return String(a.brand || '').localeCompare(String(b.brand || ''), 'ja');
    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  });

  // 本社が店舗を指定せず全体を見る場合のみ、店舗別合計と総合計を付与する(店舗ロールには渡さない)。
  if (session.role === 'hq' && !store) {
    var totalsByStore = {};
    var grandTotal = 0;
    items.forEach(function (r) {
      totalsByStore[r.store] = (totalsByStore[r.store] || 0) + r.currentStock;
      grandTotal += r.currentStock;
    });
    return { items: items, totalsByStore: totalsByStore, grandTotal: grandTotal };
  }
  return { items: items };
}

function getOutOfStock_(session, storeParam) {
  var summary = getInventorySummary_(session, storeParam);
  return { items: summary.items.filter(function (i) { return i.outOfStock; }) };
}

/** 「在庫検索」画面用。バーコード1件をスキャンして、自店の最新在庫だけを軽量に返す。 */
function lookupCurrentStock_(session, p) {
  if (!p.code) throw new Error('コードが指定されていません');
  var store = session.role === 'hq' ? (p.store || session.store) : session.store;

  var product = lookupProduct_(store, p.code);
  if (!product) return { found: false };

  var all = computeAllSummary_();
  var entry = null;
  for (var i = 0; i < all.length; i++) {
    if (all[i].store === store && String(all[i].code) === String(p.code)) {
      entry = all[i];
      break;
    }
  }
  var currentStock = entry ? entry.current : 0;
  return {
    found: true,
    store: store,
    code: product.code,
    name: product.name,
    brand: product.brand,
    colorNo: product.colorNo,
    currentStock: currentStock,
    outOfStock: currentStock <= 0
  };
}

// ---- 棚卸履歴・承認 ----
// 「先月末在庫 + 今月の入荷 − 今月の廃棄」と「今月末在庫(棚卸後)」を商品ごとに突き合わせ、
// 棚卸で生じた差異を確認できるようにする。本社が内容を確認して承認すると、SHEET_APPROVALS
// に店舗×年月の記録が残る(データの修正・削除自体は制限しない、確認済みの記録のみ)。
// 本社・店舗どちらのアカウントからも閲覧できるが、承認できるのは本社のみ。

/** "YYYY-MM" の前月を "YYYY-MM" で返す。 */
function previousMonthString_(month) {
  var parts = month.split('-');
  var year = Number(parts[0]);
  var m = Number(parts[1]) - 1;
  if (m === 0) {
    m = 12;
    year -= 1;
  }
  return year + '-' + (m < 10 ? '0' + m : String(m));
}

function getApprovalStatus_(store, month) {
  var sheet = getSheet_(SHEET_APPROVALS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === store && data[i][1] === month) {
      var approver = data[i][2], approvedAt = data[i][3];
      // 状態列(E列)が無い古い行は、承認者が入っていれば「承認済み」扱いにする(後方互換)。
      var status = data[i][4] || (approver ? '承認済み' : '');
      return {
        approved: status === '承認済み',
        approver: approver,
        approvedAt: approvedAt,
        status: status,
        rejectedReason: data[i][5] || '',
        rejectedAt: data[i][6] || null
      };
    }
  }
  return { approved: false, status: '' };
}

/**
 * 店舗×年月の棚卸状態(未実施/承認待ち/差し戻し/承認済み)をまとめて返す。
 * 承認済み・差し戻しはSHEET_APPROVALSに明示的な行がある場合のみ。それ以外(未実施/承認待ち)は
 * 行を作らず、取引ログにその月の棚卸行があるかどうかだけで判定する。
 */
function getMonthlyStocktakeStatus_(store, month) {
  var approval = getApprovalStatus_(store, month);
  var implementedDate = getLatestStocktakeDateThisMonth_(store, month);
  if (approval.status === '承認済み' || approval.status === '差し戻し') {
    return {
      status: approval.status,
      implementedDate: implementedDate,
      approver: approval.approver || '',
      approvedAt: approval.approvedAt || null,
      rejectedReason: approval.rejectedReason || '',
      rejectedAt: approval.rejectedAt || null
    };
  }
  return {
    status: implementedDate ? '承認待ち' : '未実施',
    implementedDate: implementedDate,
    approver: '', approvedAt: null, rejectedReason: '', rejectedAt: null
  };
}

/** その店舗・その月の棚卸ログのうち、最新の実施日(Date)。1件もなければnull。 */
function getLatestStocktakeDateThisMonth_(store, month) {
  var curCutoff = monthEndCutoff_(month);
  var prevCutoff = monthEndCutoff_(previousMonthString_(month));
  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var latest = null;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] !== store || row[6] !== '棚卸') continue;
    var ts = new Date(row[0]);
    if (ts.getTime() < prevCutoff.getTime() || ts.getTime() >= curCutoff.getTime()) continue;
    if (!latest || ts.getTime() > latest.getTime()) latest = ts;
  }
  return latest;
}

/** 棚卸開始前の状態確認(軽量)。店舗は自店のみ、本社はstore指定で他店舗も見れる。 */
function getStocktakeStatusForClient_(session, storeParam, month) {
  var store = session.role === 'hq' ? storeParam : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  if (!store) throw new Error('店舗を指定してください');
  if (!month) throw new Error('対象月を指定してください');
  return getMonthlyStocktakeStatus_(store, month);
}

/** 指定した店舗・年月の、商品ごとの棚卸差異一覧と、その月の棚卸実施記録・承認状況をまとめて返す。 */
function getStocktakeReview_(session, storeParam, month) {
  var store = session.role === 'hq' ? storeParam : session.store;
  if (session.role !== 'hq' && storeParam && storeParam !== session.store) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
  if (!store) throw new Error('店舗を指定してください');
  if (!month) throw new Error('対象月を指定してください');

  var curCutoff = monthEndCutoff_(month);
  var prevCutoff = monthEndCutoff_(previousMonthString_(month));

  var prevMap = {};
  computeAllSummaryAsOf_(prevCutoff).forEach(function (e) {
    if (e.store === store) prevMap[e.code] = e.current;
  });
  var curMap = {};
  computeAllSummaryAsOf_(curCutoff).forEach(function (e) {
    if (e.store === store) curMap[e.code] = e.current;
  });

  var incomingMap = {}, disposalMap = {};
  // 棚卸は1回の送信で複数商品を記録するが、appendLog_で送信内の全行に同じ時刻を
  // 持たせているため、時刻をキーにまとめれば「1回の送信」単位のイベントに戻せる。
  var stocktakeEventMap = {};
  var logSheet = getSheet_(SHEET_LOG);
  var logData = logSheet.getDataRange().getValues();
  for (var i = 1; i < logData.length; i++) {
    var row = logData[i];
    if (row[1] !== store) continue;
    var ts = new Date(row[0]).getTime();
    if (ts < prevCutoff.getTime() || ts >= curCutoff.getTime()) continue;
    var code = row[3], type = row[6], qty = Number(row[7]) || 0;
    if (type === '入荷') {
      incomingMap[code] = (incomingMap[code] || 0) + qty;
    } else if (type === '廃棄') {
      disposalMap[code] = (disposalMap[code] || 0) + qty;
    } else if (type === '棚卸') {
      var eventKey = String(ts) + '||' + row[2];
      if (!stocktakeEventMap[eventKey]) {
        stocktakeEventMap[eventKey] = { timestamp: row[0], staffName: row[2], itemCount: 0 };
      }
      stocktakeEventMap[eventKey].itemCount++;
    }
  }

  var stocktakeEvents = Object.keys(stocktakeEventMap)
    .map(function (k) { return stocktakeEventMap[k]; })
    .sort(function (a, b) { return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime(); });
  if (stocktakeEvents.length) {
    stocktakeEvents[stocktakeEvents.length - 1].isLatest = true;
  }

  var products = listProducts_(session, store).products;
  var items = products.map(function (p) {
    var prevStock = prevMap[p.code] || 0;
    var incoming = incomingMap[p.code] || 0;
    var disposal = disposalMap[p.code] || 0;
    var currentStock = hasKey_(curMap, p.code) ? curMap[p.code] : prevStock + incoming - disposal;
    var expected = prevStock + incoming - disposal;
    return {
      code: p.code, brand: p.brand, name: p.name, colorNo: p.colorNo,
      prevStock: prevStock, incoming: incoming, disposal: disposal,
      currentStock: currentStock, diff: currentStock - expected
    };
  });

  return {
    store: store,
    month: month,
    previousMonth: previousMonthString_(month),
    items: items,
    stocktakeEvents: stocktakeEvents,
    approval: getApprovalStatus_(store, month),
    monthlyStatus: getMonthlyStocktakeStatus_(store, month)
  };
}

/** map[key]の存在確認(値が0だとfalsyになってしまうため、存在確認だけの用途に使う)。 */
function hasKey_(map, key) {
  return Object.prototype.hasOwnProperty.call(map, key);
}

/** 本社が、店舗×年月の棚卸内容を確認済みとして記録する。差し戻し中だった場合はその内容をクリアする。 */
function approveStocktake_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store || !p.month) throw new Error('店舗と対象月を指定してください');

  var sheet = getSheet_(SHEET_APPROVALS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === p.store && data[i][1] === p.month) {
      sheet.getRange(i + 1, 3).setValue(session.username);
      sheet.getRange(i + 1, 4).setValue(new Date());
      sheet.getRange(i + 1, 5).setValue('承認済み');
      sheet.getRange(i + 1, 6).setValue('');
      sheet.getRange(i + 1, 7).setValue('');
      return { store: p.store, month: p.month, approver: session.username };
    }
  }
  sheet.appendRow([p.store, p.month, session.username, new Date(), '承認済み', '', '']);
  return { store: p.store, month: p.month, approver: session.username };
}

/**
 * 本社が、店舗×年月の棚卸を差し戻す(店舗が再度カウントできるようにする)。
 * 承認済みの月はsubmitStocktake_でロックされるため、修正が必要なときはここから解除する。
 */
function rejectStocktake_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store || !p.month) throw new Error('店舗と対象月を指定してください');

  var sheet = getSheet_(SHEET_APPROVALS);
  var data = sheet.getDataRange().getValues();
  var reason = p.reason || '';
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === p.store && data[i][1] === p.month) {
      sheet.getRange(i + 1, 3).setValue('');
      sheet.getRange(i + 1, 4).setValue('');
      sheet.getRange(i + 1, 5).setValue('差し戻し');
      sheet.getRange(i + 1, 6).setValue(reason);
      sheet.getRange(i + 1, 7).setValue(now);
      return { store: p.store, month: p.month, status: '差し戻し' };
    }
  }
  sheet.appendRow([p.store, p.month, '', '', '差し戻し', reason, now]);
  return { store: p.store, month: p.month, status: '差し戻し' };
}

// ---- 在庫のリセット(本社限定・取り消し不可の操作) ----
// どちらもパスワードの再検証(verifyOwnPassword_)を必須にしている。
// 実体としては、対象の店舗×商品に対して数量0の「棚卸」を記録することで在庫を0に戻す。

/**
 * 入力ミスやシステムの不具合などで実際と違う数になった在庫を、正しい数に直接書き換える。
 * 種別「調整」として記録するため、棚卸承認画面の棚卸実施回数には数えられない。
 */
function adjustProductStock_(session, p) {
  requireRole_(session, ['hq']);
  verifyOwnPassword_(session, p.password);
  if (!p.store || !p.code) throw new Error('店舗と商品コードを指定してください');
  var newStock = Number(p.newStock);
  if (p.newStock === undefined || p.newStock === '' || isNaN(newStock) || newStock < 0) {
    throw new Error('正しい在庫数を指定してください');
  }

  var product = lookupProduct_(p.store, p.code);
  if (!product) throw new Error('商品が見つかりません');

  appendLog_(p.store, session.username, product, '調整', newStock, p.memo || '本社による在庫調整');
  refreshSummary_();
  return { store: p.store, code: p.code, newStock: newStock };
}

/** カラージェルの劣化などで、特定商品1点だけの在庫を0にリセットする。 */
function resetProductStock_(session, p) {
  requireRole_(session, ['hq']);
  verifyOwnPassword_(session, p.password);
  if (!p.store || !p.code) throw new Error('店舗と商品コードを指定してください');

  var product = lookupProduct_(p.store, p.code);
  if (!product) throw new Error('商品が見つかりません');

  appendLog_(p.store, session.username, product, '棚卸', 0, '本社による個別リセット');
  refreshSummary_();
  return { store: p.store, code: p.code };
}

/** 指定した店舗(または全店舗)の在庫をすべて0にリセットする。 */
function resetStoreInventory_(session, p) {
  requireRole_(session, ['hq']);
  verifyOwnPassword_(session, p.password);
  if (!p.store) throw new Error('店舗を指定してください(全店舗の場合は "ALL")');

  var all = computeAllSummary_();
  var targets = p.store === 'ALL' ? all : all.filter(function (e) { return e.store === p.store; });

  targets.forEach(function (e) {
    appendLog_(e.store, session.username, { code: e.code, name: e.name, brand: e.brand }, '棚卸', 0, '本社による一括リセット');
  });
  refreshSummary_();
  return { resetCount: targets.length };
}

function refreshSummary_() {
  var sheet = getSheet_(SHEET_SUMMARY);
  var all = computeAllSummary_();

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }
  if (!all.length) return;

  var now = new Date();
  var rows = all.map(function (e) {
    return [e.store, e.brand, e.name, e.code, e.current, e.lastStocktakeCount, e.current <= 0, now];
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * 「現在庫サマリ」シートを最新の状態に書き直す(参考用シートの手動更新)。
 * 入荷登録・廃棄登録では速度のためこれを自動実行しなくなったので、スプレッドシート上で
 * 最新の集計を見たいときは、Apps Scriptエディタからこの関数を手動で実行する。
 */
function refreshCurrentStockSummary() {
  refreshSummary_();
  Logger.log('現在庫サマリを更新しました。');
}
