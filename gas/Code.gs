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
      return lookupProduct_(p.code);
    case 'registerProduct':
      return registerProduct_(validateToken_(p.token), p);
    case 'recordIncoming':
      return recordIncoming_(validateToken_(p.token), p);
    case 'recordUsed':
      return recordUsed_(validateToken_(p.token), p);
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
    case 'listStores':
      return { stores: STORES };
    case 'listProducts':
      return listProducts_(validateToken_(p.token));
    case 'deleteProduct':
      return deleteProduct_(validateToken_(p.token), p);
    case 'getLogEntries':
      return getLogEntries_(validateToken_(p.token), p.store, p.limit);
    case 'createAccount':
      return createAccount_(validateToken_(p.token), p);
    case 'resetPassword':
      return resetPassword_(validateToken_(p.token), p);
    case 'ocrProductLabel':
      return ocrProductLabel_(validateToken_(p.token), p);
    case 'updateProduct':
      return updateProduct_(validateToken_(p.token), p);
    case 'getBrandList':
      return getBrandList_(validateToken_(p.token));
    case 'addBrand':
      return addBrand_(validateToken_(p.token), p);
    case 'resetProductStock':
      return resetProductStock_(validateToken_(p.token), p);
    case 'resetStoreInventory':
      return resetStoreInventory_(validateToken_(p.token), p);
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

function lookupProduct_(code) {
  if (!code) return null;
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(code)) {
      return {
        code: data[i][0],
        name: data[i][1],
        brand: data[i][2],
        category: data[i][3],
        maker: data[i][4],
        unit: data[i][5],
        colorNo: data[i][8]
      };
    }
  }
  return null;
}

/**
 * 新規商品登録。バーコードがある商品はそのコードで、ない商品は code を空にして
 * 呼ぶと自前QRコードを発行する。店舗・本社どちらのアカウントからも呼べるが、
 * 既存商品の上書きはできない(コードが重複していればエラーになる)。
 * 既存商品の修正は updateProduct_ (本社のみ)を使う。
 */
function registerProduct_(session, p) {
  if (!p.name) throw new Error('品名を入力してください');

  var code = p.code;
  if (!code) {
    code = 'QR-' + Utilities.getUuid().split('-')[0].toUpperCase();
  } else if (lookupProduct_(code)) {
    throw new Error('このコードは既に登録されています');
  }

  var sheet = getSheet_(SHEET_PRODUCTS);
  sheet.appendRow([
    code, p.name, p.brand || '', p.category || '', p.maker || '',
    p.unit || '本', p.memo || '', new Date(), p.colorNo || ''
  ]);
  return { code: code, name: p.name, brand: p.brand || '', colorNo: p.colorNo || '' };
}

/** 既存商品の品名・ブランド・カテゴリー・カラーNOの修正。本社アカウントのみ。 */
function updateProduct_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.code) throw new Error('コードが指定されていません');

  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.code)) {
      var row = i + 1;
      sheet.getRange(row, 2).setValue(p.name || '');
      sheet.getRange(row, 3).setValue(p.brand || '');
      sheet.getRange(row, 4).setValue(p.category || '');
      sheet.getRange(row, 9).setValue(p.colorNo || '');
      refreshSummary_(); // ブランド名等の変更をサマリ表示にも反映する
      return { code: p.code };
    }
  }
  throw new Error('商品が見つかりません');
}

function listProducts_(session) {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      code: data[i][0], name: data[i][1], brand: data[i][2],
      category: data[i][3], maker: data[i][4], unit: data[i][5],
      colorNo: data[i][8]
    });
  }
  return { products: list };
}

/** 既存商品の削除。本社アカウントのみ。 */
function deleteProduct_(session, p) {
  requireRole_(session, ['hq']);
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.code)) {
      sheet.deleteRow(i + 1);
      return { deleted: p.code };
    }
  }
  throw new Error('商品が見つかりません');
}

// ---- ブランドマスタ ----
// 本社が追加登録したブランド名が、店舗側の新規商品登録フォームのプルダウンに反映される。

function getBrandList_(session) {
  var sheet = getSheet_(SHEET_BRANDS);
  var data = sheet.getDataRange().getValues();
  var brands = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) brands.push(data[i][0]);
  }
  return { brands: brands };
}

function addBrand_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.name) throw new Error('ブランド名を入力してください');
  var existing = getBrandList_(session).brands;
  if (existing.indexOf(p.name) !== -1) throw new Error('このブランド名は既に登録されています');
  getSheet_(SHEET_BRANDS).appendRow([p.name]);
  return { name: p.name };
}

// ---- 取引ログ(入荷・使用済・棚卸) ----

function appendLog_(store, staffName, product, type, quantity, memo) {
  var sheet = getSheet_(SHEET_LOG);
  sheet.appendRow([new Date(), store, staffName || '', product.code, product.name, product.brand, type, quantity, memo || '']);
}

function recordIncoming_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);
  var product = lookupProduct_(p.code);
  if (!product) throw new Error('商品が見つかりません。先に新規商品登録をしてください');
  var quantity = Number(p.quantity) || 1;
  appendLog_(store, p.staffName, product, '入荷', quantity, p.memo);
  refreshSummary_();
  return { product: product, quantity: quantity };
}

function recordUsed_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);
  var product = lookupProduct_(p.code);
  if (!product) throw new Error('商品が見つかりません');
  appendLog_(store, p.staffName, product, '使用済', 1, p.memo);
  refreshSummary_();
  return { product: product };
}

/** items: [{ code, count }, ...] 棚卸で連続スキャンした結果をまとめて送信する。 */
function submitStocktake_(session, p) {
  var store = session.role === 'hq' ? p.store : session.store;
  requireStoreAccess_(session, store);

  var items = p.items || [];
  var recorded = [];
  var unknownCodes = [];
  items.forEach(function (item) {
    var product = lookupProduct_(item.code);
    if (!product) {
      unknownCodes.push(item.code);
      return;
    }
    appendLog_(store, p.staffName, product, '棚卸', item.count, '');
    recorded.push({ code: item.code, name: product.name, brand: product.brand, count: item.count });
  });
  refreshSummary_();

  return {
    recorded: recorded,
    unknownCodes: unknownCodes,
    summary: getInventorySummary_(session, store)
  };
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

function getLogEntries_(session, storeParam, limit) {
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
// 直近の棚卸カウントをその時点の実在庫として採用し、それ以降の入荷(+)・使用済(-)を
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
    } else if (type === '使用済') {
      entry.current -= qty;
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

/** コード(バーコード)→カラーNOの対応表。在庫一覧の検索用にカラーNOも一緒に返すために使う。 */
function getProductColorNoMap_() {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[data[i][0]] = data[i][8] || '';
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
      colorNo: colorNoMap[e.code] || '',
      currentStock: e.current,
      outOfStock: e.current <= 0,
      lastStocktakeCount: e.lastStocktakeCount
    };
  });

  items.sort(function (a, b) {
    if (a.store !== b.store) return a.store < b.store ? -1 : 1;
    if (a.brand !== b.brand) return (a.brand || '').localeCompare(b.brand || '', 'ja');
    return (a.name || '').localeCompare(b.name || '', 'ja');
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

  var product = lookupProduct_(p.code);
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

// ---- 在庫のリセット(本社限定・取り消し不可の操作) ----
// どちらもパスワードの再検証(verifyOwnPassword_)を必須にしている。
// 実体としては、対象の店舗×商品に対して数量0の「棚卸」を記録することで在庫を0に戻す。

/** カラージェルの劣化などで、特定商品1点だけの在庫を0にリセットする。 */
function resetProductStock_(session, p) {
  requireRole_(session, ['hq']);
  verifyOwnPassword_(session, p.password);
  if (!p.store || !p.code) throw new Error('店舗と商品コードを指定してください');

  var product = lookupProduct_(p.code);
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
