/**
 * 発注と入荷の紐づけ(2026-11-17〜)。
 *
 * 各店舗の「発注表」(Googleスプレッドシート、タブ「発注入力」)のうち、本社が状態を
 * 「発注済み」「欠品・入荷待ち」にした行を「届く予定」として扱う。
 * - TATの商品は、発注表のTAT品番(URLの ?id=)と商品マスタの「品番」で棚卸しの商品とつなぐ。
 * - 届いた数は、取引ログの「入荷」から数える(発注日以降・古い発注から順に割り当てる)。
 * - 棚卸しの商品とつながらない行(楽天・Amazon・新しい商品など)は、お店が「届いた」を押して記録する。
 * - 足りない商品の理由(欠品・入荷待ち／数がちがう／別の商品が届いた／破損していた／その他)と
 *   追記は、シート「入荷の報告」に1行ずつ残す(同じ行の報告は新しいものが有効)。
 * 発注表そのものは読むだけで、書きかえない。
 */

var DELIVERY_ORDER_SHEETS = {
  '西大寺宝来店': '1tmlgW2hc6Gtd58nM1iZ1614ip9L6FVBLv1EL4pSmsyA',
  '生駒店': '1hzRAMk71EOGmwREnhCwhlC8YU9plD3RuqmYNznJ86mc',
  '木津店': '1bpqY4fnHYgzy3wMbN2t73tlw0qDzRdSni5IPAvQGSws'
};
var DELIVERY_ORDER_TAB = '発注入力';
var DELIVERY_FIRST_ROW = 6; // 1〜5行目は見出し
var DELIVERY_OPEN_STATES = ['発注済み', '欠品・入荷待ち'];
var DELIVERY_REASONS = ['欠品・入荷待ち', '数がちがう', '別の商品が届いた', '破損していた', 'その他'];
var SHEET_DELIVERY_REPORTS = '入荷の報告';
var DELIVERY_REPORT_HEADERS = ['記録日時', '店舗', '行のキー', '発注月', '商品', '内容', '個数', '追記', '記録者'];
var DELIVERY_ARRIVED = '届いた（手で記録）';

function deliveryReportSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheet_(ss, SHEET_DELIVERY_REPORTS, DELIVERY_REPORT_HEADERS);
  return ss.getSheetByName(SHEET_DELIVERY_REPORTS);
}

function deliveryDate_(v) {
  if (v instanceof Date) return v;
  if (v === '' || v == null) return null;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/** 7日締め：1〜7日に入力した分はその月、8日以降は次の月の発注。 */
function deliveryOrderMonth_(inputDate, orderDate) {
  var d = inputDate || orderDate;
  if (!d) return '';
  var y = d.getFullYear(), m = d.getMonth() + (d.getDate() <= 7 ? 0 : 1);
  return Utilities.formatDate(new Date(y, m, 1), 'Asia/Tokyo', 'yyyy-MM');
}

function deliveryNorm_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
}

/** 1店舗の発注表から、届く予定の行を読む(状態が「発注済み」「欠品・入荷待ち」の行だけ)。 */
function readOrderLines_(store) {
  var id = DELIVERY_ORDER_SHEETS[store];
  if (!id) return [];
  var sheet = SpreadsheetApp.openById(id).getSheetByName(DELIVERY_ORDER_TAB);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < DELIVERY_FIRST_ROW) return [];
  var values = sheet.getRange(DELIVERY_FIRST_ROW, 1, last - DELIVERY_FIRST_ROW + 1, 14).getValues();
  var lines = [], seen = {};
  values.forEach(function (r, i) {
    var state = String(r[13] || '').trim();
    if (DELIVERY_OPEN_STATES.indexOf(state) === -1) return;
    var qty = Number(r[5]) || 0;
    if (qty <= 0) return;
    var inputDate = deliveryDate_(r[0]), orderDate = deliveryDate_(r[12]);
    var url = String(r[4] || '').trim(), name = String(r[6] || '').trim();
    var base = [store, orderDate ? Utilities.formatDate(orderDate, 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      deliveryNorm_(url) || deliveryNorm_(name)].join('|');
    seen[base] = (seen[base] || 0) + 1;
    lines.push({
      key: base + '|' + seen[base],
      store: store,
      row: DELIVERY_FIRST_ROW + i,
      month: deliveryOrderMonth_(inputDate, orderDate),
      orderDate: orderDate,
      type: String(r[2] || ''),
      vendor: String(r[3] || ''),
      url: url,
      qty: qty,
      itemName: name,
      itemNumber: r[8] === '' || r[8] == null ? '' : String(r[8]).replace(/\.0$/, ''),
      stocktakeName: String(r[9] || ''),
      state: state
    });
  });
  return lines;
}

/** 商品マスタの品番 → 棚卸しの商品(その店舗だけ)。 */
function itemNumberMap_(store) {
  var data = getSheet_(SHEET_PRODUCTS).getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== store) continue;
    var no = String(data[i][10] == null ? '' : data[i][10]).trim().replace(/\.0$/, '');
    if (!/^\d{4,7}$/.test(no) || map[no]) continue;
    map[no] = { code: String(data[i][1]), name: data[i][2], brand: data[i][3], colorNo: data[i][9] };
  }
  return map;
}

/** 入荷の報告(理由・手で記録した「届いた」)を行のキーごとにまとめる。 */
function deliveryReportsByKey_(store) {
  var data = deliveryReportSheet_().getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    if (store && data[i][1] !== store) continue;
    var key = String(data[i][2]);
    var o = out[key] || (out[key] = { reason: null, arrivedManual: 0 });
    var kind = String(data[i][5]);
    if (kind === DELIVERY_ARRIVED) {
      o.arrivedManual += Number(data[i][6]) || 0;
    } else {
      o.reason = { reason: kind, memo: String(data[i][7] || ''), at: data[i][0], by: String(data[i][8] || '') };
    }
  }
  return out;
}

/**
 * 店舗の届く予定に、届いた数・理由をつけて返す。
 * 届いた数は、取引ログの「入荷」を発注日以降のものだけ、古い発注から順に割り当てる。
 */
function buildDeliveryLines_(store) {
  var lines = readOrderLines_(store);
  if (!lines.length) return [];
  var numMap = itemNumberMap_(store);
  var reports = deliveryReportsByKey_(store);

  lines.forEach(function (l) {
    var p = l.itemNumber ? numMap[l.itemNumber] : null;
    l.code = p ? p.code : '';
    l.displayName = p ? (p.name + (p.colorNo ? '（' + p.colorNo + '）' : '')) : (l.itemName || l.stocktakeName || '（品名なし）');
    l.brand = p ? (p.brand || '') : '';
    l.got = 0;
  });

  // コードごとに、発注日の古い順に並べておく
  var byCode = {};
  lines.forEach(function (l) { if (l.code) (byCode[l.code] = byCode[l.code] || []).push(l); });
  Object.keys(byCode).forEach(function (c) {
    byCode[c].sort(function (a, b) { return (a.orderDate ? a.orderDate.getTime() : 0) - (b.orderDate ? b.orderDate.getTime() : 0); });
  });

  if (Object.keys(byCode).length) {
    var log = getSheet_(SHEET_LOG).getDataRange().getValues();
    // 入荷(取り消しのマイナスも含む)をコードごとに時刻順で集める
    var incoming = {};
    for (var i = 1; i < log.length; i++) {
      var r = log[i];
      if (r[1] !== store || r[6] !== '入荷') continue;
      var code = String(r[3]);
      if (!byCode[code]) continue;
      var ts = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      (incoming[code] = incoming[code] || []).push({ ts: ts, qty: Number(r[7]) || 0 });
    }
    Object.keys(byCode).forEach(function (code) {
      var list = byCode[code];
      var earliest = list[0].orderDate ? new Date(list[0].orderDate.getFullYear(), list[0].orderDate.getMonth(), list[0].orderDate.getDate()).getTime() : 0;
      var total = 0;
      (incoming[code] || []).forEach(function (e) { if (e.ts >= earliest) total += e.qty; });
      total = Math.max(0, total);
      list.forEach(function (l) {
        var take = Math.min(total, l.qty);
        l.got = take;
        total -= take;
      });
    });
  }

  lines.forEach(function (l) {
    var rep = reports[l.key];
    if (!l.code && rep) l.got = Math.min(l.qty, rep.arrivedManual);
    l.reason = rep && rep.reason ? rep.reason : null;
    l.done = l.got >= l.qty;
    l.orderDate = l.orderDate ? Utilities.formatDate(l.orderDate, 'Asia/Tokyo', 'M/d') : '';
    if (l.reason && l.reason.at instanceof Date) l.reason.at = Utilities.formatDate(l.reason.at, 'Asia/Tokyo', 'M/d HH:mm');
  });
  return lines;
}

/** お店の入荷画面用：まだ全部届いていない発注月の行(その月の届いた行も含む)。 */
function getDeliveryPlan_(session, p) {
  requireRole_(session, ['hq', 'store']);
  var store = session.role === 'hq' ? (p.store || session.store) : session.store;
  if (!store) throw new Error('店舗を指定してください');
  var lines = buildDeliveryLines_(store);
  var openMonths = {};
  lines.forEach(function (l) { if (!l.done) openMonths[l.month] = true; });
  var shown = lines.filter(function (l) { return openMonths[l.month]; });
  shown.sort(function (a, b) {
    if (a.month !== b.month) return a.month < b.month ? -1 : 1;
    if (a.done !== b.done) return a.done ? 1 : -1;
    return 0;
  });
  return { store: store, lines: shown, reasons: DELIVERY_REASONS };
}

/**
 * お店からの報告。items: [{ key, reason, memo }] または [{ key, arrived: 個数 }]。
 * 理由は5つのどれか・追記は必須。
 */
function reportDelivery_(session, p) {
  requireRole_(session, ['hq', 'store']);
  var store = session.role === 'hq' ? (p.store || session.store) : session.store;
  if (!store) throw new Error('店舗を指定してください');
  var items = p.items || [];
  if (!items.length) return { saved: 0 };
  var lines = buildDeliveryLines_(store);
  var byKey = {};
  lines.forEach(function (l) { byKey[l.key] = l; });
  var who = recorderName_(session, p);
  var now = new Date();
  var rows = [];
  items.forEach(function (it) {
    var l = byKey[it.key];
    if (!l) throw new Error('発注表に見つからない行があります。画面を開き直してください');
    if (it.arrived) {
      rows.push([now, store, l.key, l.month, l.displayName, DELIVERY_ARRIVED, Number(it.arrived) || 0, '', who]);
      return;
    }
    if (DELIVERY_REASONS.indexOf(it.reason) === -1) throw new Error('理由を選んでください：' + l.displayName);
    var memo = String(it.memo || '').trim();
    if (!memo) throw new Error('追記を書いてください：' + l.displayName);
    rows.push([now, store, l.key, l.month, l.displayName, it.reason, l.qty - l.got, memo, who]);
  });
  var sheet = deliveryReportSheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { saved: rows.length };
}

/** 本社用：発注月を指定して、3店舗の届く予定と届いた数・理由。 */
function getDeliveryOverview_(session, p) {
  requireRole_(session, ['hq']);
  var byStore = {}, errors = {}, months = {};
  Object.keys(DELIVERY_ORDER_SHEETS).forEach(function (store) {
    try { byStore[store] = buildDeliveryLines_(store); } catch (e) { errors[store] = e.message; byStore[store] = []; }
    byStore[store].forEach(function (l) { if (l.month) months[l.month] = true; });
  });
  var monthList = Object.keys(months).sort().reverse();
  // 指定がなければ、まだ届いていない行がある一番古い月(なければ一番新しい月)
  var month = p.month || '';
  if (!month) {
    var open = [];
    Object.keys(byStore).forEach(function (s) { byStore[s].forEach(function (l) { if (!l.done && l.month) open.push(l.month); }); });
    month = open.length ? open.sort()[0] : (monthList[0] || '');
  }
  var stores = [], all = [];
  Object.keys(byStore).forEach(function (store) {
    if (errors[store]) { stores.push({ store: store, error: errors[store], total: 0, got: 0, open: 0, noReason: 0 }); return; }
    var mine = byStore[store].filter(function (l) { return l.month === month; });
    stores.push({
      store: store,
      total: mine.length,
      got: mine.filter(function (l) { return l.done; }).length,
      open: mine.filter(function (l) { return !l.done; }).length,
      noReason: mine.filter(function (l) { return !l.done && !l.reason; }).length
    });
    all = all.concat(mine);
  });
  return { month: month, months: monthList, stores: stores, lines: all };
}
