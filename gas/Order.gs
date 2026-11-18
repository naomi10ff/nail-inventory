/**
 * SOFA STOCKから発注(2026-11-18〜)。
 *
 * お店がほしい商品のバーコードを読む → その店舗の発注表(タブ「発注入力」)に1行足す。
 * - その店舗の商品マスタにある商品だけ(新しい商品は、これまでどおり発注表にURLを貼る)。
 * - 商品マスタに品番(TAT)がある商品は、発注先「TAT」とTATの商品ページのURLも入れる
 *   (URLからTAT品番・棚卸しの商品名・確認の列が自動で出る)。
 * - 品番がない商品は品名だけ入れ、メモに「品番なし」と書く(本社が品番・URLを足す)。
 * - 同じ発注月で、まだ発注していない同じ商品の行があれば、新しい行は足さずに知らせる
 *   (お店が選べば、その行の個数を書きかえる)。
 * 書くのは青い列(A〜H)だけ。灰色(自動)・ベージュ(本社)の列はさわらない。
 */

var ORDER_MEMO_MARK = 'SOFA STOCKから';
var ORDER_MEMO_NO_NUMBER = 'SOFA STOCKから（品番なし：本社で入れます）';

function tatUrlFromNumber_(no) {
  return 'https://www.nailtat.com/tatweb/item/' + no.slice(0, 3) + '/ItemDetail.do?id=' + no;
}

function orderItemLabel_(product) {
  var name = String(product.name || '');
  if (product.colorNo && name.indexOf(String(product.colorNo)) === -1) name += '（' + product.colorNo + '）';
  return (product.brand ? product.brand + ' ' : '') + name;
}

function orderSheet_(store) {
  var id = DELIVERY_ORDER_SHEETS[store];
  if (!id) throw new Error('この店舗の発注表が見つかりません');
  var sheet = SpreadsheetApp.openById(id).getSheetByName(DELIVERY_ORDER_TAB);
  if (!sheet) throw new Error('発注表のタブ「' + DELIVERY_ORDER_TAB + '」が見つかりません');
  return sheet;
}

/** 発注表の青い列(A〜H)と本社の列(M・N)を読む。lastFilled＝A〜Hに何か入っている最後の行。 */
function readOrderRows_(sheet) {
  var last = sheet.getLastRow();
  var rows = [], lastFilled = DELIVERY_FIRST_ROW - 1;
  if (last < DELIVERY_FIRST_ROW) return { rows: rows, lastFilled: lastFilled };
  var values = sheet.getRange(DELIVERY_FIRST_ROW, 1, last - DELIVERY_FIRST_ROW + 1, 14).getValues();
  values.forEach(function (r, i) {
    var filled = r.slice(0, 8).some(function (v) { return v !== '' && v != null; });
    if (filled) lastFilled = DELIVERY_FIRST_ROW + i;
    rows.push({
      row: DELIVERY_FIRST_ROW + i, filled: filled,
      inputDate: deliveryDate_(r[0]), who: String(r[1] || ''), url: String(r[4] || '').trim(),
      qty: Number(r[5]) || 0, name: String(r[6] || '').trim(), memo: String(r[7] || ''),
      itemNumber: r[8] === '' || r[8] == null ? '' : String(r[8]).replace(/\.0$/, ''),
      orderDate: r[12], state: String(r[13] || '').trim()
    });
  });
  return { rows: rows, lastFilled: lastFilled };
}

/** まだ本社が発注していない、今の発注月の行。 */
function openOrderRows_(rows, month) {
  return rows.filter(function (r) {
    if (!r.filled || r.state || r.orderDate) return false;
    return deliveryOrderMonth_(r.inputDate || new Date(), null) === month;
  });
}

/** 今の発注月(7日締め)と、その月にSOFA STOCKから入れた・まだ発注していない行の一覧。 */
function getOrderList_(session, p) {
  requireRole_(session, ['hq', 'store']);
  var store = session.role === 'hq' ? p.store : session.store;
  if (!store) throw new Error('店舗を指定してください');
  var month = deliveryOrderMonth_(new Date(), null);
  var data = readOrderRows_(orderSheet_(store));
  var list = openOrderRows_(data.rows, month).map(function (r) {
    return { row: r.row, name: r.name || r.url, qty: r.qty, who: r.who, fromApp: r.memo.indexOf(ORDER_MEMO_MARK) === 0 };
  });
  return { store: store, month: month, lines: list };
}

/**
 * 1商品を発注表に入れる。p: code, quantity, staffName, replaceRow(同じ商品の行の個数を書きかえるとき)
 * 戻り値 duplicate:true のときは書いていない(お店に確認してから replaceRow つきで呼び直す)。
 */
function addOrderLine_(session, p) {
  requireRole_(session, ['hq', 'store']);
  var store = session.role === 'hq' ? p.store : session.store;
  if (!store) throw new Error('店舗を指定してください');
  var product = lookupProduct_(store, p.code);
  if (!product) throw new Error('この店舗の商品リストにない商品です。新しい商品は、これまでどおり発注表にURLを貼ってください');
  var quantity = Math.floor(Number(p.quantity) || 0);
  if (quantity < 1) throw new Error('個数を1以上にしてください');
  // 発注表の「入力者」は設定タブのスタッフ一覧しか入らない(入力規則)。本社アカウントでも選んだ担当者名を使う。
  var staff = String(p.staffName || recorderName_(session, p)).trim();
  var no = String(product.itemNumber == null ? '' : product.itemNumber).trim().replace(/\.0$/, '');
  if (!/^\d{4,7}$/.test(no)) no = '';
  var label = orderItemLabel_(product);
  var url = no ? tatUrlFromNumber_(no) : '';

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = orderSheet_(store);
    var month = deliveryOrderMonth_(new Date(), null);
    var data = readOrderRows_(sheet);
    var open = openOrderRows_(data.rows, month);

    if (p.replaceRow) {
      var target = open.filter(function (r) { return r.row === Number(p.replaceRow); })[0];
      if (!target) throw new Error('その行は本社が発注を始めたため、書きかえられません。本社に連絡してください');
      sheet.getRange(target.row, 6).setValue(quantity);
      return { product: product, label: label, quantity: quantity, row: target.row, replaced: true, hasNumber: !!no };
    }

    var same = open.filter(function (r) {
      return (no && r.itemNumber === no) || deliveryNorm_(r.name) === deliveryNorm_(label);
    })[0];
    if (same) {
      return { duplicate: true, product: product, label: label, row: same.row, existingQty: same.qty, existingWho: same.who };
    }

    var row = data.lastFilled + 1;
    var today = new Date();
    var memo = no ? ORDER_MEMO_MARK : ORDER_MEMO_NO_NUMBER;
    var staffList = sheet.getParent().getSheetByName('設定').getRange('A2:A50').getValues()
      .map(function (r) { return String(r[0]).trim(); });
    if (staffList.indexOf(staff) === -1) { memo += '（入力：' + staff + '）'; staff = ''; }
    try {
      sheet.getRange(row, 1, 1, 8).setValues([[
        new Date(today.getFullYear(), today.getMonth(), today.getDate()),
        staff, 'ネイル商材', no ? 'TAT' : '', url, quantity, label, memo
      ]]);
      SpreadsheetApp.flush();
    } catch (err) {
      sheet.getRange(row, 1, 1, 8).clearContent(); // 途中まで書けた行を残さない
      throw new Error('発注表に入れられませんでした。本社に連絡してください（' + err.message + '）');
    }
    return { product: product, label: label, quantity: quantity, row: row, month: month, hasNumber: !!no };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 送る前のリストをまとめて発注表に入れる(「発注表に送る」ボタン)。p: items[{code, quantity}], staffName
 * 1商品ずつ結果を返す：added(入れた)／duplicate(今月の発注表にもう入っていたので入れていない)／error。
 */
function addOrderLines_(session, p) {
  requireRole_(session, ['hq', 'store']);
  var store = session.role === 'hq' ? p.store : session.store;
  if (!store) throw new Error('店舗を指定してください');
  var items = Array.isArray(p.items) ? p.items : [];
  if (!items.length) throw new Error('送る商品がありません');
  var staff = String(p.staffName || recorderName_(session, p)).trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = orderSheet_(store);
    var month = deliveryOrderMonth_(new Date(), null);
    var data = readOrderRows_(sheet);
    var open = openOrderRows_(data.rows, month);
    var staffList = sheet.getParent().getSheetByName('設定').getRange('A2:A50').getValues()
      .map(function (r) { return String(r[0]).trim(); });
    var staffCell = staffList.indexOf(staff) === -1 ? '' : staff;
    var today = new Date();
    var day = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var row = data.lastFilled + 1;
    var results = [];

    items.forEach(function (it) {
      var res = { code: String(it.code || '') };
      try {
        var product = lookupProduct_(store, it.code);
        if (!product) throw new Error('この店舗の商品リストにない商品です');
        var quantity = Math.floor(Number(it.quantity) || 0);
        if (quantity < 1) throw new Error('個数が0です');
        var no = String(product.itemNumber == null ? '' : product.itemNumber).trim().replace(/\.0$/, '');
        if (!/^\d{4,7}$/.test(no)) no = '';
        var label = orderItemLabel_(product);
        res.label = label; res.quantity = quantity;
        var same = open.filter(function (r) {
          return (no && r.itemNumber === no) || deliveryNorm_(r.name) === deliveryNorm_(label);
        })[0];
        if (same) {
          res.status = 'duplicate'; res.row = same.row; res.existingQty = same.qty; res.existingWho = same.who;
          results.push(res);
          return;
        }
        var memo = no ? ORDER_MEMO_MARK : ORDER_MEMO_NO_NUMBER;
        if (!staffCell && staff) memo += '（入力：' + staff + '）';
        try {
          sheet.getRange(row, 1, 1, 8).setValues([[day, staffCell, 'ネイル商材', no ? 'TAT' : '',
            no ? tatUrlFromNumber_(no) : '', quantity, label, memo]]);
        } catch (err) {
          sheet.getRange(row, 1, 1, 8).clearContent();
          throw new Error('発注表に入れられませんでした（' + err.message + '）');
        }
        open.push({ row: row, itemNumber: no, name: label, qty: quantity, who: staffCell || staff });
        res.status = 'added'; res.row = row;
        row++;
      } catch (e) {
        res.status = 'error'; res.error = e.message;
      }
      results.push(res);
    });
    SpreadsheetApp.flush();
    return { month: month, results: results };
  } finally {
    lock.releaseLock();
  }
}
