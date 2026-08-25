/* eslint-disable no-console */
// ここをGASのWebアプリURL(.../exec)に差し替えてください
const CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbwGwA8PfZjVi_IgDPvjcQkJIVSxdimPasRP1eej312KVBYG-S7q8mJBZS_3QOPrc4IY/exec'
};

// 店舗用・本社用の両方の画面をこの1ファイルで扱う。ログイン時のroleで
// 表示する画面群を振り分ける(店舗用の状態と本社用の状態を1つのstateにまとめている)。
const state = {
  token: localStorage.getItem('token') || null,
  store: localStorage.getItem('store') || null,
  role: localStorage.getItem('role') || null,
  staffName: null,
  tally: {}, // code -> { code, name, brand, count } (棚卸で使用)
  stocktakeMode: 'overwrite', // 'add'(既存カウントに加算)または'overwrite'(数え直し)
  stores: [] // 本社ダッシュボードで使用
};

// 棚卸の途中経過(state.tally)は、送信するまでは画面上のメモリ上にしかなかったため、
// アプリを閉じる(ホーム画面アプリをタスクスワイプで消す等)とそれまでスキャンした分が
// 全部消えてしまっていた。1件スキャン・修正するたびにlocalStorageへも保存しておき、
// 棚卸に入るときに残っていれば復元できるようにする。
function stocktakeDraftKey_() {
  return 'stocktakeDraft_' + (state.store || '');
}
function saveStocktakeDraft() {
  try {
    localStorage.setItem(stocktakeDraftKey_(), JSON.stringify({ tally: state.tally, mode: state.stocktakeMode }));
  } catch (e) { /* 保存に失敗しても棚卸自体は続行できるようにする */ }
}
function loadStocktakeDraft() {
  try {
    const raw = localStorage.getItem(stocktakeDraftKey_());
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}
function clearStocktakeDraft() {
  try {
    localStorage.removeItem(stocktakeDraftKey_());
  } catch (e) { /* 無視 */ }
}

let scannerStocktake = null;
let scannerHqIncoming = null;
let scannerHqDisposal = null;
let scannerHqStockLookup = null;

// 以前はタブ復帰(visibilitychange)で自動的にスキャナーを再起動していたが、
// 新規商品登録フォームへの入力中にも発火してカメラが壊れる不具合が出たため廃止。
// カメラが固まったときは、各スキャン画面の手動「再起動」ボタンを使う。
let activeScannerCtx = null;
let scannerRestarting = false;

// ---- API通信 ----
// GAS Web AppはCORSのpreflightに対応していないため、
// Content-Type を text/plain にして単純リクエストとして送る。
async function apiCall(action, payload) {
  const res = await fetch(CONFIG.API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, payload: Object.assign({ token: state.token }, payload || {}) })
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || '不明なエラーが発生しました');
  return json.data;
}

// ---- 画面切替 ----
function showScreen(id) {
  document.querySelectorAll('.screen').forEach((el) => el.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

function setText(id, text) {
  document.getElementById(id).textContent = text;
}

/**
 * 通信中はボタンを無効化して押せなくする。GASのWeb Appは1回の呼び出しに数秒
 * かかることがあり、無効化しないと「反応が無い」と思って連打され、二重送信に
 * つながることがある。押したときのラベルに変え、完了後は元のラベルに戻す。
 */
async function withButtonBusy(button, busyLabel, fn) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = busyLabel;
  try {
    return await fn();
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

// ---- CSVエクスポート ----
function csvEscape(value) {
  const str = String(value == null ? '' : value);
  return /[",\r\n]/.test(str) ? '"' + str.replace(/"/g, '""') + '"' : str;
}

function downloadCsv(filename, rows) {
  // 先頭にUTF-8のBOMを付けることで、Excel/Googleスプレッドシートで開いたときに
  // 日本語の文字化けを防ぐ。
  const csvContent = rows.map((row) => row.map(csvEscape).join(',')).join('\r\n');
  const blob = new Blob(['﻿' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---- 取り消せない操作の確認モーダル ----
// パスワード再入力+警告表示+実行ボタンの3点で、ワンクリックでの実行を防ぐ。
function openDangerModal(message, onConfirm) {
  const overlay = document.getElementById('danger-modal');
  const passwordInput = document.getElementById('danger-modal-password');
  const errorEl = document.getElementById('danger-modal-error');
  const confirmBtn = document.getElementById('danger-modal-confirm');
  const cancelBtn = document.getElementById('danger-modal-cancel');

  document.getElementById('danger-modal-message').textContent = message;
  passwordInput.value = '';
  errorEl.textContent = '';
  overlay.style.display = 'flex';

  function close() {
    overlay.style.display = 'none';
    confirmBtn.removeEventListener('click', onConfirmClick);
    cancelBtn.removeEventListener('click', onCancelClick);
  }
  async function onConfirmClick() {
    const password = passwordInput.value;
    if (!password) {
      errorEl.textContent = 'パスワードを入力してください';
      return;
    }
    confirmBtn.disabled = true;
    try {
      await onConfirm(password);
      close();
    } catch (e) {
      errorEl.textContent = e.message;
    } finally {
      confirmBtn.disabled = false;
    }
  }
  function onCancelClick() {
    close();
  }

  confirmBtn.addEventListener('click', onConfirmClick);
  cancelBtn.addEventListener('click', onCancelClick);
}

// ---- 総在庫一覧から、ブランド名・品名・カラーNO・カテゴリー・在庫数をまとめて直せるようにする ----
// パスワード再入力を必須にする点はdanger-modalと同じ。
async function loadAdjustBrandOptions(store) {
  const data = await apiCall('getBrandList', { store });
  const select = document.getElementById('adjust-modal-brand-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-brand-toggle-adjust').addEventListener('click', () => {
  document.getElementById('adjust-modal-brand-new-form').style.display = 'block';
});

function currentAdjustBrandValue() {
  const newForm = document.getElementById('adjust-modal-brand-new-form');
  const newValue = document.getElementById('adjust-modal-brand-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('adjust-modal-brand-select').value;
}

function resetAdjustBrandNewForm() {
  document.getElementById('adjust-modal-brand-new-form').style.display = 'none';
  document.getElementById('adjust-modal-brand-new').value = '';
}

function setAdjustBrandValue(brand) {
  resetAdjustBrandNewForm();
  const select = document.getElementById('adjust-modal-brand-select');
  if (!brand) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === brand);
  if (hasOption) {
    select.value = brand;
  } else {
    document.getElementById('adjust-modal-brand-new-form').style.display = 'block';
    document.getElementById('adjust-modal-brand-new').value = brand;
  }
}

async function loadAdjustCategoryOptions() {
  const data = await apiCall('getCategoryList', {});
  const select = document.getElementById('adjust-modal-category-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.categories.forEach((category) => {
    const opt = document.createElement('option');
    opt.value = category;
    opt.textContent = category;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-category-toggle-adjust').addEventListener('click', () => {
  document.getElementById('adjust-modal-category-new-form').style.display = 'block';
});

function currentAdjustCategoryValue() {
  const newForm = document.getElementById('adjust-modal-category-new-form');
  const newValue = document.getElementById('adjust-modal-category-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('adjust-modal-category-select').value;
}

function resetAdjustCategoryNewForm() {
  document.getElementById('adjust-modal-category-new-form').style.display = 'none';
  document.getElementById('adjust-modal-category-new').value = '';
}

function setAdjustCategoryValue(category) {
  resetAdjustCategoryNewForm();
  const select = document.getElementById('adjust-modal-category-select');
  if (!category) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === category);
  if (hasOption) {
    select.value = category;
  } else {
    document.getElementById('adjust-modal-category-new-form').style.display = 'block';
    document.getElementById('adjust-modal-category-new').value = category;
  }
}

async function openAdjustModal(item, onConfirm) {
  const overlay = document.getElementById('adjust-modal');
  const codeInput = document.getElementById('adjust-modal-code');
  const nameInput = document.getElementById('adjust-modal-name');
  const colorInput = document.getElementById('adjust-modal-colorno');
  const newStockInput = document.getElementById('adjust-modal-newstock');
  const memoInput = document.getElementById('adjust-modal-memo');
  const passwordInput = document.getElementById('adjust-modal-password');
  const errorEl = document.getElementById('adjust-modal-error');
  const confirmBtn = document.getElementById('adjust-modal-confirm');
  const cancelBtn = document.getElementById('adjust-modal-cancel');

  document.getElementById('adjust-modal-label').textContent = `${item.store} / コード: ${item.code}`;
  codeInput.value = item.code;
  await loadAdjustBrandOptions(item.store);
  setAdjustBrandValue(item.brand || '');
  nameInput.value = item.name || '';
  colorInput.value = item.colorNo || '';
  await loadAdjustCategoryOptions();
  setAdjustCategoryValue(item.category || '');
  newStockInput.value = item.currentStock;
  memoInput.value = '';
  passwordInput.value = '';
  errorEl.textContent = '';
  overlay.style.display = 'flex';

  function close() {
    overlay.style.display = 'none';
    confirmBtn.removeEventListener('click', onConfirmClick);
    cancelBtn.removeEventListener('click', onCancelClick);
  }
  async function onConfirmClick() {
    const newCode = codeInput.value.trim();
    const name = nameInput.value.trim();
    const newStock = newStockInput.value.trim();
    const password = passwordInput.value;
    if (!newCode) {
      errorEl.textContent = '商品コードを入力してください';
      return;
    }
    if (!name) {
      errorEl.textContent = '品名を入力してください';
      return;
    }
    if (newStock === '' || Number(newStock) < 0) {
      errorEl.textContent = '正しい在庫数を入力してください';
      return;
    }
    if (!password) {
      errorEl.textContent = 'パスワードを入力してください';
      return;
    }
    confirmBtn.disabled = true;
    try {
      await onConfirm({
        newCode, brand: currentAdjustBrandValue(), name, colorNo: colorInput.value.trim(),
        category: currentAdjustCategoryValue(), newStock, memo: memoInput.value.trim(), password
      });
      close();
    } catch (e) {
      errorEl.textContent = e.message;
    } finally {
      confirmBtn.disabled = false;
    }
  }
  function onCancelClick() {
    close();
  }

  confirmBtn.addEventListener('click', onConfirmClick);
  cancelBtn.addEventListener('click', onCancelClick);
}

// ---- ログイン(店舗・本社共通、roleで振り分け) ----
document.getElementById('btn-login').addEventListener('click', async () => {
  const username = document.getElementById('login-username').value.trim();
  const password = document.getElementById('login-password').value;
  const errEl = document.getElementById('login-error');
  errEl.textContent = '';
  try {
    const data = await apiCall('login', { username, password });
    state.token = data.token;
    state.store = data.store;
    state.role = data.role;
    localStorage.setItem('token', state.token);
    localStorage.setItem('store', state.store);
    localStorage.setItem('role', state.role);
    if (data.role === 'hq') {
      await enterDashboard();
    } else {
      await goToStaffSelect();
    }
  } catch (e) {
    errEl.textContent = e.message;
  }
});

[['btn-logout-1', 'screen-login'], ['btn-logout-2', 'screen-login'], ['hq-btn-logout', 'screen-login']].forEach(([btnId]) => {
  document.getElementById(btnId).addEventListener('click', logout);
});

function logout() {
  state.token = null;
  state.store = null;
  state.role = null;
  state.staffName = null;
  state.stores = [];
  localStorage.removeItem('token');
  localStorage.removeItem('store');
  localStorage.removeItem('role');
  document.getElementById('login-error').textContent = '';
  showScreen('screen-login');
}

// ==================== 店舗用 ====================

// ---- スタッフ選択 ----
// エラーを内部で握りつぶさず呼び出し元に投げる。ログインボタンのハンドラから
// 呼ばれた場合はそのtry/catchでログイン画面にエラー表示され、init()(再読み込み時の
// 自動ログイン)から呼ばれた場合はそのcatchでlogout()が呼ばれ、無効なセッションの
// まま古い店舗名だけが表示され続ける状態を防ぐ。
async function goToStaffSelect() {
  setText('staff-store-name', state.store);
  const errEl = document.getElementById('staff-error');
  errEl.textContent = '';
  const data = await apiCall('getStaffList', {});
  const select = document.getElementById('staff-select');
  select.innerHTML = '';
  data.staff.forEach((s) => {
    const opt = document.createElement('option');
    opt.value = s.name;
    opt.textContent = s.name;
    select.appendChild(opt);
  });
  showScreen('screen-staff');
}

document.getElementById('btn-add-staff').addEventListener('click', () => {
  document.getElementById('add-staff-form').style.display = 'block';
});

document.getElementById('btn-save-staff').addEventListener('click', async () => {
  const name = document.getElementById('new-staff-name').value.trim();
  const errEl = document.getElementById('staff-error');
  if (!name) return;
  try {
    await apiCall('addStaff', { name });
    document.getElementById('new-staff-name').value = '';
    document.getElementById('add-staff-form').style.display = 'none';
    await goToStaffSelect();
  } catch (e) {
    errEl.textContent = e.message;
  }
});

document.getElementById('btn-goto-menu').addEventListener('click', () => {
  const select = document.getElementById('staff-select');
  if (!select.value) {
    document.getElementById('staff-error').textContent = 'スタッフを選択してください';
    return;
  }
  state.staffName = select.value;
  setText('menu-store-name', state.store);
  setText('menu-staff-label', '棚卸実施者: ' + state.staffName);
  showScreen('screen-menu');
});

// ---- メニュー ----
// 棚卸に入る前に、今月の状態(未実施/承認待ち/差し戻し/承認済み)を確認する。
// 承認済みならロックして進めない。今月すでに棚卸の記録があれば「追加する/新しく数え直す」を選ばせる。
document.getElementById('btn-nav-stocktake').addEventListener('click', async (ev) => {
  // アプリを閉じた等で送信前に途中経過が残っている場合は、まずそれを復元するか確認する。
  // (通信不要でその場で分かるので、確認中の通信の前に聞く)
  const draft = loadStocktakeDraft();
  const draftItems = draft ? Object.values(draft.tally || {}) : [];
  if (draftItems.length) {
    const draftTotal = draftItems.reduce((sum, item) => sum + item.count, 0);
    const resume = confirm(
      `前回、棚卸し作業の途中(${draftItems.length}品目・${draftTotal}本)で中断されたデータが残っています。\n` +
      '続きから再開しますか?\n(「キャンセル」を選ぶと、この途中データは削除され、最初からになります)'
    );
    if (resume) {
      resumeStocktakeDraft(draft);
      return;
    }
    clearStocktakeDraft();
  }
  // 結果が分かるまで画面を切り替えない(「未実施」で結局そのままスキャン画面へ進む場合に
  // 中間画面が一瞬映って消えるのを避けるため)。ただし何も反応がないように見えないよう、
  // ボタン自体は確認中だと分かる表示にする。
  const lockedEl = document.getElementById('stocktake-mode-locked');
  const choiceEl = document.getElementById('stocktake-mode-choice');
  lockedEl.style.display = 'none';
  choiceEl.style.display = 'none';
  await withButtonBusy(ev.currentTarget, '確認中...', async () => {
  try {
    const status = await apiCall('getStocktakeStatus', { month: currentYearMonth() });
    if (status.status === '承認済み') {
      lockedEl.querySelector('p').textContent =
        '今月は本社承認済みのため、棚卸はできません。修正が必要な場合は本社にご連絡ください(差し戻しを依頼)。';
      lockedEl.style.display = 'block';
      showScreen('screen-stocktake-mode');
    } else if (status.implementedDate) {
      const dateStr = new Date(status.implementedDate).toLocaleDateString('ja-JP');
      document.getElementById('stocktake-mode-info').textContent =
        status.status === '差し戻し'
          ? `本社から差し戻されています(理由: ${status.rejectedReason || '記載なし'})。前回の実施日: ${dateStr}`
          : `今月はすでに ${dateStr} に棚卸を実施しています。`;
      choiceEl.style.display = 'block';
      showScreen('screen-stocktake-mode');
    } else {
      startStocktakeScan('overwrite');
    }
  } catch (e) {
    lockedEl.querySelector('p').textContent = e.message;
    lockedEl.style.display = 'block';
    showScreen('screen-stocktake-mode');
  }
  });
});

document.getElementById('btn-back-stocktake-mode').addEventListener('click', () => showScreen('screen-menu'));
document.getElementById('btn-stocktake-mode-add').addEventListener('click', () => startStocktakeScan('add'));

// 「新しく数え直す」は、スキャンし忘れた商品に前回のカウントが残ってしまわないよう、
// 今月すでに記録されている棚卸データを削除してから数え直す。取り消せない操作のため
// 必ず確認ダイアログを出す。
document.getElementById('btn-stocktake-mode-overwrite').addEventListener('click', async (ev) => {
  const confirmed = confirm(
    '「新しく数え直す」を選ぶと、今月すでに記録されている棚卸データはすべて削除されます。\n' +
    'この操作は取り消せません。よろしいですか?'
  );
  if (!confirmed) return;
  await withButtonBusy(ev.currentTarget, '削除中...', async () => {
    try {
      await apiCall('resetStocktakeThisMonth', {});
      startStocktakeScan('overwrite');
    } catch (e) {
      document.getElementById('stocktake-mode-info').textContent = e.message;
    }
  });
});

function startStocktakeScan(mode) {
  state.tally = {};
  state.stocktakeMode = mode;
  clearStocktakeDraft();
  loadStocktakeProductList().catch((e) => console.error(e));
  renderTally();
  document.getElementById('tally-container').style.display = 'none';
  document.getElementById('stocktake-review').style.display = 'none';
  document.getElementById('stocktake-live-status').style.display = 'none';
  showScreen('screen-stocktake');
  rearmGate(stocktakeGate);
  startScanner('reader', onStocktakeScan, stocktakeGate).catch((e) => console.error(e));
}

/** アプリを閉じる等で送信前に失われかけた棚卸データを、保存されていた分から再開する。 */
function resumeStocktakeDraft(draft) {
  state.tally = draft.tally || {};
  state.stocktakeMode = draft.mode || 'overwrite';
  loadStocktakeProductList().catch((e) => console.error(e));
  document.getElementById('tally-container').style.display = 'none';
  document.getElementById('stocktake-review').style.display = 'none';
  document.getElementById('stocktake-live-status').style.display = 'none';
  showScreen('screen-stocktake');
  rearmGate(stocktakeGate);
  startScanner('reader', onStocktakeScan, stocktakeGate).catch((e) => console.error(e));
  renderTally();
}

document.getElementById('btn-switch-camera-stocktake').addEventListener('click', async (ev) => {
  await withButtonBusy(ev.currentTarget, '切り替え中...', async () => {
    await switchCamera('reader', onStocktakeScan, stocktakeGate, 'stocktake-status');
    rearmGate(stocktakeGate);
  });
});

// ---- 棚卸: 未スキャン商品一覧 ----
// 「内容を確認する」を押した時点で送信はせず、まず商品マスタの全件と今スキャン済みの
// 一覧を突き合わせて未スキャン商品(欠品・スキャン漏れの可能性がある)を確認してもらい、
// 「本社へ送信する」を押した時点で初めてsubmitStocktakeを呼ぶ2段階にしている。
let stocktakeAllProducts = [];
let stocktakeProductsLoaded = false;

async function loadStocktakeProductList() {
  stocktakeAllProducts = [];
  stocktakeProductsLoaded = false;
  document.getElementById('unscanned-container').innerHTML = '';
  document.getElementById('stocktake-review').style.display = 'none';
  const data = await apiCall('listProducts', {});
  stocktakeAllProducts = data.products;
  stocktakeProductsLoaded = true;
}

function renderUnscanned() {
  const container = document.getElementById('unscanned-container');
  container.innerHTML = '';

  const unscanned = stocktakeAllProducts.filter((p) => !state.tally[p.code]);
  if (!unscanned.length) {
    container.textContent = '全商品をスキャン済みです';
    return;
  }
  const byBrand = {};
  unscanned.forEach((p) => {
    const brand = p.brand || '(ブランド未設定)';
    if (!byBrand[brand]) byBrand[brand] = [];
    byBrand[brand].push(p);
  });
  Object.keys(byBrand).sort().forEach((brand) => {
    const group = document.createElement('div');
    group.className = 'tally-group';
    const h3 = document.createElement('h3');
    h3.textContent = brand;
    group.appendChild(h3);
    byBrand[brand]
      .sort((a, b) => String(a.name).localeCompare(String(b.name), 'ja'))
      .forEach((p) => {
        const row = document.createElement('div');
        row.className = 'tally-row';
        row.innerHTML = `<span class="tally-name">${p.name}${p.colorNo ? ' (' + p.colorNo + ')' : ''}</span>`;
        group.appendChild(row);
      });
    container.appendChild(group);
  });
}

document.getElementById('btn-nav-inventory').addEventListener('click', () => {
  document.getElementById('inventory-search').value = '';
  document.getElementById('inventory-brand-filter').innerHTML = '<option value="">全ブランド</option>';
  document.getElementById('inventory-result').style.display = 'none';
  const monthInput = document.getElementById('inventory-month');
  if (!monthInput.value) monthInput.value = currentYearMonth();
  showScreen('screen-inventory');
});

function currentYearMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

document.getElementById('btn-confirm-inventory-month').addEventListener('click', async () => {
  const month = document.getElementById('inventory-month').value;
  if (!month) return;
  document.getElementById('inventory-result').style.display = 'block';
  await loadInventory(month);
});

// ---- 在庫検索(スキャンで自店の最新在庫を確認) ----
const stockLookupGate = createGate({ rearmOnMiss: true, requiredMisses: 8 });
let scannerStockLookup = null;

document.getElementById('btn-nav-stock-lookup').addEventListener('click', () => {
  document.getElementById('stock-lookup-result').style.display = 'none';
  document.getElementById('stock-lookup-status').textContent = '';
  document.getElementById('btn-rescan-stock-lookup').style.display = 'none';
  showScreen('screen-stock-lookup');
  rearmGate(stockLookupGate);
  startScanner('reader-4', onStockLookupScan, stockLookupGate).catch((e) => console.error(e));
});

document.getElementById('btn-rescan-stock-lookup').addEventListener('click', () => {
  document.getElementById('stock-lookup-result').style.display = 'none';
  document.getElementById('stock-lookup-status').textContent = '';
  document.getElementById('btn-rescan-stock-lookup').style.display = 'none';
  rearmGate(stockLookupGate);
});

document.getElementById('btn-back-4').addEventListener('click', async () => {
  await stopScanner(scannerStockLookup);
  showScreen('screen-menu');
});

async function onStockLookupScan(code) {
  try {
    const result = await apiCall('lookupCurrentStock', { code });
    if (!result.found) {
      document.getElementById('stock-lookup-result').style.display = 'none';
      document.getElementById('stock-lookup-status').textContent = '未登録のコードです: ' + code;
      document.getElementById('btn-rescan-stock-lookup').style.display = 'block';
      return;
    }
    document.getElementById('stock-lookup-status').textContent = '';
    document.getElementById('stock-lookup-result').style.display = 'block';
    document.getElementById('stock-lookup-name').textContent = result.name;
    document.getElementById('stock-lookup-brand').textContent = result.brand || '';
    document.getElementById('stock-lookup-colorno').textContent = result.colorNo || '';
    document.getElementById('stock-lookup-count').textContent =
      result.currentStock + '本' + (result.outOfStock ? '(欠品)' : '');
    document.getElementById('btn-rescan-stock-lookup').style.display = 'block';
  } catch (e) {
    document.getElementById('stock-lookup-status').textContent = e.message;
  }
}

document.getElementById('btn-manual-add-3').addEventListener('click', () => {
  const input = document.getElementById('manual-code-3');
  const code = input.value.trim();
  if (code) onStockLookupScan(code);
});

document.getElementById('btn-back-1').addEventListener('click', async () => {
  await stopScanner(scannerStocktake);
  showScreen('screen-menu');
});
document.getElementById('btn-back-2').addEventListener('click', () => showScreen('screen-menu'));

// ---- バーコード/QRスキャナ共通処理 ----
//
// 同じ商品を1回スキャンしただけなのに何度もカウントされてしまう問題への対処として、
// 「スキャン成功後は、そのコードが画面から一旦外れて再度検出されなくなるまで次の
// スキャンを無視する」というゲート(gate)を導入している。html5-qrcodeは検出できな
// かったフレームでも毎回コールバックを呼ぶので、それを利用して「連続で一定回数
// 検出できなかった=現物が画面から外れた」とみなし、そこで初めて次のスキャンを
// 受け付ける。棚卸(rearmOnMiss:true)は現物を離せば自動で次を受け付け、
// 入荷登録・破棄登録(rearmOnMiss:false)は登録操作が終わるまで自動では再開しない。
//
// 入荷登録・破棄登録で登録操作が完了した直後にrearmGate()で即座に再開してしまうと、
// 登録した現物がまだカメラに映ったままの場合、間を置かずに同じバーコードを再検出して
// しまい、次の商品にスキャンを移せなくなる(固まったように見える)。そのため、登録が
// 完了した直後だけは「現物が画面から一旦外れるまで待ってから次を受け付ける」
// rearmGateAfterMiss()を使う。手動の「読み直す」ボタンや通信エラー時のリトライは、
// 今映っているものをすぐ読み直したいので rearmGate() のまま即座に再開する。
// valueAware: true のゲート(棚卸・在庫検索など、確定操作を挟まず連続でスキャンし続ける
// 画面用)は「画面から消えたはずの時間」という推測ではなく、実際にデコードされた値そのもの
// で判断する。同じ値が続く間は無視し、違う値が読めた瞬間に即カウントするので、次の商品に
// すぐ移れる(=入荷登録が安定して感じられるのと同じ理由)一方、同じ商品を持ち続けても
// 手ブレやピントの一瞬のズレで誤って再カウントされることがない。
function createGate(options) {
  return {
    armed: true,
    lastValue: null,
    missCount: 0,
    rearmOnMiss: !!options.rearmOnMiss,
    valueAware: !!options.valueAware,
    pendingRearm: false,
    requiredMisses: options.requiredMisses || 8
  };
}

function rearmGate(gate) {
  gate.armed = true;
  gate.lastValue = null;
  gate.missCount = 0;
  gate.pendingRearm = false;
}

function rearmGateAfterMiss(gate) {
  gate.pendingRearm = true;
  gate.missCount = 0;
}

// 自動再開(タイムアウトや値の比較)は2回試して2回とも誤カウントが再発したため、
// 入荷登録と同じ「明示的に次へ進むまで完全に止める」方式に統一した。
const stocktakeGate = createGate({ rearmOnMiss: false });
const hqIncomingGate = createGate({ rearmOnMiss: false });
const hqDisposalGate = createGate({ rearmOnMiss: false });

// 一部の機種(iPhone13など)は「背面カメラ」指定だとiOSがピントの合いにくいレンズを
// 自動選択してしまうことがある。カメラの一覧を取得しておき、ユーザーが手動で別の
// 物理レンズに切り替えられるようにする(切り替えた選択はこのページを開いている間は
// 覚えておき、同じ画面に入り直しても引き継ぐ)。
let cachedCameraList = null;
const cameraIndexByElement = {};

async function getCameraList() {
  if (!cachedCameraList) {
    try {
      cachedCameraList = await Html5Qrcode.getCameras();
    } catch (e) {
      cachedCameraList = [];
    }
  }
  return cachedCameraList;
}

function getSelectedCameraId(elementId) {
  const idx = cameraIndexByElement[elementId];
  if (idx === undefined || !cachedCameraList || !cachedCameraList[idx]) return undefined;
  return cachedCameraList[idx].id;
}

/** 次の物理カメラに切り替えて、同じ画面のスキャンをそのカメラで再開する。 */
async function switchCamera(elementId, onSuccess, gate, statusElId) {
  const statusEl = statusElId ? document.getElementById(statusElId) : null;
  const cameras = await getCameraList();
  if (cameras.length < 2) {
    if (statusEl) statusEl.textContent = '切り替えられる別のカメラが見つかりませんでした';
    return;
  }
  const nextIndex = ((cameraIndexByElement[elementId] || 0) + 1) % cameras.length;
  cameraIndexByElement[elementId] = nextIndex;
  await stopScanner(scannerForElement_(elementId));
  if (statusEl) statusEl.textContent = `カメラを切り替えました(${nextIndex + 1}/${cameras.length})`;
  await startScanner(elementId, onSuccess, gate);
}

async function startScanner(elementId, onSuccess, gate) {
  const formats = [
    Html5QrcodeSupportedFormats.QR_CODE,
    Html5QrcodeSupportedFormats.EAN_13,
    Html5QrcodeSupportedFormats.EAN_8,
    Html5QrcodeSupportedFormats.UPC_A,
    Html5QrcodeSupportedFormats.UPC_E,
    Html5QrcodeSupportedFormats.CODE_128,
    Html5QrcodeSupportedFormats.CODE_39
  ];
  const scanner = new Html5Qrcode(elementId, { formatsToSupport: formats, verbose: false });
  if (elementId === 'reader') scannerStocktake = scanner;
  if (elementId === 'reader-4') scannerStockLookup = scanner;
  if (elementId === 'reader-hq-incoming') scannerHqIncoming = scanner;
  if (elementId === 'reader-hq-disposal') scannerHqDisposal = scanner;
  if (elementId === 'reader-hq-stock-lookup') scannerHqStockLookup = scanner;

  const selectedCameraId = getSelectedCameraId(elementId);
  // 特定のカメラに切り替え済みならdeviceIdで指定し、そうでなければ従来通り
  // facingMode(背面カメラ)にまかせる。videoConstraintsを指定すると外側の
  // 最初の引数は無視されるため、両方に同じ内容を渡す。
  const cameraTarget = selectedCameraId ? { deviceId: { exact: selectedCameraId } } : { facingMode: 'environment' };
  const videoConstraints = selectedCameraId
    ? { deviceId: { exact: selectedCameraId }, width: { ideal: 1920 }, height: { ideal: 1080 } }
    : { facingMode: { exact: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } };

  await scanner.start(
    cameraTarget,
    {
      fps: 10,
      // 正方形だと横長のバーコードに対して余白が多く読み取りにくいため、横長の矩形にする
      qrbox: (viewfinderWidth, viewfinderHeight) => ({
        width: Math.max(240, Math.floor(viewfinderWidth * 0.85)),
        height: Math.max(120, Math.floor(viewfinderHeight * 0.4))
      }),
      // iOSは既定だと解像度が低く、バーコードの細い線がつぶれやすいため高解像度も要求する。
      videoConstraints: videoConstraints
    },
    (decodedText) => {
      if (gate.valueAware) {
        if (decodedText === gate.lastValue) {
          gate.missCount = 0; // 同じ商品がまだ画面内にある(まだ数えない)
          return;
        }
        gate.lastValue = decodedText; // 違う値が読めた=次の商品に移ったとみなし、即カウントする
        gate.missCount = 0;
        onSuccess(decodedText);
        return;
      }
      if (!gate.armed) {
        gate.missCount = 0; // まだ画面内に写っている(検出できている)ので外れた判定をリセット
        return;
      }
      gate.armed = false;
      gate.missCount = 0;
      onSuccess(decodedText);
    },
    () => {
      if (gate.valueAware) {
        if (gate.lastValue !== null) {
          gate.missCount += 1;
          if (gate.missCount >= gate.requiredMisses) {
            gate.lastValue = null; // 完全に画面から外れたとみなし、同じ値が再度出てきたら数えられるようにする
            gate.missCount = 0;
          }
        }
        return;
      }
      if (!gate.armed && (gate.rearmOnMiss || gate.pendingRearm)) {
        gate.missCount += 1;
        if (gate.missCount >= gate.requiredMisses) {
          gate.armed = true;
          gate.pendingRearm = false;
        }
      }
    }
  );
  activeScannerCtx = { elementId, onSuccess, gate };
}

async function stopScanner(scanner) {
  activeScannerCtx = null;
  if (scanner) {
    try {
      await scanner.stop();
      scanner.clear();
    } catch (e) {
      /* 既に停止している場合は無視 */
    }
  }
}

function scannerForElement_(elementId) {
  if (elementId === 'reader') return scannerStocktake;
  if (elementId === 'reader-4') return scannerStockLookup;
  if (elementId === 'reader-hq-incoming') return scannerHqIncoming;
  if (elementId === 'reader-hq-disposal') return scannerHqDisposal;
  if (elementId === 'reader-hq-stock-lookup') return scannerHqStockLookup;
  return null;
}

/** カメラが固まって見えるとき用に、スキャナーを止めてから起動し直す。連打による多重実行は無視する。 */
async function restartScanner(elementId, onSuccess, gate) {
  if (scannerRestarting) return;
  scannerRestarting = true;
  try {
    await stopScanner(scannerForElement_(elementId));
    rearmGate(gate);
    await startScanner(elementId, onSuccess, gate).catch((e) => console.error(e));
  } finally {
    scannerRestarting = false;
  }
}

// ---- 棚卸 ----
async function onStocktakeScan(code) {
  try {
    // 商品マスタは棚卸画面に入った時点で1回だけ読み込み済み(stocktakeAllProducts)。
    // スキャンのたびに毎回サーバーへ問い合わせると通信待ちの分だけ画面の反応が
    // 遅く感じられ、スキャンする商品が増えるほど積み重なって負担になるため、
    // まずは読み込み済みのデータから探す。読み込みがまだ終わっていない場合
    // (画面を開いた直後など)だけ、念のためサーバーに問い合わせる。
    let product = stocktakeAllProducts.find((p) => String(p.code) === String(code)) || null;
    if (!product && !stocktakeProductsLoaded) {
      product = await apiCall('lookupProduct', { code });
    }
    if (!product) {
      document.getElementById('stocktake-status').textContent = '未登録のコードです: ' + code;
      rearmGate(stocktakeGate); // 登録できなかった場合はすぐ次を試せるようにする
      return;
    }
    if (!state.tally[code]) {
      state.tally[code] = { code, name: product.name, brand: product.brand || '(ブランド未設定)', count: 0 };
    }
    state.tally[code].count += 1;
    document.getElementById('stocktake-status').textContent = '';
    document.getElementById('stocktake-live-status').style.display = 'block';
    document.getElementById('stocktake-last-name').textContent = product.name;
    document.getElementById('stocktake-last-count').textContent = state.tally[code].count;
    renderTally();
    // 「次の商品をスキャンする」を押すまでカメラは反応しない(入荷登録と同じ、1件ずつ
    // 明示的に確定する方式)。ここでは意図的に再開しない。
  } catch (e) {
    document.getElementById('stocktake-status').textContent = e.message;
    rearmGate(stocktakeGate); // 通信エラー時はすぐ再試行できるようにする
  }
}

document.getElementById('btn-stocktake-next').addEventListener('click', () => {
  // 今映っている現物からカメラが外れるまで待って再開する(入荷登録の「登録した直後」と同じ
  // 考え方。押した瞬間まだ同じ商品が映っていることが多いため、即再開だと再カウントしてしまう)
  rearmGateAfterMiss(stocktakeGate);
  document.getElementById('stocktake-status').textContent = '次の商品をスキャンしてください';
  // 前の商品名を表示したままだと押した反応が見えず不安になるため、次にスキャンできる
  // ようになるまで「直前のスキャン」カードを一旦隠す
  document.getElementById('stocktake-live-status').style.display = 'none';
});

document.getElementById('btn-manual-add-1').addEventListener('click', () => {
  const input = document.getElementById('manual-code-1');
  const code = input.value.trim();
  if (code) onStocktakeScan(code);
  input.value = '';
});

function renderTally() {
  const container = document.getElementById('tally-container');
  container.innerHTML = '';
  const byBrand = {};
  Object.values(state.tally).forEach((item) => {
    if (!byBrand[item.brand]) byBrand[item.brand] = [];
    byBrand[item.brand].push(item);
  });
  Object.keys(byBrand).sort().forEach((brand) => {
    const group = document.createElement('div');
    group.className = 'tally-group';
    const h3 = document.createElement('h3');
    h3.textContent = brand;
    group.appendChild(h3);
    byBrand[brand]
      .sort((a, b) => String(a.name).localeCompare(String(b.name), 'ja'))
      .forEach((item) => {
        const row = document.createElement('div');
        row.className = 'tally-row';
        row.innerHTML = `
          <span class="tally-name">${item.name}</span>
          <span class="tally-stepper">
            <button type="button" class="stepper-btn" data-code="${item.code}" data-delta="-1" aria-label="1本減らす">−</button>
            <input type="number" class="tally-count-input" data-code="${item.code}" value="${item.count}" min="0">
            <button type="button" class="stepper-btn" data-code="${item.code}" data-delta="1" aria-label="1本増やす">+</button>
          </span>`;
        group.appendChild(row);
      });
    container.appendChild(group);
  });

  container.querySelectorAll('.stepper-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const code = btn.dataset.code;
      const current = state.tally[code] ? state.tally[code].count : 0;
      setTallyCount(code, current + Number(btn.dataset.delta));
    });
  });
  container.querySelectorAll('.tally-count-input').forEach((input) => {
    input.addEventListener('change', () => {
      setTallyCount(input.dataset.code, Number(input.value));
    });
  });

  updateTallySummary();
}

/** 個数を手動修正する。0にした場合はスキャン一覧から除く。 */
function setTallyCount(code, newCount) {
  if (!state.tally[code]) return;
  const count = Math.max(0, Math.floor(newCount) || 0);
  if (count === 0) {
    delete state.tally[code];
  } else {
    state.tally[code].count = count;
  }
  renderTally();
}

function updateTallySummary() {
  const items = Object.values(state.tally);
  const totalCount = items.reduce((sum, item) => sum + item.count, 0);
  const summaryEl = document.getElementById('tally-summary');
  const confirmBtn = document.getElementById('btn-confirm-stocktake');
  if (!items.length) {
    summaryEl.textContent = 'まだ何もスキャンしていません';
    confirmBtn.disabled = true;
    document.getElementById('stocktake-live-status').style.display = 'none';
    clearStocktakeDraft();
  } else {
    summaryEl.textContent = `${items.length}品目 / 合計${totalCount}本をスキャン済み。よろしければ内容を確認してください`;
    confirmBtn.disabled = false;
    document.getElementById('stocktake-total-items').textContent = items.length;
    document.getElementById('stocktake-total-count').textContent = totalCount;
    saveStocktakeDraft();
  }
}

document.getElementById('btn-confirm-stocktake').addEventListener('click', () => {
  const items = Object.values(state.tally);
  if (!items.length) {
    document.getElementById('stocktake-status').textContent = 'まだ何もスキャンしていません';
    return;
  }
  document.getElementById('stocktake-status').textContent = '';
  renderUnscanned();
  // 未スキャン一覧はボタンを押すまで非表示にする(見たい人だけ確認できればよく、
  // 「本社へ送信する」への到達を毎回さえぎらないようにするため)
  document.getElementById('unscanned-wrap').style.display = 'none';
  document.getElementById('btn-toggle-unscanned').textContent = '未スキャンの商品を確認する';
  // スキャン中は隠していた商品ごとの一覧を、確認のタイミングでだけ表示する
  document.getElementById('tally-container').style.display = 'block';
  // 「直前のスキャン」はもう関係ないので、確認画面に移ったら隠す(表示され続けて
  // 混乱するとの指摘があったため)
  document.getElementById('stocktake-live-status').style.display = 'none';
  const reviewEl = document.getElementById('stocktake-review');
  reviewEl.style.display = 'block';
  // 商品数が多いとタリー一覧が長くなり、表示が変わったこと自体が画面外で見えないことがあるため、
  // 「本社へ送信する」ボタンが確実に見えるようスクロールする
  reviewEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

document.getElementById('btn-toggle-unscanned').addEventListener('click', () => {
  const wrap = document.getElementById('unscanned-wrap');
  const isHidden = wrap.style.display === 'none';
  wrap.style.display = isHidden ? 'block' : 'none';
  document.getElementById('btn-toggle-unscanned').textContent = isHidden ? '未スキャンの商品を隠す' : '未スキャンの商品を確認する';
});

document.getElementById('btn-back-to-scan').addEventListener('click', () => {
  document.getElementById('stocktake-review').style.display = 'none';
  document.getElementById('tally-container').style.display = 'none';
  // カメラは前回スキャンした商品のあとまだ止まったままなので、「次の商品をスキャンする」
  // ボタンに再びアクセスできるよう、直前のスキャン情報を再表示する
  if (Object.keys(state.tally).length) {
    document.getElementById('stocktake-live-status').style.display = 'block';
  }
});

document.getElementById('btn-send-stocktake').addEventListener('click', async () => {
  const items = Object.values(state.tally).map((t) => ({ code: t.code, count: t.count }));
  try {
    const data = await apiCall('submitStocktake', { staffName: state.staffName, items, mode: state.stocktakeMode });
    document.getElementById('stocktake-status').textContent =
      `棚卸を送信しました(${data.recorded.length}品目)。` +
      (data.unknownCodes.length ? ` 未登録コード: ${data.unknownCodes.join(', ')}` : '');
    state.tally = {};
    state.stocktakeMode = 'overwrite';
    renderTally();
    document.getElementById('stocktake-review').style.display = 'none';
    document.getElementById('tally-container').style.display = 'none';
  } catch (e) {
    document.getElementById('stocktake-status').textContent = e.message;
  }
});

// ---- 在庫情報(店舗) ----
let inventoryItems = [];
let currentInventoryMonth = '';

async function loadInventory(month) {
  const container = document.getElementById('inventory-container');
  container.textContent = '読み込み中...';
  try {
    const data = await apiCall('getInventorySummary', { month });
    currentInventoryMonth = month;
    inventoryItems = data.items;
    updateInventoryBrandFilterOptions();
    renderInventoryList(document.getElementById('inventory-search').value);
  } catch (e) {
    container.textContent = e.message;
  }
}

/** ブランドの絞り込みプルダウンを、いま表示している範囲に実在するブランドだけで作り直す。 */
function updateInventoryBrandFilterOptions() {
  const select = document.getElementById('inventory-brand-filter');
  const current = select.value;
  const brands = Array.from(new Set(inventoryItems.map((item) => item.brand).filter(Boolean)))
    .sort((a, b) => String(a).localeCompare(String(b), 'ja'));
  select.innerHTML = '<option value="">全ブランド</option>';
  brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
  select.value = brands.includes(current) ? current : '';
}

document.getElementById('inventory-brand-filter').addEventListener('change', () => {
  renderInventoryList(document.getElementById('inventory-search').value);
});

/**
 * 検索用に文字列を正規化する。
 * - 全角英数記号 → 半角(例: "０６" → "06")
 * - カタカナ → ひらがな(例: "イシコッパ" → "いしこっぱ")
 * - 大文字 → 小文字
 * これにより「06」「０６」、"nfa"「ＮＦＡ」、「イシコッパ」「いしこっぱ」などの
 * 表記ゆれを同一視して検索できる。
 */
function normalizeSearchText(str) {
  if (!str) return '';
  return String(str)
    .replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/　/g, ' ')
    .replace(/[ァ-ヶ]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0x60))
    .toLowerCase();
}

let inventoryFilteredForExport = [];

function renderInventoryList(query) {
  const container = document.getElementById('inventory-container');
  if (!inventoryItems.length) {
    container.textContent = '在庫データがまだありません';
    inventoryFilteredForExport = [];
    return;
  }

  const brandFilter = document.getElementById('inventory-brand-filter').value;
  // スペース区切りの複数キーワードはすべて満たす(AND検索)商品だけに絞り込む
  const keywords = normalizeSearchText(query).split(/\s+/).filter(Boolean);
  const filtered = inventoryItems.filter((item) => {
    if (brandFilter && item.brand !== brandFilter) return false;
    if (!keywords.length) return true;
    const haystack = normalizeSearchText([item.name, item.colorNo].filter(Boolean).join(' '));
    return keywords.every((kw) => haystack.includes(kw));
  });
  inventoryFilteredForExport = filtered;

  if (!filtered.length) {
    container.textContent = '該当する商品が見つかりません';
    return;
  }

  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>ブランド</th><th>品名</th><th>カラーNO</th><th>現在庫</th></tr>';
  filtered.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.outOfStock) tr.className = 'out-of-stock';
    tr.innerHTML = `<td>${item.brand || ''}</td><td>${item.name}</td><td>${item.colorNo || ''}</td><td>${item.currentStock}</td>`;
    table.appendChild(tr);
  });
  container.innerHTML = '';
  container.appendChild(table);
}

document.getElementById('inventory-search').addEventListener('input', () => {
  renderInventoryList(document.getElementById('inventory-search').value);
});

document.getElementById('btn-export-inventory').addEventListener('click', () => {
  if (!inventoryFilteredForExport.length) return;
  const rows = [['ブランド', '品名', 'カラーNO', '現在庫']];
  inventoryFilteredForExport.forEach((item) => {
    rows.push([item.brand || '', item.name, item.colorNo || '', item.currentStock]);
  });
  downloadCsv(`在庫情報_${state.store}_${currentInventoryMonth}.csv`, rows);
});

// ==================== 本社用 ====================

async function enterDashboard() {
  const stores = await apiCall('listStores', {});
  state.stores = stores.stores;
  fillStoreSelect('store-filter', true);
  fillStoreSelect('staff-store-select', false);
  fillStoreSelect('log-store-filter', true);
  fillStoreSelect('product-store-select', false);
  fillStoreSelect('hq-incoming-store-select', false);
  fillStoreSelect('hq-disposal-store-select', false);
  fillStoreSelect('hq-review-store-select', false);
  fillStoreSelect('dedup-store-select', false);
  fillStoreSelect('hq-stock-lookup-store-select', false);
  showScreen('screen-dashboard');
  await loadDashboard();
}

function fillStoreSelect(elementId, includeAll) {
  const select = document.getElementById(elementId);
  select.innerHTML = includeAll ? '<option value="">全店舗</option>' : '';
  state.stores.forEach((s) => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    select.appendChild(opt);
  });
}

// ---- ナビゲーション ----
document.getElementById('nav-total-inventory').addEventListener('click', async () => {
  showScreen('screen-total-inventory');
  await loadTotalInventoryScreen();
});
// この3つは互いに依存しない別々のAPI呼び出しなので、順番にawaitすると遅い方の
// 待ち時間が積み重なってしまう(「ブランドで絞り込み」が出てくるまで妙に時間が
// かかる、という指摘の原因)。Promise.allでまとめて並行実行する。
document.getElementById('nav-products').addEventListener('click', async () => {
  showScreen('screen-products');
  cancelEditProduct();
  document.getElementById('product-brand-filter').value = '';
  await Promise.all([loadHqBrandOptions(), loadHqCategoryOptions(), loadProducts()]);
});
document.getElementById('product-store-select').addEventListener('change', async () => {
  cancelEditProduct();
  document.getElementById('product-brand-filter').value = '';
  await Promise.all([loadHqBrandOptions(), loadHqCategoryOptions(), loadProducts()]);
});

document.getElementById('nav-staff').addEventListener('click', async () => { showScreen('hq-screen-staff'); await loadStaff(); });
document.getElementById('nav-accounts').addEventListener('click', () => showScreen('screen-accounts'));
document.getElementById('nav-logs').addEventListener('click', async () => { showScreen('screen-logs'); await loadLogs(); });
document.getElementById('nav-hq-review').addEventListener('click', () => {
  showScreen('screen-hq-review');
  document.getElementById('hq-review-month').value = currentYearMonth();
  document.getElementById('hq-review-summary').innerHTML = '';
  document.getElementById('hq-review-table').innerHTML = '';
  document.getElementById('btn-approve-review').style.display = 'none';
  document.getElementById('btn-reject-review').style.display = 'none';
  document.getElementById('btn-export-hq-review').style.display = 'none';
});

document.getElementById('nav-dedup').addEventListener('click', () => {
  showScreen('screen-dedup');
  document.getElementById('dedup-status').textContent = '';
  document.getElementById('dedup-brand-results').innerHTML = '';
  document.getElementById('dedup-name-results').innerHTML = '';
});
document.getElementById('btn-back-dedup').addEventListener('click', () => showScreen('screen-dashboard'));

document.getElementById('nav-hq-incoming').addEventListener('click', () => {
  resetHqIncomingScreen();
  showScreen('screen-hq-incoming');
  rearmGate(hqIncomingGate);
  loadHqIncomingBrandOptions().catch((e) => console.error(e));
  loadHqIncomingCategoryOptions().catch((e) => console.error(e));
  loadHqIncomingProductList().catch((e) => console.error(e));
  startScanner('reader-hq-incoming', onHqIncomingScan, hqIncomingGate).catch((e) => console.error(e));
});
document.getElementById('nav-hq-disposal').addEventListener('click', () => {
  resetHqDisposalScreen();
  showScreen('screen-hq-disposal');
  rearmGate(hqDisposalGate);
  startScanner('reader-hq-disposal', onHqDisposalScan, hqDisposalGate).catch((e) => console.error(e));
});
// スマホでカメラが完全に固まった場合、ページ内でのスキャナー再起動(stop→start)では
// 直らないことがあり、実際にはタブを閉じて開き直す(=ブラウザによる完全な再取得)しか
// 効かないケースが確認された。ページの再読み込みも同じくブラウザにカメラを完全に
// 手放させる操作なので、こちらの方が確実。
document.getElementById('btn-restart-camera-hq-incoming').addEventListener('click', () => {
  location.reload();
});
document.getElementById('btn-restart-camera-hq-disposal').addEventListener('click', () => {
  location.reload();
});
document.getElementById('hq-incoming-store-select').addEventListener('change', () => {
  resetHqIncomingScreen();
  loadHqIncomingBrandOptions().catch((e) => console.error(e));
  loadHqIncomingProductList().catch((e) => console.error(e));
});
document.getElementById('hq-disposal-store-select').addEventListener('change', resetHqDisposalScreen);

async function loadHqIncomingBrandOptions() {
  const store = document.getElementById('hq-incoming-store-select').value;
  if (!store) return;
  const data = await apiCall('getBrandList', { store });
  const select = document.getElementById('hq-incoming-new-brand-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-brand-toggle-incoming').addEventListener('click', () => {
  document.getElementById('hq-incoming-new-brand-new-form').style.display = 'block';
});

function currentHqIncomingBrandValue() {
  const newForm = document.getElementById('hq-incoming-new-brand-new-form');
  const newValue = document.getElementById('hq-incoming-new-brand-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('hq-incoming-new-brand-select').value;
}

function resetHqIncomingBrandNewForm() {
  document.getElementById('hq-incoming-new-brand-new-form').style.display = 'none';
  document.getElementById('hq-incoming-new-brand-new').value = '';
}

/** OCR自動入力などで、プルダウンに無いブランド名が来た場合は新規入力欄に入れる。 */
function setHqIncomingBrandValue(brand) {
  resetHqIncomingBrandNewForm();
  const select = document.getElementById('hq-incoming-new-brand-select');
  if (!brand) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === brand);
  if (hasOption) {
    select.value = brand;
  } else {
    document.getElementById('hq-incoming-new-brand-new-form').style.display = 'block';
    document.getElementById('hq-incoming-new-brand-new').value = brand;
  }
}

async function loadHqIncomingCategoryOptions() {
  const data = await apiCall('getCategoryList', {});
  const select = document.getElementById('hq-incoming-new-category-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.categories.forEach((category) => {
    const opt = document.createElement('option');
    opt.value = category;
    opt.textContent = category;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-category-toggle-incoming').addEventListener('click', () => {
  document.getElementById('hq-incoming-new-category-new-form').style.display = 'block';
});

function currentHqIncomingCategoryValue() {
  const newForm = document.getElementById('hq-incoming-new-category-new-form');
  const newValue = document.getElementById('hq-incoming-new-category-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('hq-incoming-new-category-select').value;
}

function resetHqIncomingCategoryNewForm() {
  document.getElementById('hq-incoming-new-category-new-form').style.display = 'none';
  document.getElementById('hq-incoming-new-category-new').value = '';
}

function setHqIncomingCategoryValue(category) {
  resetHqIncomingCategoryNewForm();
  const select = document.getElementById('hq-incoming-new-category-select');
  if (!category) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === category);
  if (hasOption) {
    select.value = category;
  } else {
    document.getElementById('hq-incoming-new-category-new-form').style.display = 'block';
    document.getElementById('hq-incoming-new-category-new').value = category;
  }
}

// ---- 入荷登録(未登録バーコード): ブランドを選ぶと、事前登録済みのその商品名一覧を
// プルダウンで出す。既存商品を選べば入力し直さずバーコードだけ紐づけられる。 ----
let hqIncomingProductList = [];

async function loadHqIncomingProductList() {
  const store = document.getElementById('hq-incoming-store-select').value;
  if (!store) return;
  const data = await apiCall('listProducts', { store });
  hqIncomingProductList = data.products;
}

function updateHqIncomingExistingNameOptions() {
  const brand = currentHqIncomingBrandValue();
  const select = document.getElementById('hq-incoming-existing-name-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  if (!brand) return;
  hqIncomingProductList
    .filter((p) => (p.brand || '') === brand)
    .forEach((p) => {
      const opt = document.createElement('option');
      opt.value = p.code;
      opt.textContent = p.name + (p.colorNo ? ` (${p.colorNo})` : '');
      select.appendChild(opt);
    });
}

document.getElementById('hq-incoming-new-brand-select').addEventListener('change', updateHqIncomingExistingNameOptions);
document.getElementById('hq-incoming-new-brand-new').addEventListener('input', updateHqIncomingExistingNameOptions);

document.getElementById('hq-incoming-existing-name-select').addEventListener('change', () => {
  const code = document.getElementById('hq-incoming-existing-name-select').value;
  const product = hqIncomingProductList.find((p) => String(p.code) === String(code));
  if (!product) return;
  // プルダウンで既存商品を選んだので、新規入力欄に残っていた文字は消して優先順位を明確にする
  document.getElementById('hq-incoming-new-name').value = '';
  document.getElementById('hq-incoming-new-color-no').value = product.colorNo || '';
  document.getElementById('hq-incoming-new-item-number').value = product.itemNumber || '';
  setHqIncomingCategoryValue(product.category || '');
});

document.getElementById('hq-incoming-new-name').addEventListener('input', () => {
  // 新規品名を入力し始めたら、プルダウンの選択は解除する(どちらが優先か紛らわしくならないように)
  document.getElementById('hq-incoming-existing-name-select').value = '';
});

function resetHqIncomingNameNewForm() {
  document.getElementById('hq-incoming-new-name').value = '';
}

/** 選択中の品名が、既存商品(バーコードを紐づけるだけ)か新規入力かを返す。 */
function currentHqIncomingNameSelection() {
  const newValue = document.getElementById('hq-incoming-new-name').value.trim();
  if (newValue) {
    return { isNew: true, name: newValue };
  }
  const code = document.getElementById('hq-incoming-existing-name-select').value;
  if (!code) return { isNew: true, name: '' };
  const product = hqIncomingProductList.find((p) => String(p.code) === String(code));
  return { isNew: false, code, product };
}

document.getElementById('btn-back-total-inventory').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-products').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-staff').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-accounts').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-logs').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-hq-review').addEventListener('click', () => showScreen('screen-dashboard'));
document.getElementById('btn-back-hq-incoming').addEventListener('click', async () => {
  await stopScanner(scannerHqIncoming);
  showScreen('screen-dashboard');
});
document.getElementById('btn-back-hq-disposal').addEventListener('click', async () => {
  await stopScanner(scannerHqDisposal);
  showScreen('screen-dashboard');
});

// ---- 入荷登録(本社が店舗を選び、その店舗への納品を登録する) ----
/** 入荷登録・破棄登録で店舗を選び間違えないよう、選択中の店舗名を大きく表示し続ける。 */
function updateStoreBanner(selectId, bannerId) {
  const store = document.getElementById(selectId).value;
  const banner = document.getElementById(bannerId);
  if (store) {
    banner.textContent = store;
    banner.classList.remove('store-banner-empty');
  } else {
    banner.textContent = '店舗を選択してください';
    banner.classList.add('store-banner-empty');
  }
}

function resetHqIncomingScreen() {
  updateStoreBanner('hq-incoming-store-select', 'hq-incoming-store-banner');
  document.getElementById('hq-incoming-known').style.display = 'none';
  document.getElementById('hq-incoming-qr-result').style.display = 'none';
  document.getElementById('hq-incoming-unknown').style.display = 'none';
  document.getElementById('hq-incoming-unknown-hint').textContent =
    'このコードは選択した店舗の商品マスタに未登録です。事前登録済みの商品ならブランド→品名を選ぶだけでバーコードを紐づけられます。無ければ「リストにない商品名はこちら」から新規登録してください。';
  document.getElementById('btn-rescan-hq-incoming').style.display = 'none';
  document.getElementById('hq-incoming-status').textContent = '';
  resetHqIncomingBrandNewForm();
  document.getElementById('hq-incoming-new-brand-select').selectedIndex = 0;
  resetHqIncomingNameNewForm();
  document.getElementById('hq-incoming-existing-name-select').innerHTML = '';
  document.getElementById('hq-incoming-new-color-no').value = '';
  document.getElementById('hq-incoming-new-item-number').value = '';
  resetHqIncomingCategoryNewForm();
  document.getElementById('hq-incoming-new-category-select').selectedIndex = 0;
}

let hqIncomingScannedCode = null;

async function onHqIncomingScan(code) {
  hqIncomingScannedCode = code;
  const store = document.getElementById('hq-incoming-store-select').value;
  try {
    const product = await apiCall('lookupProduct', { store, code });
    resetHqIncomingScreen();
    if (product) {
      document.getElementById('hq-incoming-known').style.display = 'block';
      document.getElementById('hq-incoming-product-name').textContent = product.name;
      document.getElementById('hq-incoming-product-brand').textContent = product.brand || '';
      document.getElementById('hq-incoming-quantity').value = 1;
      document.getElementById('hq-incoming-confirm-store').textContent = '店舗: ' + store;
    } else {
      document.getElementById('hq-incoming-unknown').style.display = 'block';
    }
    document.getElementById('btn-rescan-hq-incoming').style.display = 'block';
  } catch (e) {
    document.getElementById('hq-incoming-status').textContent = e.message;
    rearmGate(hqIncomingGate); // 通信エラー時はすぐ再スキャンできるようにする
  }
}

document.getElementById('btn-manual-add-hq-incoming').addEventListener('click', () => {
  const input = document.getElementById('manual-code-hq-incoming');
  const code = input.value.trim();
  if (code) onHqIncomingScan(code);
});

document.getElementById('btn-rescan-hq-incoming').addEventListener('click', () => {
  resetHqIncomingScreen();
  rearmGate(hqIncomingGate);
});

/** バーコードが無い商品用: スキャン・手入力を介さず、いきなり新規登録フォームを開く。 */
document.getElementById('btn-new-product-no-barcode').addEventListener('click', () => {
  const store = document.getElementById('hq-incoming-store-select').value;
  if (!store) {
    document.getElementById('hq-incoming-status').textContent = '店舗を選択してください';
    return;
  }
  resetHqIncomingScreen();
  hqIncomingScannedCode = null;
  document.getElementById('hq-incoming-unknown-hint').textContent =
    'バーコード無しで新規登録します。コード欄は空欄のまま登録され、自動でQRコードが発行されます。';
  document.getElementById('hq-incoming-unknown').style.display = 'block';
  document.getElementById('btn-rescan-hq-incoming').style.display = 'block';
});

document.getElementById('btn-register-new-from-incoming').addEventListener('click', async (e) => {
  const store = document.getElementById('hq-incoming-store-select').value;
  const brand = currentHqIncomingBrandValue();
  const selection = currentHqIncomingNameSelection();
  const colorNo = document.getElementById('hq-incoming-new-color-no').value.trim();
  const itemNumber = document.getElementById('hq-incoming-new-item-number').value.trim();
  const category = currentHqIncomingCategoryValue();

  if (selection.isNew && !selection.name) {
    document.getElementById('hq-incoming-status').textContent = '品名を選択するか、新しい品名を入力してください';
    return;
  }

  await withButtonBusy(e.currentTarget, '処理中...', async () => {
  try {
    let name;
    let generatedCode = null;
    if (selection.isNew) {
      // 新規商品として登録。コードが空欄(バーコード無しの新規登録)だった場合は
      // 自動発行されたQRコードが返ってくるので、以降の入荷登録にはそちらを使う
      name = selection.name;
      const wasBlank = !hqIncomingScannedCode;
      const result = await apiCall('registerProduct', {
        store, code: hqIncomingScannedCode, brand, name, colorNo, itemNumber, category
      });
      hqIncomingScannedCode = result.code;
      if (wasBlank) generatedCode = result.code;
    } else {
      // 事前登録済みの商品に、いまスキャンしたバーコードを紐づける(新規登録はしない)
      name = selection.product.name;
      await apiCall('updateProduct', {
        store, code: selection.code, newCode: hqIncomingScannedCode,
        brand, name, colorNo, itemNumber, category, memo: selection.product.memo || ''
      });
    }
    document.getElementById('hq-incoming-status').textContent =
      (selection.isNew ? '商品を登録しました。' : 'バーコードを紐づけました。') + '続けて入荷本数を入力してください';
    document.getElementById('hq-incoming-unknown').style.display = 'none';
    document.getElementById('hq-incoming-known').style.display = 'block';
    document.getElementById('hq-incoming-product-name').textContent = name;
    document.getElementById('hq-incoming-product-brand').textContent = brand;
    document.getElementById('hq-incoming-quantity').value = 1;
    document.getElementById('hq-incoming-confirm-store').textContent = '店舗: ' + store;
    const qrResult = document.getElementById('hq-incoming-qr-result');
    if (generatedCode) {
      const holder = document.getElementById('hq-incoming-qr-canvas-holder');
      holder.innerHTML = '';
      new QRCode(holder, { text: String(generatedCode), width: 120, height: 120, correctLevel: QRCode.CorrectLevel.H });
      qrResult.style.display = 'block';
    } else {
      qrResult.style.display = 'none';
    }
    await loadHqIncomingBrandOptions();
    await loadHqIncomingProductList();
  } catch (e) {
    document.getElementById('hq-incoming-status').textContent = e.message;
  }
  });
});

document.getElementById('btn-print-hq-incoming-qr').addEventListener('click', printCurrentModal);

document.getElementById('btn-submit-hq-incoming').addEventListener('click', async (e) => {
  const store = document.getElementById('hq-incoming-store-select').value;
  const quantity = Number(document.getElementById('hq-incoming-quantity').value) || 1;
  await withButtonBusy(e.currentTarget, '処理中...', async () => {
  try {
    await apiCall('recordIncoming', { store, code: hqIncomingScannedCode, quantity });
    resetHqIncomingScreen();
    document.getElementById('hq-incoming-status').textContent = '入荷を登録しました';
    // すぐに再開すると、登録した現物がまだカメラに映ったままの場合に同じ商品を
    // 再検出してしまうため、一旦画面から外れてから次を受け付けるようにする
    rearmGateAfterMiss(hqIncomingGate);
  } catch (e) {
    document.getElementById('hq-incoming-status').textContent = e.message;
  }
  });
});

// ---- 破棄登録(本社が店舗を選び、劣化・不良などによる在庫の廃棄を登録する) ----
function resetHqDisposalScreen() {
  updateStoreBanner('hq-disposal-store-select', 'hq-disposal-store-banner');
  document.getElementById('hq-disposal-known').style.display = 'none';
  document.getElementById('hq-disposal-unknown').style.display = 'none';
  document.getElementById('btn-rescan-hq-disposal').style.display = 'none';
  document.getElementById('hq-disposal-status').textContent = '';
}

let hqDisposalScannedCode = null;

async function onHqDisposalScan(code) {
  hqDisposalScannedCode = code;
  const store = document.getElementById('hq-disposal-store-select').value;
  try {
    const result = await apiCall('lookupCurrentStock', { store, code });
    resetHqDisposalScreen();
    if (result.found) {
      document.getElementById('hq-disposal-known').style.display = 'block';
      document.getElementById('hq-disposal-product-name').textContent = result.name;
      document.getElementById('hq-disposal-product-brand').textContent = result.brand || '';
      document.getElementById('hq-disposal-product-colorno').textContent = result.colorNo || '';
      document.getElementById('hq-disposal-current-stock').textContent =
        result.currentStock + '本' + (result.outOfStock ? '(欠品)' : '');
      document.getElementById('hq-disposal-quantity').value = 1;
      document.getElementById('hq-disposal-confirm-store').textContent = '店舗: ' + store;
    } else {
      document.getElementById('hq-disposal-unknown').style.display = 'block';
    }
    document.getElementById('btn-rescan-hq-disposal').style.display = 'block';
  } catch (e) {
    document.getElementById('hq-disposal-status').textContent = e.message;
    rearmGate(hqDisposalGate); // 通信エラー時はすぐ再スキャンできるようにする
  }
}

document.getElementById('btn-manual-add-hq-disposal').addEventListener('click', () => {
  const input = document.getElementById('manual-code-hq-disposal');
  const code = input.value.trim();
  if (code) onHqDisposalScan(code);
});

document.getElementById('btn-rescan-hq-disposal').addEventListener('click', () => {
  resetHqDisposalScreen();
  rearmGate(hqDisposalGate);
});

document.getElementById('btn-submit-hq-disposal').addEventListener('click', async (e) => {
  const store = document.getElementById('hq-disposal-store-select').value;
  const quantity = Number(document.getElementById('hq-disposal-quantity').value) || 1;
  await withButtonBusy(e.currentTarget, '処理中...', async () => {
  try {
    await apiCall('recordDisposal', { store, code: hqDisposalScannedCode, quantity });
    resetHqDisposalScreen();
    document.getElementById('hq-disposal-status').textContent = '廃棄を登録しました';
    // 入荷登録と同じ理由で、現物が画面から外れてから次を受け付けるようにする
    rearmGateAfterMiss(hqDisposalGate);
  } catch (e) {
    document.getElementById('hq-disposal-status').textContent = e.message;
  }
  });
});

// ---- 本社ダッシュボード(サマリーのみ) ----
async function loadDashboard() {
  const summaryAll = await apiCall('getInventorySummary', {});
  document.getElementById('grand-total').textContent = '総在庫本数: ' + summaryAll.grandTotal;
  const totalsEl = document.getElementById('totals-by-store');
  totalsEl.innerHTML = '';
  Object.keys(summaryAll.totalsByStore).forEach((store) => {
    const p = document.createElement('p');
    p.textContent = store + ' : ' + summaryAll.totalsByStore[store] + '本';
    totalsEl.appendChild(p);
  });

  const outOfStock = await apiCall('getOutOfStock', {});
  const oosEl = document.getElementById('out-of-stock-container');
  oosEl.innerHTML = '';
  if (!outOfStock.items.length) {
    oosEl.textContent = '欠品はありません';
  } else {
    outOfStock.items.forEach((item) => {
      const p = document.createElement('p');
      p.textContent = `${item.store} / ${item.brand || ''} ${item.name}`;
      oosEl.appendChild(p);
    });
  }
}

// ---- 総在庫(店舗別在庫一覧・検索・エクスポート・リセット) ----
async function loadTotalInventoryScreen() {
  document.getElementById('hq-inventory-search').value = '';
  await renderInventoryTable(document.getElementById('store-filter').value);
}

document.getElementById('store-filter').addEventListener('change', (e) => {
  document.getElementById('hq-inventory-search').value = '';
  renderInventoryTable(e.target.value);
});

let hqInventoryAllItems = [];
let hqInventoryForExport = [];

async function renderInventoryTable(store) {
  const container = document.getElementById('hq-inventory-container');
  let data;
  try {
    data = await apiCall('getInventorySummary', { store: store || undefined });
  } catch (e) {
    container.textContent = e.message;
    return;
  }
  hqInventoryAllItems = data.items;
  updateHqInventoryBrandFilterOptions();
  renderHqInventoryRows(document.getElementById('hq-inventory-search').value);
}

/** ブランドの絞り込みプルダウンを、いま表示している範囲に実在するブランドだけで作り直す。 */
function updateHqInventoryBrandFilterOptions() {
  const select = document.getElementById('hq-inventory-brand-filter');
  const current = select.value;
  const brands = Array.from(new Set(hqInventoryAllItems.map((item) => item.brand).filter(Boolean)))
    .sort((a, b) => String(a).localeCompare(String(b), 'ja'));
  select.innerHTML = '<option value="">全ブランド</option>';
  brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
  select.value = brands.includes(current) ? current : '';
}

document.getElementById('hq-inventory-brand-filter').addEventListener('change', () => {
  renderHqInventoryRows(document.getElementById('hq-inventory-search').value);
});

document.getElementById('hq-inventory-search').addEventListener('input', () => {
  renderHqInventoryRows(document.getElementById('hq-inventory-search').value);
});

function renderHqInventoryRows(query) {
  const container = document.getElementById('hq-inventory-container');
  container.innerHTML = '';
  if (!hqInventoryAllItems.length) {
    hqInventoryForExport = [];
    container.textContent = '在庫データがありません';
    return;
  }

  const brandFilter = document.getElementById('hq-inventory-brand-filter').value;
  // スペース区切りの複数キーワードはすべて満たす(AND検索)商品だけに絞り込む
  const keywords = normalizeSearchText(query).split(/\s+/).filter(Boolean);
  const items = hqInventoryAllItems.filter((item) => {
    if (brandFilter && item.brand !== brandFilter) return false;
    if (!keywords.length) return true;
    const haystack = normalizeSearchText([item.code, item.brand, item.name, item.colorNo].filter(Boolean).join(' '));
    return keywords.every((kw) => haystack.includes(kw));
  });
  hqInventoryForExport = items;

  if (!items.length) {
    container.textContent = '該当する商品が見つかりません';
    return;
  }
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>店舗</th><th>ブランド</th><th>品名</th><th>カラーNO</th><th>現在庫</th><th colspan="2"></th></tr>';
  items.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.outOfStock) tr.className = 'out-of-stock';
    tr.innerHTML = `<td>${item.store}</td><td>${item.brand || ''}</td><td>${item.name}</td><td>${item.colorNo || ''}</td><td>${item.currentStock}</td>`;

    const label = `${item.store}の「${item.brand || ''} ${item.name}」`;

    const adjustTd = document.createElement('td');
    const adjustBtn = document.createElement('button');
    adjustBtn.type = 'button';
    adjustBtn.className = 'link';
    adjustBtn.textContent = '修正';
    adjustBtn.addEventListener('click', async (e) => {
      await withButtonBusy(e.currentTarget, '読み込み中...', async () => {
        let product;
        try {
          const productsData = await apiCall('listProducts', { store: item.store });
          product = productsData.products.find((p) => String(p.code) === String(item.code));
        } catch (err) {
          alert(err.message);
          return;
        }
        await openAdjustModal(
          { ...item, category: product ? product.category : '' },
          async ({ newCode, brand, name, colorNo, category, newStock, memo, password }) => {
            await apiCall('updateProduct', {
              store: item.store, code: item.code, newCode, brand, name, colorNo, category,
              itemNumber: product ? product.itemNumber : '', memo: product ? product.memo : ''
            });
            const finalCode = newCode && newCode !== String(item.code) ? newCode : item.code;
            if (Number(newStock) !== item.currentStock) {
              await apiCall('adjustProductStock', { store: item.store, code: finalCode, newStock, memo, password });
            }
            await loadTotalInventoryScreen();
          }
        );
      });
    });
    adjustTd.appendChild(adjustBtn);
    tr.appendChild(adjustTd);

    const resetTd = document.createElement('td');
    const resetBtn = document.createElement('button');
    resetBtn.type = 'button';
    resetBtn.className = 'link';
    resetBtn.textContent = 'リセット';
    resetBtn.addEventListener('click', () => {
      openDangerModal(
        `${label}の在庫を0にリセットします。カラージェルの劣化などで在庫から外す場合に使ってください。この操作は取り消せません。`,
        async (password) => {
          await apiCall('resetProductStock', { store: item.store, code: item.code, password });
          await loadTotalInventoryScreen();
        }
      );
    });
    resetTd.appendChild(resetBtn);
    tr.appendChild(resetTd);

    table.appendChild(tr);
  });
  container.appendChild(table);
}

document.getElementById('btn-export-hq-inventory').addEventListener('click', () => {
  if (!hqInventoryForExport.length) return;
  const rows = [['店舗', 'ブランド', '品名', 'カラーNO', '現在庫']];
  hqInventoryForExport.forEach((item) => {
    rows.push([item.store, item.brand || '', item.name, item.colorNo || '', item.currentStock]);
  });
  downloadCsv('在庫情報_全店舗.csv', rows);
});

document.getElementById('btn-bulk-reset').addEventListener('click', () => {
  const store = document.getElementById('store-filter').value;
  const targetLabel = store || '全店舗';
  const scope = store || 'ALL';
  openDangerModal(
    `${targetLabel}の在庫を、表示されている${hqInventoryForExport.length}品目すべて0にリセットします。この操作は取り消せません。`,
    async (password) => {
      await apiCall('resetStoreInventory', { store: scope, password });
      await loadTotalInventoryScreen();
    }
  );
});

// ---- 在庫検索(本社。店舗を選んでスキャンし、その店舗の最新在庫を確認する) ----
const hqStockLookupGate = createGate({ rearmOnMiss: true, requiredMisses: 8 });

document.getElementById('nav-hq-stock-lookup').addEventListener('click', () => {
  document.getElementById('hq-stock-lookup-result').style.display = 'none';
  document.getElementById('hq-stock-lookup-status').textContent = '';
  document.getElementById('btn-rescan-hq-stock-lookup').style.display = 'none';
  updateStoreBanner('hq-stock-lookup-store-select', 'hq-stock-lookup-store-banner');
  showScreen('screen-hq-stock-lookup');
  rearmGate(hqStockLookupGate);
  startScanner('reader-hq-stock-lookup', onHqStockLookupScan, hqStockLookupGate).catch((e) => console.error(e));
});

document.getElementById('hq-stock-lookup-store-select').addEventListener('change', () => {
  updateStoreBanner('hq-stock-lookup-store-select', 'hq-stock-lookup-store-banner');
});

document.getElementById('btn-back-hq-stock-lookup').addEventListener('click', async () => {
  await stopScanner(scannerHqStockLookup);
  showScreen('screen-dashboard');
});

// カメラが固まったときの対処は他のHQスキャン画面と同じ理由でページ再読み込みにしている
document.getElementById('btn-restart-camera-hq-stock-lookup').addEventListener('click', () => {
  location.reload();
});

async function onHqStockLookupScan(code) {
  const store = document.getElementById('hq-stock-lookup-store-select').value;
  if (!store) {
    document.getElementById('hq-stock-lookup-status').textContent = '店舗を選択してください';
    rearmGate(hqStockLookupGate);
    return;
  }
  try {
    const result = await apiCall('lookupCurrentStock', { store, code });
    if (!result.found) {
      document.getElementById('hq-stock-lookup-result').style.display = 'none';
      document.getElementById('hq-stock-lookup-status').textContent = '未登録のコードです: ' + code;
      document.getElementById('btn-rescan-hq-stock-lookup').style.display = 'block';
      return;
    }
    document.getElementById('hq-stock-lookup-status').textContent = '';
    document.getElementById('hq-stock-lookup-result').style.display = 'block';
    document.getElementById('hq-stock-lookup-name').textContent = result.name;
    document.getElementById('hq-stock-lookup-brand').textContent = result.brand || '';
    document.getElementById('hq-stock-lookup-colorno').textContent = result.colorNo || '';
    document.getElementById('hq-stock-lookup-count').textContent =
      result.currentStock + '本' + (result.outOfStock ? '(欠品)' : '');
    document.getElementById('btn-rescan-hq-stock-lookup').style.display = 'block';
  } catch (e) {
    document.getElementById('hq-stock-lookup-status').textContent = e.message;
    rearmGate(hqStockLookupGate);
  }
}

document.getElementById('btn-manual-add-hq-stock-lookup').addEventListener('click', () => {
  const input = document.getElementById('manual-code-hq-stock-lookup');
  const code = input.value.trim();
  if (code) onHqStockLookupScan(code);
});

document.getElementById('btn-rescan-hq-stock-lookup').addEventListener('click', () => {
  document.getElementById('hq-stock-lookup-result').style.display = 'none';
  document.getElementById('hq-stock-lookup-status').textContent = '';
  document.getElementById('btn-rescan-hq-stock-lookup').style.display = 'none';
  rearmGate(hqStockLookupGate);
});

// ---- 商品マスタ(本社) ----
// 商品マスタの新規登録・修正・削除はすべて本社アカウントのみ(registerProduct/updateProduct/
// deleteProduct はサーバー側でもhq権限を要求している)。店舗アカウントからは登録できない。

async function loadHqBrandOptions() {
  const store = document.getElementById('product-store-select').value;
  if (!store) return;
  const data = await apiCall('getBrandList', { store });
  const select = document.getElementById('p-brand-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-brand-toggle').addEventListener('click', () => {
  document.getElementById('p-brand-new-form').style.display = 'block';
});

/** プルダウンで選んだブランドと、「リストにない」場合の新規入力のどちらかを返す。 */
function currentPBrandValue() {
  const newForm = document.getElementById('p-brand-new-form');
  const newValue = document.getElementById('p-brand-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('p-brand-select').value;
}

function resetPBrandNewForm() {
  document.getElementById('p-brand-new-form').style.display = 'none';
  document.getElementById('p-brand-new').value = '';
}

/** 編集時に既存商品のブランドを反映する。プルダウンに無い名前なら新規入力欄に入れる。 */
function setPBrandValue(brand) {
  resetPBrandNewForm();
  const select = document.getElementById('p-brand-select');
  if (!brand) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === brand);
  if (hasOption) {
    select.value = brand;
  } else {
    document.getElementById('p-brand-new-form').style.display = 'block';
    document.getElementById('p-brand-new').value = brand;
  }
}

// ---- カテゴリー(商品マスタと同じ「プルダウン+リストにない場合の新規入力」パターン) ----
async function loadHqCategoryOptions() {
  const data = await apiCall('getCategoryList', {});
  const select = document.getElementById('p-category-select');
  select.innerHTML = '<option value="">(未選択)</option>';
  data.categories.forEach((category) => {
    const opt = document.createElement('option');
    opt.value = category;
    opt.textContent = category;
    select.appendChild(opt);
  });
}

document.getElementById('btn-new-category-toggle').addEventListener('click', () => {
  document.getElementById('p-category-new-form').style.display = 'block';
});

function currentPCategoryValue() {
  const newForm = document.getElementById('p-category-new-form');
  const newValue = document.getElementById('p-category-new').value.trim();
  if (newForm.style.display !== 'none' && newValue) return newValue;
  return document.getElementById('p-category-select').value;
}

function resetPCategoryNewForm() {
  document.getElementById('p-category-new-form').style.display = 'none';
  document.getElementById('p-category-new').value = '';
}

function setPCategoryValue(category) {
  resetPCategoryNewForm();
  const select = document.getElementById('p-category-select');
  if (!category) {
    select.value = '';
    return;
  }
  const hasOption = Array.from(select.options).some((o) => o.value === category);
  if (hasOption) {
    select.value = category;
  } else {
    document.getElementById('p-category-new-form').style.display = 'block';
    document.getElementById('p-category-new').value = category;
  }
}

let editingProductCode = null;
let currentProductList = [];
let productsFilteredForExport = [];
let selectedQrProducts = new Set();

async function loadProducts() {
  const store = document.getElementById('product-store-select').value;
  const container = document.getElementById('products-table');
  selectedQrProducts.clear(); // 店舗を切り替えたら商品データ自体が別物になるので選択もリセットする
  if (!store) {
    currentProductList = [];
    container.textContent = '店舗を選択してください';
    return;
  }
  const data = await apiCall('listProducts', { store });
  currentProductList = data.products;
  updateProductBrandFilterOptions();
  renderProductsTable();
}

/** ブランドの絞り込みプルダウンを、いま選んでいる店舗に実在するブランドだけで作り直す。 */
function updateProductBrandFilterOptions() {
  const select = document.getElementById('product-brand-filter');
  const current = select.value;
  const brands = Array.from(new Set(currentProductList.map((p) => p.brand).filter(Boolean)))
    .sort((a, b) => String(a).localeCompare(String(b), 'ja'));
  select.innerHTML = '<option value="">全ブランド</option>';
  brands.forEach((brand) => {
    const opt = document.createElement('option');
    opt.value = brand;
    opt.textContent = brand;
    select.appendChild(opt);
  });
  select.value = brands.includes(current) ? current : '';
}

document.getElementById('product-brand-filter').addEventListener('change', () => {
  renderProductsTable();
});

function renderProductsTable() {
  const store = document.getElementById('product-store-select').value;
  const container = document.getElementById('products-table');
  if (!currentProductList.length) {
    container.textContent = '商品がまだ登録されていません';
    productsFilteredForExport = [];
    return;
  }

  const brandFilter = document.getElementById('product-brand-filter').value;
  const filtered = brandFilter ? currentProductList.filter((p) => p.brand === brandFilter) : currentProductList;
  productsFilteredForExport = filtered;

  if (!filtered.length) {
    container.textContent = '該当する商品が見つかりません';
    return;
  }

  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th></th><th>コード</th><th>品名</th><th></th></tr>';
  // data-*属性経由でコードを突き合わせると、数字だけのバーコードがGoogleスプレッドシート
  // 側で数値型として返ってきた場合に文字列(属性値)と型が合わず一致しない不具合が過去に
  // あったため、ボタンにその場でpをクロージャとして持たせて直接参照する。
  filtered.forEach((p) => {
    const tr = document.createElement('tr');

    const checkTd = document.createElement('td');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    // ブランドで絞り込みを変えても、他のブランドで選んだQRコード発行対象は消えないようにする
    checkbox.checked = selectedQrProducts.has(p);
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) selectedQrProducts.add(p);
      else selectedQrProducts.delete(p);
    });
    checkTd.appendChild(checkbox);
    tr.appendChild(checkTd);

    const codeTd = document.createElement('td');
    codeTd.textContent = p.code;
    tr.appendChild(codeTd);

    // 横に何列も並べると狭い画面では横スクロールが必要になるため、品名の下に
    // ブランド・品番・カラーNO・メモをまとめて小さく表示する。
    const nameTd = document.createElement('td');
    const nameLine = document.createElement('div');
    nameLine.textContent = p.name;
    nameTd.appendChild(nameLine);
    const subParts = [];
    if (p.brand) subParts.push(p.brand);
    if (p.itemNumber) subParts.push('品番: ' + p.itemNumber);
    if (p.colorNo) subParts.push('カラーNO: ' + p.colorNo);
    if (p.memo) subParts.push('メモ: ' + p.memo);
    if (subParts.length) {
      const subLine = document.createElement('div');
      subLine.className = 'sub-info';
      subLine.textContent = subParts.join(' / ');
      nameTd.appendChild(subLine);
    }
    tr.appendChild(nameTd);

    const actionsTd = document.createElement('td');

    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'link';
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', () => startEditProduct(p));
    actionsTd.appendChild(editBtn);

    const qrBtn = document.createElement('button');
    qrBtn.type = 'button';
    qrBtn.className = 'link';
    qrBtn.textContent = 'QR表示';
    qrBtn.addEventListener('click', () => showQrViewModal(p.code, p.name, p.brand));
    actionsTd.appendChild(qrBtn);

    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'link';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      // 誤操作を防ぐため、消えるものと消えないものを具体的に書いた警告文にする
      // (「削除しますか?」だけだと軽く見えてうっかり押してしまうとの指摘のため)
      const warning = `「${p.name}」(${p.code})を削除します。\n\n` +
        'これまでの在庫数の記録は消えませんが、商品マスタからは無くなり、' +
        '次にこのバーコードをスキャンすると「未登録」になります。\n\n' +
        '本当に削除しますか?';
      if (!confirm(warning)) return;
      await apiCall('deleteProduct', { store, code: p.code });
      await loadProducts();
    });
    actionsTd.appendChild(delBtn);

    tr.appendChild(actionsTd);
    table.appendChild(tr);
  });
  container.appendChild(table);
}

document.getElementById('btn-export-products').addEventListener('click', () => {
  if (!productsFilteredForExport.length) return;
  const store = document.getElementById('product-store-select').value;
  const rows = [['ブランド', '品名', 'カラーNO', '品番(商品コード)', 'バーコード']];
  productsFilteredForExport.forEach((p) => {
    rows.push([p.brand || '', p.name, p.colorNo || '', p.itemNumber || '', p.code]);
  });
  downloadCsv(`商品マスタ_${store}.csv`, rows);
});

/**
 * QRコードのモーダルを印刷するとき、ダッシュボードの背景まで印刷されてしまわないよう
 * bodyに目印のクラスを付ける(印刷用CSSがこのクラスを見て、モーダルの中身だけを表示する)。
 */
function printCurrentModal() {
  document.body.classList.add('printing-qr');
  window.print();
}
window.addEventListener('afterprint', () => {
  document.body.classList.remove('printing-qr');
});

/** 商品一覧から、既に登録済みの商品のQRコードをいつでも呼び出して印刷できるようにする。 */
function showQrViewModal(code, name, brand) {
  const holder = document.getElementById('qr-view-canvas-holder');
  holder.innerHTML = '';
  try {
    new QRCode(holder, { text: String(code), width: 120, height: 120, correctLevel: QRCode.CorrectLevel.H });
  } catch (e) {
    alert('QRコードの作成に失敗しました: ' + e.message);
    return;
  }
  document.getElementById('qr-view-label').textContent = code + ' / ' + (brand ? brand + ' ' : '') + name;
  document.getElementById('qr-view-modal').style.display = 'flex';
}

document.getElementById('btn-print-qr-view').addEventListener('click', printCurrentModal);
document.getElementById('btn-close-qr-view').addEventListener('click', () => {
  document.getElementById('qr-view-modal').style.display = 'none';
});

// ---- QRコード一括印刷(エーワン 31171: A4・95面・ラベル35×12mm・5列×19段) ----
// 実際の用紙とズレる場合に備えて上・左の余白を調整できるようにし、localStorageに
// 保存して次回以降も使う。
const AONE_COLS = 5;
const AONE_ROWS = 19;
const AONE_PER_PAGE = AONE_COLS * AONE_ROWS;
const AONE_PREVIEW_SCALE = 0.5;
// エーワン公式のテストプリント用紙(フォーマット番号F95A4-1)で実測した値。
// 上16.5mm・左13.5mm、ラベルの間は上下左右とも2mmの隙間(.aone-sheetのgapで指定)。
const AONE_DEFAULT_MARGIN_TOP = 16.5;
const AONE_DEFAULT_MARGIN_LEFT = 13.5;

function getAoneMargins() {
  const top = localStorage.getItem('aoneMarginTop');
  const left = localStorage.getItem('aoneMarginLeft');
  return {
    top: top !== null ? Number(top) : AONE_DEFAULT_MARGIN_TOP,
    left: left !== null ? Number(left) : AONE_DEFAULT_MARGIN_LEFT
  };
}

function renderAoneSheets() {
  const margins = getAoneMargins();
  document.getElementById('aone-margin-top').value = margins.top;
  document.getElementById('aone-margin-left').value = margins.left;

  const products = Array.from(selectedQrProducts);
  const container = document.getElementById('qr-bulk-sheets');
  container.innerHTML = '';

  for (let start = 0; start < products.length; start += AONE_PER_PAGE) {
    const pageItems = products.slice(start, start + AONE_PER_PAGE);

    const wrap = document.createElement('div');
    wrap.className = 'aone-page-wrap';
    wrap.style.width = (210 * AONE_PREVIEW_SCALE) + 'mm';
    wrap.style.height = (297 * AONE_PREVIEW_SCALE) + 'mm';

    const page = document.createElement('div');
    page.className = 'aone-page';
    page.style.transform = `scale(${AONE_PREVIEW_SCALE})`;
    page.style.paddingTop = margins.top + 'mm';
    page.style.paddingLeft = margins.left + 'mm';

    const sheet = document.createElement('div');
    sheet.className = 'aone-sheet';
    pageItems.forEach((p) => {
      const cell = document.createElement('div');
      cell.className = 'aone-cell';
      const qrHolder = document.createElement('div');
      qrHolder.className = 'aone-qr';
      cell.appendChild(qrHolder);
      new QRCode(qrHolder, { text: String(p.code), width: 80, height: 80, correctLevel: QRCode.CorrectLevel.H });
      const text = document.createElement('div');
      text.className = 'aone-text';
      text.textContent = (p.brand ? p.brand + ' ' : '') + p.name;
      cell.appendChild(text);
      sheet.appendChild(cell);
    });

    page.appendChild(sheet);
    wrap.appendChild(page);
    container.appendChild(wrap);
  }
}

/** 商品一覧でチェックした商品のQRコードを、エーワン31171の面付けに合わせて並べる。 */
document.getElementById('btn-print-selected-qr').addEventListener('click', () => {
  if (!selectedQrProducts.size) {
    alert('QRコードを発行したい商品にチェックを入れてください');
    return;
  }
  try {
    renderAoneSheets();
  } catch (e) {
    alert('QRコードの作成に失敗しました: ' + e.message);
    return;
  }
  document.getElementById('qr-bulk-modal').style.display = 'flex';
});

document.getElementById('btn-clear-selected-qr').addEventListener('click', () => {
  if (!selectedQrProducts.size) return;
  selectedQrProducts.clear();
  renderProductsTable();
});

document.getElementById('btn-aone-apply-margin').addEventListener('click', () => {
  const top = Number(document.getElementById('aone-margin-top').value);
  const left = Number(document.getElementById('aone-margin-left').value);
  if (!Number.isFinite(top) || !Number.isFinite(left)) {
    alert('余白は数値で入力してください');
    return;
  }
  localStorage.setItem('aoneMarginTop', top);
  localStorage.setItem('aoneMarginLeft', left);
  renderAoneSheets();
});

document.getElementById('btn-print-qr-bulk').addEventListener('click', printCurrentModal);
document.getElementById('btn-close-qr-bulk').addEventListener('click', () => {
  document.getElementById('qr-bulk-modal').style.display = 'none';
});

function startEditProduct(product) {
  editingProductCode = product.code;
  document.getElementById('p-code').value = product.code;
  setPBrandValue(product.brand || '');
  document.getElementById('p-name').value = product.name || '';
  document.getElementById('p-color-no').value = product.colorNo || '';
  document.getElementById('p-item-number').value = product.itemNumber || '';
  setPCategoryValue(product.category || '');
  document.getElementById('p-memo').value = product.memo || '';
  document.getElementById('product-status').textContent = '';
  document.getElementById('product-edit-form-card').style.display = 'block';
  // 商品一覧は下の方までスクロールされていることが多く、フォームが画面外だと
  // 「編集」を押しても何も起きていないように見えるため、フォームまで自動でスクロールする
  document.getElementById('product-form-title').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function cancelEditProduct() {
  editingProductCode = null;
  ['p-code', 'p-name', 'p-color-no', 'p-item-number', 'p-memo'].forEach((id) => (document.getElementById(id).value = ''));
  resetPBrandNewForm();
  document.getElementById('p-brand-select').selectedIndex = 0;
  resetPCategoryNewForm();
  document.getElementById('p-category-select').selectedIndex = 0;
  document.getElementById('product-edit-form-card').style.display = 'none';
}

document.getElementById('btn-cancel-edit-product').addEventListener('click', cancelEditProduct);

// 新規登録は入荷登録の画面から行うため、ここでの「更新する」は既存商品の編集のみを扱う
document.getElementById('btn-add-product').addEventListener('click', async (e) => {
  const statusEl = document.getElementById('product-status');
  const store = document.getElementById('product-store-select').value;
  if (!store || !editingProductCode) return;
  const codeInput = document.getElementById('p-code').value.trim();
  const payload = {
    store,
    brand: currentPBrandValue(),
    name: document.getElementById('p-name').value.trim(),
    colorNo: document.getElementById('p-color-no').value.trim(),
    itemNumber: document.getElementById('p-item-number').value.trim(),
    category: currentPCategoryValue(),
    memo: document.getElementById('p-memo').value.trim()
  };
  await withButtonBusy(e.currentTarget, '処理中...', async () => {
  try {
    await apiCall('updateProduct', Object.assign({ code: editingProductCode, newCode: codeInput }, payload));
    statusEl.textContent = '更新しました';
    cancelEditProduct();
    await loadProducts();
  } catch (e) {
    statusEl.textContent = e.message;
  }
  });
});

// ---- マスタ整理(ブランド名・品名の表記ゆれチェック) ----
// 大文字/小文字・全角/半角・スペースの有無などの違いだけで、本来同じはずのブランド名/品名が
// 別表記として登録されてしまうケースを見つけるための重複候補検出キー。検索用の
// normalizeSearchText よりさらに厳しく、スペースも取り除いて比較する。
function dedupeKey(str) {
  return normalizeSearchText(str).replace(/\s+/g, '');
}

document.getElementById('btn-scan-dedup').addEventListener('click', async (e) => {
  const store = document.getElementById('dedup-store-select').value;
  const statusEl = document.getElementById('dedup-status');
  if (!store) {
    statusEl.textContent = '店舗を選択してください';
    return;
  }
  await withButtonBusy(e.currentTarget, '検索中...', async () => {
    try {
      await runDedupScan(store);
      document.getElementById('dedup-status').textContent = '';
    } catch (err) {
      statusEl.textContent = err.message;
    }
  });
});

/** ブランド/品名の重複候補を検索して描画する。削除後の再読み込みにも使う。 */
async function runDedupScan(store) {
  const data = await apiCall('listProducts', { store });
  renderBrandDuplicates(store, data.products);
  renderNameDuplicates(store, data.products);
}

function renderBrandDuplicates(store, products) {
  const container = document.getElementById('dedup-brand-results');
  container.innerHTML = '';

  const groups = {};
  products.forEach((p) => {
    if (!p.brand) return;
    const key = dedupeKey(p.brand);
    if (!groups[key]) groups[key] = new Map();
    groups[key].set(p.brand, (groups[key].get(p.brand) || 0) + 1);
  });
  const candidates = Object.values(groups).filter((m) => m.size > 1);

  if (!candidates.length) {
    container.textContent = 'ブランド名の表記ゆれ候補は見つかりませんでした';
    return;
  }

  candidates.forEach((variantMap) => {
    const variants = Array.from(variantMap.entries()); // [[brand, count], ...]
    const card = document.createElement('div');
    card.className = 'card';

    const summary = document.createElement('p');
    summary.textContent = variants.map(([brand, count]) => `${brand}(${count}件)`).join(' / ');
    card.appendChild(summary);

    const label = document.createElement('label');
    label.textContent = '統一後の表記';
    card.appendChild(label);

    const select = document.createElement('select');
    variants.forEach(([brand]) => {
      const opt = document.createElement('option');
      opt.value = brand;
      opt.textContent = brand;
      select.appendChild(opt);
    });
    card.appendChild(select);

    const mergeStatus = document.createElement('p');
    mergeStatus.className = 'status';
    card.appendChild(mergeStatus);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'primary';
    btn.textContent = 'この表記に統一する';
    btn.addEventListener('click', async () => {
      const to = select.value;
      const from = variants.map(([brand]) => brand).filter((brand) => brand !== to);
      const totalCount = variants.reduce((sum, [, count]) => sum + count, 0);
      if (!confirm(`${totalCount}件の商品のブランド名を「${to}」に統一します。よろしいですか?`)) return;
      await withButtonBusy(btn, '処理中...', async () => {
        try {
          const result = await apiCall('mergeBrand', { store, from, to });
          mergeStatus.textContent = `${result.updatedCount}件を「${to}」に統一しました`;
          card.querySelectorAll('select, button').forEach((el) => { el.disabled = true; });
        } catch (err) {
          mergeStatus.textContent = err.message;
        }
      });
    });
    card.appendChild(btn);

    container.appendChild(card);
  });
}

function renderNameDuplicates(store, products) {
  const container = document.getElementById('dedup-name-results');
  container.innerHTML = '';

  const groups = {};
  products.forEach((p) => {
    const key = (p.brand || '') + '||' + dedupeKey(p.name);
    if (!groups[key]) groups[key] = [];
    groups[key].push(p);
  });
  const candidates = Object.values(groups).filter((arr) => new Set(arr.map((p) => p.name)).size > 1);

  if (!candidates.length) {
    container.textContent = '品名の表記ゆれ候補は見つかりませんでした';
    return;
  }

  candidates.forEach((items) => {
    const card = document.createElement('div');
    card.className = 'card';
    const table = document.createElement('table');
    table.className = 'stock-table';
    table.innerHTML = '<tr><th>コード</th><th>ブランド</th><th>品名</th><th>カラーNO</th><th></th></tr>';
    // 商品管理と同様、data-*属性ではなくpをそのままクロージャで持たせてボタンから参照する
    items.forEach((p) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${p.code}</td><td>${p.brand || ''}</td><td>${p.name}</td><td>${p.colorNo || ''}</td>`;

      const actionsTd = document.createElement('td');

      const editBtn = document.createElement('button');
      editBtn.type = 'button';
      editBtn.className = 'link';
      editBtn.textContent = '編集';
      editBtn.addEventListener('click', async () => {
        document.getElementById('product-store-select').value = store;
        document.getElementById('product-brand-filter').value = '';
        showScreen('screen-products');
        await loadHqBrandOptions();
        await loadHqCategoryOptions();
        await loadProducts();
        startEditProduct(p);
      });
      actionsTd.appendChild(editBtn);

      const delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'link';
      delBtn.textContent = '削除';
      delBtn.addEventListener('click', async () => {
        const warning = `「${p.name}」(${p.code})を削除します。\n\n` +
          'これまでの在庫数の記録は消えませんが、商品マスタからは無くなり、' +
          '次にこのバーコードをスキャンすると「未登録」になります。\n\n' +
          '本当に削除しますか?';
        if (!confirm(warning)) return;
        await apiCall('deleteProduct', { store, code: p.code });
        await runDedupScan(store);
      });
      actionsTd.appendChild(delBtn);

      tr.appendChild(actionsTd);
      table.appendChild(tr);
    });
    card.appendChild(table);
    container.appendChild(card);
  });
}

// ---- スタッフ管理(本社) ----
document.getElementById('staff-store-select').addEventListener('change', loadStaff);

async function loadStaff() {
  const store = document.getElementById('staff-store-select').value;
  if (!store) return;
  const data = await apiCall('getStaffList', { store });
  const container = document.getElementById('staff-table');
  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>スタッフ名</th><th></th></tr>';
  data.staff.forEach((s) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${s.name}</td><td><button type="button" class="link" data-name="${s.name}">無効化</button></td>`;
    table.appendChild(tr);
  });
  container.appendChild(table);
  container.querySelectorAll('button[data-name]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      await apiCall('setStaffActive', { store, name: btn.dataset.name, active: false });
      await loadStaff();
    });
  });
}

document.getElementById('hq-btn-add-staff').addEventListener('click', async () => {
  const store = document.getElementById('staff-store-select').value;
  const name = document.getElementById('hq-new-staff-name').value.trim();
  if (!store || !name) return;
  await apiCall('addStaff', { store, name });
  document.getElementById('hq-new-staff-name').value = '';
  await loadStaff();
});

// ---- アカウント管理 ----
document.getElementById('btn-create-account').addEventListener('click', async () => {
  const statusEl = document.getElementById('account-status');
  try {
    await apiCall('createAccount', {
      store: document.getElementById('a-store').value.trim(),
      username: document.getElementById('a-username').value.trim(),
      password: document.getElementById('a-password').value,
      role: document.getElementById('a-role').value
    });
    statusEl.textContent = '作成しました';
    ['a-store', 'a-username', 'a-password'].forEach((id) => (document.getElementById(id).value = ''));
  } catch (e) {
    statusEl.textContent = e.message;
  }
});

document.getElementById('btn-reset-password').addEventListener('click', async () => {
  const statusEl = document.getElementById('reset-status');
  try {
    await apiCall('resetPassword', {
      username: document.getElementById('r-username').value.trim(),
      newPassword: document.getElementById('r-password').value
    });
    statusEl.textContent = 'リセットしました';
    ['r-username', 'r-password'].forEach((id) => (document.getElementById(id).value = ''));
  } catch (e) {
    statusEl.textContent = e.message;
  }
});

// ---- 取引ログ ----
document.getElementById('log-store-filter').addEventListener('change', loadLogs);
document.getElementById('log-type-filter').addEventListener('change', loadLogs);
document.getElementById('btn-search-logs').addEventListener('click', loadLogs);
// スタッフ名・ブランドは入力しながら絞り込むと打鍵のたびに通信してしまうため、Enterキーでも検索できるようにする
['log-staff-filter', 'log-brand-filter'].forEach((id) => {
  document.getElementById(id).addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') loadLogs();
  });
});

document.getElementById('btn-clear-store-log').addEventListener('click', () => {
  const store = document.getElementById('log-store-filter').value;
  const type = document.getElementById('log-type-filter').value;
  if (!store) {
    alert('店舗を選択してください(全店舗のログを一括削除することはできません)');
    return;
  }
  const target = type ? `${store}の「${type}」の取引ログ` : `${store}の取引ログ全て`;
  openDangerModal(
    `${target}を削除します。この操作は取り消せません。テストデータの整理などに使ってください。`,
    async (password) => {
      await apiCall('clearStoreLog', { store, type: type || undefined, password });
      await loadLogs();
    }
  );
});

async function loadLogs() {
  const store = document.getElementById('log-store-filter').value;
  const type = document.getElementById('log-type-filter').value;
  const staffName = document.getElementById('log-staff-filter').value.trim();
  const brand = document.getElementById('log-brand-filter').value.trim();
  const data = await apiCall('getLogEntries', {
    store: store || undefined,
    type: type || undefined,
    staffName: staffName || undefined,
    brand: brand || undefined,
    limit: 200
  });
  const container = document.getElementById('logs-table');
  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>日時</th><th>店舗</th><th>スタッフ</th><th>ブランド</th><th>商品</th><th>種別</th><th>数量</th><th></th></tr>';
  data.logs.forEach((log) => {
    const tr = document.createElement('tr');
    const delBtn = `<button type="button" class="link" data-row="${log.rowIndex}">削除</button>`;
    tr.innerHTML = `<td>${new Date(log.timestamp).toLocaleString('ja-JP')}</td><td>${log.store}</td><td>${log.staffName}</td><td>${log.brand || ''}</td><td>${log.name}</td><td>${log.type}</td><td>${log.quantity}</td><td>${delBtn}</td>`;
    table.appendChild(tr);
  });
  container.appendChild(table);
  container.querySelectorAll('button[data-row]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      if (!confirm('このログを削除しますか?(在庫集計に影響します)')) return;
      await apiCall('deleteLogEntry', { rowIndex: Number(btn.dataset.row) });
      await loadLogs();
    });
  });
}

// ---- 棚卸承認(本社・店舗共通の表示ロジック) ----
// 「先月末在庫 + 今月の入荷 − 今月の廃棄」と「今月末在庫(棚卸後)」を商品ごとに突き合わせ、
// 差異(0以外)がある行を強調表示する。承認は本社のみ行える(承認ボタンはHQ画面にしかない)。
/** "2026-07" のような月文字列を「7月末」の表記にする(先月末・今月末の列見出し用)。 */
function monthLabel(monthStr) {
  const parts = (monthStr || '').split('-');
  return parts[1] ? `${Number(parts[1])}月末` : '';
}

/**
 * 「今の状況」(状態バッジ)をまず大きく分かりやすく出し、実施記録の一覧・商品ごとの
 * 差異テーブルといった細かい情報は、必要な人だけがボタンで開けるようにする(すべて
 * 一度に表示すると情報が多すぎて分かりにくい、との指摘への対応)。
 */
function renderReviewResult(data, summaryElId, tableElId) {
  const summaryEl = document.getElementById(summaryElId);
  const tableContainer = document.getElementById(tableElId);

  const ms = data.monthlyStatus;
  let statusHtml;
  if (ms) {
    const dateStr = ms.implementedDate ? new Date(ms.implementedDate).toLocaleDateString('ja-JP') : '';
    // 「未実施・承認待ち・差し戻し・承認済み」を並べて表示し、今の状態だけを色付けして
    // 一目で分かるようにする(差し戻しは注意が必要な状態として赤色にする)。
    const steps = [
      { key: '未実施', label: '未実施' },
      { key: '承認待ち', label: '承認待ち' },
      { key: '差し戻し', label: '差し戻し' },
      { key: '承認済み', label: '承認済み' }
    ];
    const stepsHtml = steps.map((s) => {
      const isActive = s.key === ms.status;
      const cls = 'step' + (isActive ? ' active' : '') + (isActive && s.key === '差し戻し' ? ' warn' : '');
      return `<span class="${cls}">${s.label}</span>`;
    }).join('');
    statusHtml = `<div class="status-steps">${stepsHtml}</div>`;
    if (ms.status === '承認済み') {
      statusHtml += `<p class="hint" style="margin-top:0;">実施日: ${dateStr} / ${ms.approver}が承認済み</p>`;
    } else if (ms.status === '差し戻し') {
      statusHtml += `<p class="hint" style="margin-top:0;">実施日: ${dateStr} / 理由: ${ms.rejectedReason || '記載なし'}</p>`;
    } else if (ms.status === '承認待ち') {
      statusHtml += `<p class="hint" style="margin-top:0;">実施日: ${dateStr}</p>`;
    }
  } else if (data.approval.approved) {
    // 旧レスポンス互換(monthlyStatusが無い場合)
    statusHtml = `<p class="status">承認済み(${data.approval.approver} / ${new Date(data.approval.approvedAt).toLocaleString('ja-JP')})</p>`;
  } else {
    statusHtml = '<p class="hint">未承認</p>';
  }

  let eventsHtml;
  if (!data.stocktakeEvents.length) {
    eventsHtml = '<p class="hint">この月はまだ棚卸が実施されていません。</p>';
  } else {
    const lines = data.stocktakeEvents.map((e) => {
      const label = `${new Date(e.timestamp).toLocaleString('ja-JP')} / ${e.staffName || '(不明)'} / ${e.itemCount}品目`;
      return e.isLatest ? `<strong>${label}(最新)</strong>` : label;
    });
    eventsHtml = `<p class="hint">棚卸実施(1回の送信ごとに1行、古い順): ${lines.join('、')}</p>`;
  }

  const detailsId = summaryElId + '-details';
  const toggleId = summaryElId + '-toggle';
  summaryEl.innerHTML = statusHtml +
    `<button type="button" class="secondary" id="${toggleId}">詳細(実施記録・商品ごとの差異)を確認する</button>` +
    `<div id="${detailsId}" style="display:none;">${eventsHtml}</div>`;
  tableContainer.style.display = 'none';

  document.getElementById(toggleId).addEventListener('click', () => {
    const detailsEl = document.getElementById(detailsId);
    const isHidden = detailsEl.style.display === 'none';
    detailsEl.style.display = isHidden ? 'block' : 'none';
    tableContainer.style.display = isHidden ? 'block' : 'none';
    document.getElementById(toggleId).textContent = isHidden ? '詳細を隠す' : '詳細(実施記録・商品ごとの差異)を確認する';
  });

  tableContainer.innerHTML = '';
  if (!data.items.length) {
    tableContainer.textContent = '商品が登録されていません';
    return;
  }
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML =
    `<tr><th>ブランド</th><th>品名</th><th>カラーNO</th><th>${monthLabel(data.previousMonth)}</th>` +
    `<th>入荷</th><th>廃棄</th><th>${monthLabel(data.month)}</th><th>差異</th></tr>`;
  data.items.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.diff !== 0) tr.className = 'out-of-stock';
    tr.innerHTML = `<td>${item.brand || ''}</td><td>${item.name}</td><td>${item.colorNo || ''}</td>` +
      `<td>${item.prevStock}</td><td>${item.incoming}</td><td>${item.disposal}</td>` +
      `<td>${item.currentStock}</td><td>${item.diff}</td>`;
    table.appendChild(tr);
  });
  tableContainer.appendChild(table);
}

// ---- 棚卸承認(本社) ----
let currentHqReviewData = null;

document.getElementById('btn-load-hq-review').addEventListener('click', async () => {
  const store = document.getElementById('hq-review-store-select').value;
  const month = document.getElementById('hq-review-month').value;
  if (!store || !month) return;
  try {
    currentHqReviewData = await apiCall('getStocktakeReview', { store, month });
    renderReviewResult(currentHqReviewData, 'hq-review-summary', 'hq-review-table');
    const approveBtn = document.getElementById('btn-approve-review');
    approveBtn.style.display = 'block';
    approveBtn.textContent = currentHqReviewData.approval.approved ? '再承認する' : '承認する';
    document.getElementById('btn-reject-review').style.display = 'block';
    document.getElementById('btn-export-hq-review').style.display = 'inline-block';
  } catch (e) {
    document.getElementById('hq-review-summary').textContent = e.message;
  }
});

document.getElementById('btn-export-hq-review').addEventListener('click', () => {
  if (!currentHqReviewData) return;
  const rows = [[
    'ブランド', '品名', 'カラーNO',
    monthLabel(currentHqReviewData.previousMonth), '入荷', '廃棄',
    monthLabel(currentHqReviewData.month), '差異'
  ]];
  currentHqReviewData.items.forEach((item) => {
    rows.push([
      item.brand || '', item.name, item.colorNo || '',
      item.prevStock, item.incoming, item.disposal, item.currentStock, item.diff
    ]);
  });
  downloadCsv(`棚卸承認_${currentHqReviewData.store}_${currentHqReviewData.month}.csv`, rows);
});

document.getElementById('btn-approve-review').addEventListener('click', async () => {
  if (!currentHqReviewData) return;
  try {
    await apiCall('approveStocktake', { store: currentHqReviewData.store, month: currentHqReviewData.month });
    // 承認したら内容は表示したままにせず、ダッシュボードに戻る
    currentHqReviewData = null;
    document.getElementById('hq-review-summary').innerHTML = '';
    document.getElementById('hq-review-table').innerHTML = '';
    document.getElementById('btn-approve-review').style.display = 'none';
    document.getElementById('btn-reject-review').style.display = 'none';
    document.getElementById('btn-export-hq-review').style.display = 'none';
    showScreen('screen-dashboard');
  } catch (e) {
    document.getElementById('hq-review-summary').textContent = e.message;
  }
});

document.getElementById('btn-reject-review').addEventListener('click', async () => {
  if (!currentHqReviewData) return;
  const reason = prompt('差し戻す理由(店舗に伝わります。空欄でも構いません)') || '';
  try {
    await apiCall('rejectStocktake', { store: currentHqReviewData.store, month: currentHqReviewData.month, reason });
    currentHqReviewData = null;
    document.getElementById('hq-review-summary').innerHTML = '';
    document.getElementById('hq-review-table').innerHTML = '';
    document.getElementById('btn-approve-review').style.display = 'none';
    document.getElementById('btn-reject-review').style.display = 'none';
    document.getElementById('btn-export-hq-review').style.display = 'none';
    showScreen('screen-dashboard');
  } catch (e) {
    document.getElementById('hq-review-summary').textContent = e.message;
  }
});

// ---- 棚卸承認の確認(店舗。閲覧のみ) ----
document.getElementById('btn-nav-store-review').addEventListener('click', async () => {
  showScreen('screen-store-review');
  document.getElementById('store-review-month').value = currentYearMonth();
  document.getElementById('store-review-summary').innerHTML = '';
  document.getElementById('store-review-table').innerHTML = '';
  document.getElementById('store-review-logs-table').innerHTML = '';
  await loadStoreReviewStaffOptions();
});

async function loadStoreReviewStaffOptions() {
  const select = document.getElementById('store-review-staff-filter');
  select.innerHTML = '<option value="">全員</option>';
  try {
    const data = await apiCall('getStaffList', {});
    data.staff.forEach((s) => {
      const opt = document.createElement('option');
      opt.value = s.name;
      opt.textContent = s.name;
      select.appendChild(opt);
    });
  } catch (e) {
    console.error(e);
  }
  // ログイン時に選んだ実施者名を初期値にしておく(「自分の」データをすぐ検索できるように)
  select.value = state.staffName || '';
}

document.getElementById('btn-search-store-review-logs').addEventListener('click', async () => {
  const staffName = document.getElementById('store-review-staff-filter').value.trim();
  const container = document.getElementById('store-review-logs-table');
  container.innerHTML = '検索中...';
  try {
    const data = await apiCall('getLogEntries', { type: '棚卸', staffName: staffName || undefined, limit: 200 });
    container.innerHTML = '';
    if (!data.logs.length) {
      container.textContent = '該当するデータがありません';
      return;
    }
    const table = document.createElement('table');
    table.className = 'stock-table';
    table.innerHTML = '<tr><th>日時</th><th>実施者</th><th>ブランド</th><th>商品</th><th>数量</th></tr>';
    data.logs.forEach((log) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${new Date(log.timestamp).toLocaleString('ja-JP')}</td><td>${log.staffName}</td><td>${log.brand || ''}</td><td>${log.name}</td><td>${log.quantity}</td>`;
      table.appendChild(tr);
    });
    container.appendChild(table);
  } catch (e) {
    container.textContent = e.message;
  }
});

document.getElementById('btn-back-store-review').addEventListener('click', () => showScreen('screen-menu'));

document.getElementById('btn-load-store-review').addEventListener('click', async () => {
  const month = document.getElementById('store-review-month').value;
  if (!month) return;
  try {
    const data = await apiCall('getStocktakeReview', { month });
    renderReviewResult(data, 'store-review-summary', 'store-review-table');
  } catch (e) {
    document.getElementById('store-review-summary').textContent = e.message;
  }
});

// ---- 初期化(ログイン済みなら保存されたroleに応じて自動的に振り分ける) ----
(async function init() {
  if (!state.token) return;
  try {
    if (state.role === 'hq') {
      await enterDashboard();
    } else if (state.store) {
      await goToStaffSelect();
    }
  } catch (e) {
    logout();
  }
})();
