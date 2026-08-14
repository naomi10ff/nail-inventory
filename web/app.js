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
  stores: [] // 本社ダッシュボードで使用
};

let scannerStocktake = null;
let scannerHqIncoming = null;
let scannerHqDisposal = null;

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

// ---- 総在庫一覧から、ブランド名・品名・カラーNO・在庫数をまとめて直せるようにする ----
// パスワード再入力を必須にする点はdanger-modalと同じ。
function openAdjustModal(item, onConfirm) {
  const overlay = document.getElementById('adjust-modal');
  const codeInput = document.getElementById('adjust-modal-code');
  const brandInput = document.getElementById('adjust-modal-brand');
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
  brandInput.value = item.brand || '';
  nameInput.value = item.name || '';
  colorInput.value = item.colorNo || '';
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
        newCode, brand: brandInput.value.trim(), name, colorNo: colorInput.value.trim(),
        newStock, memo: memoInput.value.trim(), password
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
document.getElementById('btn-nav-stocktake').addEventListener('click', () => {
  state.tally = {};
  loadStocktakeProductList().catch((e) => console.error(e));
  renderTally();
  showScreen('screen-stocktake');
  rearmGate(stocktakeGate);
  startScanner('reader', onStocktakeScan, stocktakeGate).catch((e) => console.error(e));
});

// ---- 棚卸: 未スキャン商品一覧 ----
// 「内容を確認する」を押した時点で送信はせず、まず商品マスタの全件と今スキャン済みの
// 一覧を突き合わせて未スキャン商品(欠品・スキャン漏れの可能性がある)を確認してもらい、
// 「本社へ送信する」を押した時点で初めてsubmitStocktakeを呼ぶ2段階にしている。
let stocktakeAllProducts = [];

async function loadStocktakeProductList() {
  stocktakeAllProducts = [];
  document.getElementById('unscanned-container').innerHTML = '';
  document.getElementById('stocktake-review').style.display = 'none';
  const data = await apiCall('listProducts', {});
  stocktakeAllProducts = data.products;
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
    byBrand[brand].forEach((p) => {
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
  showScreen('screen-stock-lookup');
  rearmGate(stockLookupGate);
  startScanner('reader-4', onStockLookupScan, stockLookupGate).catch((e) => console.error(e));
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
      return;
    }
    document.getElementById('stock-lookup-status').textContent = '';
    document.getElementById('stock-lookup-result').style.display = 'block';
    document.getElementById('stock-lookup-name').textContent = result.name;
    document.getElementById('stock-lookup-brand').textContent = result.brand || '';
    document.getElementById('stock-lookup-colorno').textContent = result.colorNo || '';
    document.getElementById('stock-lookup-count').textContent =
      result.currentStock + '本' + (result.outOfStock ? '(欠品)' : '');
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
function createGate(options) {
  return {
    armed: true,
    missCount: 0,
    rearmOnMiss: !!options.rearmOnMiss,
    pendingRearm: false,
    requiredMisses: options.requiredMisses || 8
  };
}

function rearmGate(gate) {
  gate.armed = true;
  gate.missCount = 0;
  gate.pendingRearm = false;
}

function rearmGateAfterMiss(gate) {
  gate.pendingRearm = true;
  gate.missCount = 0;
}

const stocktakeGate = createGate({ rearmOnMiss: true, requiredMisses: 8 });
const hqIncomingGate = createGate({ rearmOnMiss: false });
const hqDisposalGate = createGate({ rearmOnMiss: false });

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

  await scanner.start(
    { facingMode: 'environment' },
    {
      fps: 10,
      // 正方形だと横長のバーコードに対して余白が多く読み取りにくいため、横長の矩形にする
      qrbox: (viewfinderWidth, viewfinderHeight) => ({
        width: Math.max(240, Math.floor(viewfinderWidth * 0.85)),
        height: Math.max(120, Math.floor(viewfinderHeight * 0.4))
      }),
      // videoConstraintsを指定すると外側のfacingMode指定が無視されてしまうため、
      // ここに背面カメラ指定を含めておく(含めないとインカメラになることがある)。
      // iOSは既定だと解像度が低く、バーコードの細い線がつぶれやすいため高解像度も要求する
      videoConstraints: {
        facingMode: { exact: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 }
      }
    },
    (decodedText) => {
      if (!gate.armed) {
        gate.missCount = 0; // まだ画面内に写っている(検出できている)ので外れた判定をリセット
        return;
      }
      gate.armed = false;
      gate.missCount = 0;
      onSuccess(decodedText);
    },
    () => {
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
    const product = await apiCall('lookupProduct', { code });
    if (!product) {
      document.getElementById('stocktake-status').textContent = '未登録のコードです: ' + code;
      return;
    }
    if (!state.tally[code]) {
      state.tally[code] = { code, name: product.name, brand: product.brand || '(ブランド未設定)', count: 0 };
    }
    state.tally[code].count += 1;
    document.getElementById('stocktake-status').textContent = '';
    renderTally();
  } catch (e) {
    document.getElementById('stocktake-status').textContent = e.message;
  }
}

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
    byBrand[brand].forEach((item) => {
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
  } else {
    summaryEl.textContent = `${items.length}品目 / 合計${totalCount}本をスキャン済み。よろしければ内容を確認してください`;
    confirmBtn.disabled = false;
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
  document.getElementById('stocktake-review').style.display = 'block';
});

document.getElementById('btn-back-to-scan').addEventListener('click', () => {
  document.getElementById('stocktake-review').style.display = 'none';
});

document.getElementById('btn-send-stocktake').addEventListener('click', async () => {
  const items = Object.values(state.tally).map((t) => ({ code: t.code, count: t.count }));
  try {
    const data = await apiCall('submitStocktake', { staffName: state.staffName, items });
    document.getElementById('stocktake-status').textContent =
      `棚卸を送信しました(${data.recorded.length}品目)。` +
      (data.unknownCodes.length ? ` 未登録コード: ${data.unknownCodes.join(', ')}` : '');
    state.tally = {};
    renderTally();
    document.getElementById('stocktake-review').style.display = 'none';
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
    renderInventoryList(document.getElementById('inventory-search').value);
  } catch (e) {
    container.textContent = e.message;
  }
}

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

  // スペース区切りの複数キーワードはすべて満たす(AND検索)商品だけに絞り込む
  const keywords = normalizeSearchText(query).split(/\s+/).filter(Boolean);
  const filtered = inventoryItems.filter((item) => {
    if (!keywords.length) return true;
    const haystack = normalizeSearchText([item.brand, item.name, item.colorNo].filter(Boolean).join(' '));
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
document.getElementById('nav-products').addEventListener('click', async () => {
  showScreen('screen-products');
  document.getElementById('product-search').value = '';
  await loadHqBrandOptions();
  await loadHqCategoryOptions();
  await loadProducts();
});
document.getElementById('product-store-select').addEventListener('change', async () => {
  cancelEditProduct();
  document.getElementById('product-search').value = '';
  await loadHqBrandOptions();
  await loadHqCategoryOptions();
  await loadProducts();
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
  document.getElementById('btn-export-hq-review').style.display = 'none';
});

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
  document.getElementById('hq-incoming-new-color-no').value = product.colorNo || '';
  document.getElementById('hq-incoming-new-item-number').value = product.itemNumber || '';
  setHqIncomingCategoryValue(product.category || '');
});

document.getElementById('btn-new-name-toggle-incoming').addEventListener('click', () => {
  document.getElementById('hq-incoming-new-name-form').style.display = 'block';
});

function resetHqIncomingNameNewForm() {
  document.getElementById('hq-incoming-new-name-form').style.display = 'none';
  document.getElementById('hq-incoming-new-name').value = '';
}

/** 選択中の品名が、既存商品(バーコードを紐づけるだけ)か新規入力かを返す。 */
function currentHqIncomingNameSelection() {
  const newForm = document.getElementById('hq-incoming-new-name-form');
  const newValue = document.getElementById('hq-incoming-new-name').value.trim();
  if (newForm.style.display !== 'none' && newValue) {
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
function resetHqIncomingScreen() {
  document.getElementById('hq-incoming-known').style.display = 'none';
  document.getElementById('hq-incoming-unknown').style.display = 'none';
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
    if (selection.isNew) {
      // 新規商品として登録(このバーコードで新しいコードが割り当てられる)
      name = selection.name;
      await apiCall('registerProduct', {
        store, code: hqIncomingScannedCode, brand, name, colorNo, itemNumber, category
      });
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
    await loadHqIncomingBrandOptions();
    await loadHqIncomingProductList();
  } catch (e) {
    document.getElementById('hq-incoming-status').textContent = e.message;
  }
  });
});

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
  await renderInventoryTable(document.getElementById('store-filter').value);
}

document.getElementById('store-filter').addEventListener('change', (e) => renderInventoryTable(e.target.value));

let hqInventoryForExport = [];

async function renderInventoryTable(store) {
  const data = await apiCall('getInventorySummary', { store: store || undefined });
  hqInventoryForExport = data.items;
  const container = document.getElementById('hq-inventory-container');
  container.innerHTML = '';
  if (!data.items.length) {
    container.textContent = '在庫データがありません';
    return;
  }
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>店舗</th><th>ブランド</th><th>品名</th><th>カラーNO</th><th>現在庫</th><th colspan="2"></th></tr>';
  data.items.forEach((item) => {
    const tr = document.createElement('tr');
    if (item.outOfStock) tr.className = 'out-of-stock';
    tr.innerHTML = `<td>${item.store}</td><td>${item.brand || ''}</td><td>${item.name}</td><td>${item.colorNo || ''}</td><td>${item.currentStock}</td>`;

    const label = `${item.store}の「${item.brand || ''} ${item.name}」`;

    const adjustTd = document.createElement('td');
    const adjustBtn = document.createElement('button');
    adjustBtn.type = 'button';
    adjustBtn.className = 'link';
    adjustBtn.textContent = '修正';
    adjustBtn.addEventListener('click', () => {
      openAdjustModal(item, async ({ newCode, brand, name, colorNo, newStock, memo, password }) => {
        const productsData = await apiCall('listProducts', { store: item.store });
        const product = productsData.products.find((p) => String(p.code) === String(item.code));
        await apiCall('updateProduct', {
          store: item.store, code: item.code, newCode, brand, name, colorNo,
          category: product ? product.category : '', memo: product ? product.memo : ''
        });
        const finalCode = newCode && newCode !== String(item.code) ? newCode : item.code;
        if (Number(newStock) !== item.currentStock) {
          await apiCall('adjustProductStock', { store: item.store, code: finalCode, newStock, memo, password });
        }
        await loadTotalInventoryScreen();
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

async function loadProducts() {
  const store = document.getElementById('product-store-select').value;
  const container = document.getElementById('products-table');
  if (!store) {
    currentProductList = [];
    container.textContent = '店舗を選択してください';
    return;
  }
  const data = await apiCall('listProducts', { store });
  currentProductList = data.products;
  renderProductsTable(document.getElementById('product-search').value);
}

function renderProductsTable(query) {
  const store = document.getElementById('product-store-select').value;
  const container = document.getElementById('products-table');
  if (!currentProductList.length) {
    container.textContent = '商品がまだ登録されていません';
    productsFilteredForExport = [];
    return;
  }

  // スペース区切りの複数キーワードはすべて満たす(AND検索)商品だけに絞り込む
  const keywords = normalizeSearchText(query).split(/\s+/).filter(Boolean);
  const filtered = currentProductList.filter((p) => {
    if (!keywords.length) return true;
    const haystack = normalizeSearchText([p.code, p.itemNumber, p.brand, p.name, p.colorNo].filter(Boolean).join(' '));
    return keywords.every((kw) => haystack.includes(kw));
  });
  productsFilteredForExport = filtered;

  if (!filtered.length) {
    container.textContent = '該当する商品が見つかりません';
    return;
  }

  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>コード</th><th>品番</th><th>ブランド</th><th>品名</th><th>カラーNO</th><th>メモ</th><th></th></tr>';
  filtered.forEach((p) => {
    const tr = document.createElement('tr');
    const editBtn = `<button type="button" class="link" data-edit-code="${p.code}">編集</button>`;
    const qrBtn = `<button type="button" class="link" data-qr-code="${p.code}">QR表示</button>`;
    const delBtn = `<button type="button" class="link" data-code="${p.code}">削除</button>`;
    tr.innerHTML = `<td>${p.code}</td><td>${p.itemNumber || ''}</td><td>${p.brand || ''}</td><td>${p.name}</td><td>${p.colorNo || ''}</td><td>${p.memo || ''}</td><td>${editBtn} ${qrBtn} ${delBtn}</td>`;
    table.appendChild(tr);
  });
  container.appendChild(table);

  container.querySelectorAll('button[data-code]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      if (!confirm('削除しますか?')) return;
      await apiCall('deleteProduct', { store, code: btn.dataset.code });
      await loadProducts();
    });
  });
  container.querySelectorAll('button[data-edit-code]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const product = currentProductList.find((p) => p.code === btn.dataset.editCode);
      if (product) startEditProduct(product);
    });
  });
  container.querySelectorAll('button[data-qr-code]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const product = currentProductList.find((p) => p.code === btn.dataset.qrCode);
      if (product) showQrViewModal(product.code, product.name, product.brand);
    });
  });
}

document.getElementById('product-search').addEventListener('input', () => {
  renderProductsTable(document.getElementById('product-search').value);
});

document.getElementById('btn-export-products').addEventListener('click', () => {
  if (!productsFilteredForExport.length) return;
  const store = document.getElementById('product-store-select').value;
  const rows = [['ブランド', '品名', 'カラーNO', '品番(商品コード)', 'バーコード']];
  productsFilteredForExport.forEach((p) => {
    rows.push([p.brand || '', p.name, p.colorNo || '', p.itemNumber || '', p.code]);
  });
  downloadCsv(`商品マスタ_${store}.csv`, rows);
});

/** 商品一覧から、既に登録済みの商品のQRコードをいつでも呼び出して印刷できるようにする。 */
async function showQrViewModal(code, name, brand) {
  const holder = document.getElementById('qr-view-canvas-holder');
  holder.innerHTML = '';
  const canvas = document.createElement('canvas');
  holder.appendChild(canvas);
  await QRCode.toCanvas(canvas, code, { width: 220 });
  document.getElementById('qr-view-label').textContent = code + ' / ' + (brand ? brand + ' ' : '') + name;
  document.getElementById('qr-view-modal').style.display = 'flex';
}

document.getElementById('btn-print-qr-view').addEventListener('click', () => window.print());
document.getElementById('btn-close-qr-view').addEventListener('click', () => {
  document.getElementById('qr-view-modal').style.display = 'none';
});

function startEditProduct(product) {
  editingProductCode = product.code;
  document.getElementById('product-form-title').textContent = '商品を編集';
  document.getElementById('p-code').value = product.code;
  setPBrandValue(product.brand || '');
  document.getElementById('p-name').value = product.name || '';
  document.getElementById('p-color-no').value = product.colorNo || '';
  document.getElementById('p-item-number').value = product.itemNumber || '';
  setPCategoryValue(product.category || '');
  document.getElementById('p-memo').value = product.memo || '';
  document.getElementById('btn-add-product').textContent = '更新する';
  document.getElementById('btn-cancel-edit-product').style.display = 'block';
  document.getElementById('p-qr-result').style.display = 'none';
}

function cancelEditProduct() {
  editingProductCode = null;
  document.getElementById('product-form-title').textContent = '新規登録';
  ['p-code', 'p-name', 'p-color-no', 'p-item-number', 'p-memo'].forEach((id) => (document.getElementById(id).value = ''));
  resetPBrandNewForm();
  document.getElementById('p-brand-select').selectedIndex = 0;
  resetPCategoryNewForm();
  document.getElementById('p-category-select').selectedIndex = 0;
  document.getElementById('btn-add-product').textContent = '登録';
  document.getElementById('btn-cancel-edit-product').style.display = 'none';
  document.getElementById('p-qr-result').style.display = 'none';
}

document.getElementById('btn-cancel-edit-product').addEventListener('click', cancelEditProduct);

document.getElementById('btn-add-product').addEventListener('click', async (e) => {
  const statusEl = document.getElementById('product-status');
  const store = document.getElementById('product-store-select').value;
  if (!store) {
    statusEl.textContent = '店舗を選択してください';
    return;
  }
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
    if (editingProductCode) {
      await apiCall('updateProduct', Object.assign({ code: editingProductCode, newCode: codeInput }, payload));
      statusEl.textContent = '更新しました';
      cancelEditProduct();
    } else {
      document.getElementById('p-qr-result').style.display = 'none';
      const result = await apiCall('registerProduct', Object.assign({ code: codeInput || undefined }, payload));
      statusEl.textContent = '登録しました';
      ['p-code', 'p-name', 'p-color-no', 'p-item-number', 'p-memo'].forEach((id) => (document.getElementById(id).value = ''));
      resetPBrandNewForm();
      resetPCategoryNewForm();
      document.getElementById('p-category-select').selectedIndex = 0;
      await loadHqBrandOptions();

      if (!codeInput) {
        const holder = document.getElementById('p-qr-canvas-holder');
        holder.innerHTML = '';
        const canvas = document.createElement('canvas');
        holder.appendChild(canvas);
        await QRCode.toCanvas(canvas, result.code, { width: 220 });
        const label = document.createElement('p');
        label.textContent = result.code + ' / ' + result.name;
        holder.appendChild(label);
        document.getElementById('p-qr-result').style.display = 'block';
      }
    }
    await loadProducts();
  } catch (e) {
    statusEl.textContent = e.message;
  }
  });
});

document.getElementById('btn-print-p-qr').addEventListener('click', () => window.print());

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

document.getElementById('btn-clear-store-log').addEventListener('click', () => {
  const store = document.getElementById('log-store-filter').value;
  if (!store) {
    alert('店舗を選択してください(全店舗のログを一括削除することはできません)');
    return;
  }
  openDangerModal(
    `${store}の取引ログを全て削除します。この操作は取り消せません。テストデータの整理などに使ってください。`,
    async (password) => {
      await apiCall('clearStoreLog', { store, password });
      await loadLogs();
    }
  );
});

async function loadLogs() {
  const store = document.getElementById('log-store-filter').value;
  const data = await apiCall('getLogEntries', { store: store || undefined, limit: 200 });
  const container = document.getElementById('logs-table');
  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'stock-table';
  table.innerHTML = '<tr><th>日時</th><th>店舗</th><th>スタッフ</th><th>商品</th><th>種別</th><th>数量</th><th></th></tr>';
  data.logs.forEach((log) => {
    const tr = document.createElement('tr');
    const delBtn = `<button type="button" class="link" data-row="${log.rowIndex}">削除</button>`;
    tr.innerHTML = `<td>${new Date(log.timestamp).toLocaleString('ja-JP')}</td><td>${log.store}</td><td>${log.staffName}</td><td>${log.name}</td><td>${log.type}</td><td>${log.quantity}</td><td>${delBtn}</td>`;
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

function renderReviewResult(data, summaryElId, tableElId) {
  const summaryEl = document.getElementById(summaryElId);
  let summaryHtml = '';
  if (!data.stocktakeEvents.length) {
    summaryHtml += '<p class="hint">この月はまだ棚卸が実施されていません。</p>';
  } else {
    const lines = data.stocktakeEvents.map((e) => {
      const label = `${new Date(e.timestamp).toLocaleString('ja-JP')} / ${e.staffName || '(不明)'} / ${e.itemCount}品目`;
      return e.isLatest ? `<strong>${label}(最新)</strong>` : label;
    });
    summaryHtml += `<p class="hint">棚卸実施(1回の送信ごとに1行、古い順): ${lines.join('、')}</p>`;
  }
  if (data.approval.approved) {
    summaryHtml += `<p class="status">承認済み(${data.approval.approver} / ${new Date(data.approval.approvedAt).toLocaleString('ja-JP')})</p>`;
  } else {
    summaryHtml += '<p class="hint">未承認</p>';
  }
  summaryEl.innerHTML = summaryHtml;

  const container = document.getElementById(tableElId);
  container.innerHTML = '';
  if (!data.items.length) {
    container.textContent = '商品が登録されていません';
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
  container.appendChild(table);
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
    document.getElementById('btn-export-hq-review').style.display = 'none';
    showScreen('screen-dashboard');
  } catch (e) {
    document.getElementById('hq-review-summary').textContent = e.message;
  }
});

// ---- 棚卸承認の確認(店舗。閲覧のみ) ----
document.getElementById('btn-nav-store-review').addEventListener('click', () => {
  showScreen('screen-store-review');
  document.getElementById('store-review-month').value = currentYearMonth();
  document.getElementById('store-review-summary').innerHTML = '';
  document.getElementById('store-review-table').innerHTML = '';
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
