/**
 * 店舗/本社アカウントの簡易認証とセッション管理。
 * Googleアカウント認証は使わず、アカウントシートに対して照合する。
 */

var SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30日間はログイン状態を保持

function generateSalt_() {
  return Utilities.getUuid();
}

function hashPassword_(password, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + ':' + salt);
  return raw
    .map(function (b) {
      var v = b < 0 ? b + 256 : b;
      return ('0' + v.toString(16)).slice(-2);
    })
    .join('');
}

function generateToken_() {
  return Utilities.getUuid() + Utilities.getUuid();
}

function login_(username, password) {
  if (!username || !password) throw new Error('ユーザー名とパスワードを入力してください');

  var sheet = getSheet_(SHEET_ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var store = data[i][0], user = data[i][1], hash = data[i][2], salt = data[i][3], role = data[i][4];
    if (user === username) {
      if (hashPassword_(password, salt) === hash) {
        var token = createSession_(username, store, role);
        return { token: token, store: store, role: role };
      }
      throw new Error('パスワードが違います');
    }
  }
  throw new Error('ユーザー名が見つかりません');
}

function createSession_(username, store, role) {
  var sheet = getSheet_(SHEET_SESSIONS);
  var token = generateToken_();
  sheet.appendRow([token, username, store, role, new Date()]);
  return token;
}

function validateToken_(token) {
  if (!token) throw new Error('未ログインです');
  var sheet = getSheet_(SHEET_SESSIONS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var issuedAt = new Date(data[i][4]).getTime();
      if (Date.now() - issuedAt > SESSION_MAX_AGE_MS) {
        throw new Error('セッションの有効期限が切れました。再ログインしてください');
      }
      return { username: data[i][1], store: data[i][2], role: data[i][3] };
    }
  }
  throw new Error('セッションが無効です。再ログインしてください');
}

function requireRole_(session, allowedRoles) {
  if (allowedRoles.indexOf(session.role) === -1) {
    throw new Error('この操作を行う権限がありません');
  }
}

/** hqはどの店舗にもアクセス可能。storeロールは自店以外にアクセス不可。 */
function requireStoreAccess_(session, targetStore) {
  if (session.role === 'hq') return;
  if (session.store !== targetStore) {
    throw new Error('他店舗のデータにはアクセスできません');
  }
}

// ---- 本社によるアカウント管理 ----

function createAccount_(session, p) {
  requireRole_(session, ['hq']);
  if (!p.store || !p.username || !p.password || !p.role) {
    throw new Error('店舗・ユーザー名・パスワード・権限を入力してください');
  }
  var sheet = getSheet_(SHEET_ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === p.username) throw new Error('このユーザー名は既に使われています');
  }
  var salt = generateSalt_();
  sheet.appendRow([p.store, p.username, hashPassword_(p.password, salt), salt, p.role]);
  return { username: p.username, store: p.store, role: p.role };
}

function resetPassword_(session, p) {
  requireRole_(session, ['hq']);
  var sheet = getSheet_(SHEET_ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === p.username) {
      var salt = generateSalt_();
      sheet.getRange(i + 1, 3).setValue(hashPassword_(p.newPassword, salt));
      sheet.getRange(i + 1, 4).setValue(salt);
      return { username: p.username };
    }
  }
  throw new Error('ユーザーが見つかりません');
}

/**
 * ログイン中の本人のパスワードを再検証する。在庫の一括/個別リセットのような
 * 取り消せない操作の直前に、セッションだけでなくパスワードでも再確認するために使う。
 */
function verifyOwnPassword_(session, password) {
  if (!password) throw new Error('パスワードを入力してください');
  var sheet = getSheet_(SHEET_ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === session.username) {
      if (hashPassword_(password, data[i][3]) !== data[i][2]) {
        throw new Error('パスワードが違います');
      }
      return true;
    }
  }
  throw new Error('アカウントが見つかりません');
}
