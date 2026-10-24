/**
 * 汎用ユーティリティ（数値変換・日付・ロック・管理者判定など）
 */

/** "5,000円" や "５０００" のような文字でも数値にする。読めなければ0 */
function toNumber_(x) {
  if (x === '' || x === null || x === undefined) return 0;
  var half = String(x).replace(/[０-９．－]/g, function (c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
  var n = Number(half.replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

/** 空欄・非数値かどうか（単価欠落チェック用。0円は「設定済みの単価」として区別する） */
function isBlank_(x) {
  return x === '' || x === null || x === undefined;
}

/** 12345、"12,345円" */
function yen_(n) {
  return Math.round(n).toLocaleString('ja-JP') + '円';
}

/** 設定シートの端数処理方式で四捨五入・切上げ・切下げを行う */
function roundAmount_(amount, method) {
  if (method === '切上げ') return Math.ceil(amount);
  if (method === '切下げ') return Math.floor(amount);
  return Math.round(amount);
}

/** 'yyyy-MM' / 'yyyyMM' のどちらでも 'yyyyMM' に正規化する。
 * スプレッドシート側で日付として自動変換された場合（Dateオブジェクトになっている場合）も扱う。 */
function normalizeMonth_(monthValue) {
  if (monthValue instanceof Date) {
    return Utilities.formatDate(monthValue, 'Asia/Tokyo', 'yyyyMM');
  }
  var digits = String(monthValue || '').replace(/[^\d]/g, '');
  if (digits.length !== 6) throw new Error('対象年月の形式が不正です: ' + monthValue);
  return digits;
}

/** 'yyyyMM'、表示用 'yyyy年M月' */
function formatMonthLabel_(yyyyMM) {
  var y = yyyyMM.slice(0, 4);
  var m = Number(yyyyMM.slice(4, 6));
  return y + '年' + m + '月';
}

/** 'yyyyMM' の前月を 'yyyyMM' で返す */
function previousMonth_(yyyyMM) {
  var y = Number(yyyyMM.slice(0, 4));
  var m = Number(yyyyMM.slice(4, 6));
  m -= 1;
  if (m === 0) { m = 12; y -= 1; }
  return y + ('' + m).padStart(2, '0');
}

/** 今日の日付文字列（設定のタイムゾーン基準） */
function todayStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function nowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

/** 現在時刻(ms)。6分制限の判定に使う。test/run.jsではこの関数だけを差し替えて、
 * 実際に待たずに「時間切れ、中断、再開」を再現する。 */
function nowMs_() {
  return new Date().getTime();
}

/**
 * スクリプトロックを取って fn を実行する。月次処理は同時に1人しか実行できないようにし、
 * 計算結果・合計シートが同時書き込みで壊れることを防ぐ。
 */
function withScriptLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT_MS);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function parseEmails_(commaSeparated) {
  return String(commaSeparated || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

/** settings（getSettingsWithDefaults_の戻り値）の管理者メールアドレスに実行者が含まれるか */
function isAdmin_(settings) {
  var email = Session.getEffectiveUser().getEmail();
  var admins = parseEmails_(settings.adminEmails);
  return admins.indexOf(email) !== -1;
}

/** オブジェクトの浅い内容から安定した短いハッシュ文字列を作る（再計算での金額変化検知用） */
function simpleHash_(value) {
  var str = JSON.stringify(value);
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash + str.charCodeAt(i)) | 0;
  }
  return String(hash);
}
