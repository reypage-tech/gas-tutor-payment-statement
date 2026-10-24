/**
 * 対象月ごとの進行状態（PropertiesService）。サイドバーのボタン有効・無効
 * （「順番外のボタンは無効」）はこの状態だけを見て決める。
 */

function saveMonthState_(yyyyMM, patch) {
  var props = PropertiesService.getScriptProperties();
  var current = getMonthState_(yyyyMM);
  var next = Object.assign(current, patch);
  props.setProperty(PROP_STATE_PREFIX + yyyyMM, JSON.stringify(next));
  return next;
}

function getMonthState_(yyyyMM) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_STATE_PREFIX + yyyyMM);
  var base = {
    month: yyyyMM,
    step: MONTH_STATE.NONE,
    errorCount: 0,
    warningCount: 0,
    checkedAt: '',
    calculatedAt: '',
    pdfAt: '',
    draftAt: '',
    finalizedAt: '',
    finalizedBy: '',
    needsRedraft: false
  };
  if (!raw) return base;
  return Object.assign(base, JSON.parse(raw));
}

function isMonthFinalized_(yyyyMM) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_FINALIZED_MONTHS);
  var list = raw ? JSON.parse(raw) : [];
  return list.indexOf(yyyyMM) !== -1;
}

function markMonthFinalized_(yyyyMM) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_FINALIZED_MONTHS);
  var list = raw ? JSON.parse(raw) : [];
  if (list.indexOf(yyyyMM) === -1) list.push(yyyyMM);
  props.setProperty(PROP_FINALIZED_MONTHS, JSON.stringify(list));
}

function unmarkMonthFinalized_(yyyyMM) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_FINALIZED_MONTHS);
  var list = raw ? JSON.parse(raw) : [];
  list = list.filter(function (m) { return m !== yyyyMM; });
  props.setProperty(PROP_FINALIZED_MONTHS, JSON.stringify(list));
}

// ===== 冪等性の記録（PDF・下書きのID）。同じ月を再実行しても二重作成しない =====

function getIdempotencyMap_(prefix, yyyyMM) {
  var raw = PropertiesService.getScriptProperties().getProperty(prefix + yyyyMM);
  return raw ? JSON.parse(raw) : {};
}

function saveIdempotencyMap_(prefix, yyyyMM, map) {
  PropertiesService.getScriptProperties().setProperty(prefix + yyyyMM, JSON.stringify(map));
}

// ===== 6分制限対策：バッチの進捗保存 =====

function getBatchProgress_(yyyyMM, step) {
  var raw = PropertiesService.getScriptProperties().getProperty(PROP_PROGRESS_PREFIX + yyyyMM + '_' + step);
  return raw ? JSON.parse(raw) : null;
}

function saveBatchProgress_(yyyyMM, step, progress) {
  PropertiesService.getScriptProperties().setProperty(PROP_PROGRESS_PREFIX + yyyyMM + '_' + step, JSON.stringify(progress));
}

function clearBatchProgress_(yyyyMM, step) {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROGRESS_PREFIX + yyyyMM + '_' + step);
}

// ===== 割当上限（日次）の手前で止める仕組み =====
// PDF作成・下書き作成それぞれの「その日の作成件数」を日付ごとに数える。
// 日付が変わればキーが変わるため、翌日になれば自然にカウントが0から再開する。

function dailyCountKey_(kind) {
  return PROP_DAILY_COUNT_PREFIX + kind + '_' + todayStr_();
}

function getDailyCount_(kind) {
  var raw = PropertiesService.getScriptProperties().getProperty(dailyCountKey_(kind));
  return raw ? Number(raw) : 0;
}

function incrementDailyCount_(kind) {
  var next = getDailyCount_(kind) + 1;
  PropertiesService.getScriptProperties().setProperty(dailyCountKey_(kind), String(next));
  return next;
}

var DAILY_LIMIT_LABEL_ = { PDF: 'PDF', DRAFT: 'Gmail下書き' };

/** 日次の上限に達していたら、処理を止めるための専用エラーを投げる（Batch.gsが中断として扱う） */
function assertDailyLimitNotReached_(kind, limit) {
  if (getDailyCount_(kind) >= limit) {
    var err = new Error(
      '日次上限（' + limit + '件）に達したため、本日はこれ以上' + DAILY_LIMIT_LABEL_[kind] +
      'の作成を行いません。進捗は保存済みです。翌日、同じステップを再実行すれば続きから再開します。'
    );
    err.isDailyLimitReached = true;
    throw err;
  }
}
