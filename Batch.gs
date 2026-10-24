/**
 * 6分制限対策：講師単位のバッチ処理＋続きから再開。
 * 進捗（次に処理する講師のインデックス）をPropertiesServiceに保存し、制限時間に近づいたら
 * 時間主導トリガー（resumeBatch）を1回だけ登録して処理を打ち切る。トリガーは
 * 「このスクリプトが作ったものだけ」を対象に削除・再登録する。
 * 例外は講師単位で隔離し、失敗した講師だけ「要再実行」にする（他の講師の処理は止めない）。
 * 日次の作成件数が設定の上限に達した場合（assertDailyLimitNotReached_が投げる
 * isDailyLimitReachedエラー）は、他の講師の処理は続けず、その場で中断して進捗を保存する
 * （翌日、同じ月のステップを再実行すれば保存した位置から再開できる）。
 */

/**
 * @param {string} yyyyMM
 * @param {string} step 'CALC' | 'PDF' | 'DRAFT'
 * @param {string[]} instructorIds 処理対象の講師ID一覧（呼び出し側で順序を固定しておくこと）
 * @param {function(string):void} perInstructorFn 1講師分の処理。例外はここで投げてよい（呼び出し側で隔離する）
 * @returns {{done: boolean, processedCount: number, failedIds: string[]}}
 */
function runBatchOverInstructors_(yyyyMM, step, instructorIds, perInstructorFn) {
  var progress = getBatchProgress_(yyyyMM, step) || { nextIndex: 0, failedIds: [] };
  var startedAt = nowMs_();

  for (var i = progress.nextIndex; i < instructorIds.length; i++) {
    if (nowMs_() - startedAt > BATCH_TIME_LIMIT_MS) {
      progress.nextIndex = i;
      saveBatchProgress_(yyyyMM, step, progress);
      scheduleContinuation_(yyyyMM, step);
      return { done: false, processedCount: i, failedIds: progress.failedIds };
    }
    var instructorId = instructorIds[i];
    try {
      perInstructorFn(instructorId);
    } catch (err) {
      if (err && err.isDailyLimitReached) {
        progress.nextIndex = i;
        saveBatchProgress_(yyyyMM, step, progress);
        appendErrorLog_(step, '-', yyyyMM, err.message);
        return { done: false, processedCount: i, failedIds: progress.failedIds, pausedForDailyLimit: true };
      }
      appendErrorLog_(step, instructorId, yyyyMM, err && err.message ? err.message : err);
      progress.failedIds.push(instructorId);
    }
    progress.nextIndex = i + 1;
    saveBatchProgress_(yyyyMM, step, progress);
  }

  clearBatchProgress_(yyyyMM, step);
  cleanupOwnTriggers_();
  return { done: true, processedCount: instructorIds.length, failedIds: progress.failedIds };
}

var RESUME_PENDING_PROP = 'RESUME_PENDING';

/** このスクリプトが作ったトリガーだけを対象に削除する（他の処理のトリガーには影響しない） */
function cleanupOwnTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === CONTINUATION_TRIGGER_HANDLER; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

function scheduleContinuation_(yyyyMM, step) {
  cleanupOwnTriggers_();
  PropertiesService.getScriptProperties().setProperty(RESUME_PENDING_PROP, JSON.stringify({ month: yyyyMM, step: step }));
  ScriptApp.newTrigger(CONTINUATION_TRIGGER_HANDLER).timeBased().after(10 * 1000).create();
}

var STEP_RESUME_ENTRYPOINTS_ = {
  CALC: function (yyyyMM) { return runCalculate(yyyyMM); },
  PDF: function (yyyyMM) { return runPdf(yyyyMM); },
  DRAFT: function (yyyyMM) { return runDraft(yyyyMM); }
};

/**
 * 時間主導トリガーから呼ばれる。中断していたバッチの続きを実行する。
 * 自分（このトリガー）だけを削除してから再開する。
 */
function resumeBatch() {
  cleanupOwnTriggers_();
  var raw = PropertiesService.getScriptProperties().getProperty(RESUME_PENDING_PROP);
  if (!raw) return;
  PropertiesService.getScriptProperties().deleteProperty(RESUME_PENDING_PROP);
  var pending = JSON.parse(raw);
  var fn = STEP_RESUME_ENTRYPOINTS_[pending.step];
  if (fn) fn(pending.month);
}

/**
 * 初回セットアップ（メニュー「初期設定」から手動実行）。
 * 前回の異常終了等で残った再開用トリガーを掃除し、PDF用フォルダを準備する。
 * 何度実行しても結果は同じ（冪等）。
 */
function setupInitial() {
  cleanupOwnTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(RESUME_PENDING_PROP);
  getOrCreatePdfRootFolderId_();
  SpreadsheetApp.getActive().toast('初期設定が完了しました。', '月次処理', 5);
}
