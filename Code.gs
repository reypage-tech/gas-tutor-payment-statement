/**
 * カスタムメニュー「月次処理」とサイドバー（HTML Service）の入口。
 * メニュー項目は常に設定シートの「対象年月」に対して実行する。複数の月を試したい場合は
 * サイドバーの対象月選択を使う（サイドバーは google.script.run で yyyyMM を明示的に渡す）。
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('月次処理')
    .addItem('①チェック', 'menuRunCheck')
    .addItem('②計算', 'menuRunCalculate')
    .addItem('③PDF作成', 'menuRunPdf')
    .addItem('④下書き作成', 'menuRunDraft')
    .addItem('⑤確定（管理者のみ）', 'menuRunFinalize')
    .addSeparator()
    .addItem('サイドバーを開く', 'showSidebar')
    .addItem('ダッシュボード更新', 'menuRefreshDashboard')
    .addSeparator()
    .addItem('初期設定（初回のみ）', 'setupInitial')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('月次処理').setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function alertResult_(label, fn) {
  try {
    var state = fn();
    SpreadsheetApp.getUi().alert(label + '\n' + JSON.stringify(state));
  } catch (err) {
    SpreadsheetApp.getUi().alert(label + ' でエラー:\n' + err.message);
  }
}

function menuRunCheck() { alertResult_('①チェック完了', function () { return runCheck(); }); }
function menuRunCalculate() { alertResult_('②計算完了', function () { return runCalculate(); }); }
function menuRunPdf() { alertResult_('③PDF作成完了', function () { return runPdf(); }); }
function menuRunDraft() { alertResult_('④下書き作成完了', function () { return runDraft(); }); }
function menuRunFinalize() { alertResult_('⑤確定完了', function () { return runFinalize(); }); }
function menuRefreshDashboard() { refreshDashboard_(); SpreadsheetApp.getActive().toast('ダッシュボードを更新しました'); }

/** 実績_yyyyMM シート名から、選択可能な対象月の一覧を作る */
function listAvailableMonths_() {
  return getActiveSs_().getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n.indexOf(SHEET_PREFIX_PERFORMANCE) === 0; })
    .map(function (n) { return n.slice(SHEET_PREFIX_PERFORMANCE.length); })
    .filter(function (m) { return /^\d{6}$/.test(m); })
    .sort();
}

/** どのボタンを押せるか（順番外のボタンは無効にする） */
function computeStepAvailability_(state) {
  var noErrors = state.errorCount === 0;
  return {
    CHECK: true,
    CALC: !!state.checkedAt && noErrors,
    PDF: !!state.calculatedAt && noErrors,
    DRAFT: !!state.pdfAt && noErrors,
    FINALIZE: !!state.draftAt && noErrors && !state.needsRedraft
  };
}

/** サイドバーの初期表示・各操作後の再表示に使う */
function sidebarGetState(yyyyMM) {
  var settings = getSettingsWithDefaults_();
  var month = yyyyMM || settings.targetMonth;
  var state = getMonthState_(month);
  return {
    month: month,
    label: formatMonthLabel_(month),
    availableMonths: listAvailableMonths_(),
    isAdmin: isAdmin_(settings),
    isFinalized: isMonthFinalized_(month),
    availability: computeStepAvailability_(state),
    state: state
  };
}

/** サイドバーの各ボタンから呼ばれる */
function sidebarRunStep(step, yyyyMM) {
  var entrypoints = { CHECK: runCheck, CALC: runCalculate, PDF: runPdf, DRAFT: runDraft, FINALIZE: runFinalize };
  var fn = entrypoints[step];
  if (!fn) return { ok: false, message: '不明なステップです: ' + step };
  try {
    var state = fn(yyyyMM);
    return { ok: true, sidebar: sidebarGetState(yyyyMM) };
  } catch (err) {
    return { ok: false, message: err.message, sidebar: sidebarGetState(yyyyMM) };
  }
}
