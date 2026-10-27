// node上でのGAS実行環境モック。test/gas_mock/mock_apis.js の上で全 .gs ファイルをvmコンテキストに
// 読み込み、正常系・異常系・大量件数・冪等性・6分制限からの再開・送信系APIが呼ばれないことを確認する。
// 実際のGoogle Apps Script実行環境では実行していない。
var fs = require('fs');
var path = require('path');
var vm = require('vm');
var mockApis = require('./gas_mock/mock_apis.js');

var ROOT = path.join(__dirname, '..');
var DATA_DIR = path.join(ROOT, 'data');

var GAS_FILES = [
  'Config.gs', 'Util.gs', 'Sheets.gs', 'State.gs', 'Drive.gs',
  'Batch.gs', 'Validate.gs', 'Calculate.gs', 'Pdf.gs', 'Draft.gs', 'Finalize.gs',
  'Dashboard.gs', 'Code.gs'
];

var results = [];
function logResult(name, ok, detail) {
  results.push({ name: name, ok: ok, detail: detail || '' });
  console.log((ok ? '[OK] ' : '[NG] ') + name + (detail ? ' - ' + detail : ''));
}
function assert(name, cond, detail) {
  logResult(name, !!cond, detail);
}

// ===== CSVの読み込み（data/配下の見本データをそのままテストに使う） =====

function parseCsv(relPath) {
  var text = fs.readFileSync(path.join(DATA_DIR, relPath), 'utf8');
  return text.split('\n')
    .map(function (l) { return l.replace(/\r$/, ''); })
    .filter(function (l) { return l.length > 0; })
    .map(function (l) { return l.split(','); });
}

function maybeNumber(s) {
  if (s === '' || s === undefined) return '';
  var n = Number(s);
  return isNaN(n) ? s : n;
}

function loadSheetFromRows(ss, sheetName, rows, numericCols) {
  var sheet = ss.insertSheet(sheetName);
  var converted = rows.map(function (row, rIdx) {
    return row.map(function (cell, cIdx) {
      if (rIdx === 0) return cell;
      return (numericCols && numericCols.indexOf(cIdx) !== -1) ? maybeNumber(cell) : cell;
    });
  });
  sheet.getRange(1, 1, converted.length, converted[0].length).setValues(converted);
  return sheet;
}

// ===== vmコンテキストの構築 =====

function loadContext() {
  var built = mockApis.buildMockGlobals();
  var sandbox = built.globals;
  var context = vm.createContext(sandbox);
  GAS_FILES.forEach(function (fileName) {
    var code = fs.readFileSync(path.join(ROOT, fileName), 'utf8');
    vm.runInContext(code, context, { filename: fileName });
  });
  return { context: context, world: built.world };
}

/** 標準のマスタ・設定・実績（data/配下のCSVそのまま）を積んだ新規コンテキストを返す */
function freshContextWithFixtures() {
  var loaded = loadContext();
  var ss = loaded.world.spreadsheet;
  loadSheetFromRows(ss, '設定', parseCsv('settings.csv'), []);
  loadSheetFromRows(ss, 'マスタ_講師', parseCsv('master_instructors.csv'), []);
  loadSheetFromRows(ss, 'マスタ_品目', parseCsv('master_items.csv'), [3, 4, 5]);
  loadSheetFromRows(ss, '実績_202608', parseCsv('performance_202608.csv'), [3]);
  loadSheetFromRows(ss, '実績_202609', parseCsv('performance_202609.csv'), [3]);
  return loaded;
}

function setSettingValue_(loaded, key, value) {
  var sheet = loaded.world.spreadsheet.getSheetByName('設定');
  var rows = sheet.__dumpRows();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  throw new Error('設定項目が見つかりません: ' + key);
}

function countPdfFiles_(loaded) {
  return Object.keys(loaded.world.files).filter(function (id) { return loaded.world.files[id].mimeType === 'application/pdf' && !loaded.world.files[id].trashed; }).length;
}

function countByColumn_(sheet, colIndex) {
  var rows = sheet.__dumpRows().slice(1);
  var counts = {};
  rows.forEach(function (r) {
    var key = String(r[colIndex]);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

// =====================================================================
// T01: 正常系（202608・異常データなしの月で全ステップを通す）
// =====================================================================
(function T01_normal() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');

  var checkState = c.runCheck();
  assert('T01: チェックでエラー0件（異常データなしの月）', checkState.errorCount === 0, JSON.stringify(checkState));

  var calcState = c.runCalculate();
  assert('T01: 計算が完了しCALCULATEDになる', calcState.step === 'CALCULATED', JSON.stringify(calcState));

  var calcSheet = loaded.world.spreadsheet.getSheetByName('計算結果_202608');
  var calcRows = calcSheet.__dumpRows().slice(1);
  assert('T01: 有効な講師19名分の計算結果行がある', calcRows.length === 19, '行数=' + calcRows.length);

  var pdfState = c.runPdf();
  assert('T01: PDF作成が完了しPDF_DONEになる', pdfState.step === 'PDF_DONE', JSON.stringify(pdfState));
  var pdfFileCount = countPdfFiles_(loaded);
  assert('T01: PDFファイルが19件作成される', pdfFileCount === 19, '件数=' + pdfFileCount);

  var draftState = c.runDraft();
  assert('T01: 下書き作成が完了しDRAFTEDになる', draftState.step === 'DRAFTED', JSON.stringify(draftState));
  var draftCount = Object.keys(loaded.world.drafts).length;
  assert('T01: 下書きがメール設定済みの18名分作成される（T18はメール未設定でスキップ）', draftCount === 18, '件数=' + draftCount);

  var errorLog = loaded.world.spreadsheet.getSheetByName('エラーログ');
  var hasSkipLog = errorLog && errorLog.__dumpRows().some(function (r) { return String(r[4] || '').indexOf('メールアドレス未設定') !== -1; });
  assert('T01: メール未設定講師がエラーログに記録される', !!hasSkipLog);

  loaded.world.currentUserEmail = 'admin@example.com';
  var finalState = c.runFinalize();
  assert('T01: 確定（管理者）が完了しFINALIZEDになる', finalState.step === 'FINALIZED', JSON.stringify(finalState));
  assert('T01: 確定後、対象月が保護済みと判定される', c.isMonthFinalized_('202608') === true);

  var auditLog = loaded.world.spreadsheet.getSheetByName('監査ログ');
  assert('T01: 監査ログに5ステップ分が記録される', auditLog.__dumpRows().length - 1 === 5, '行数=' + (auditLog.__dumpRows().length - 1));
})();

// =====================================================================
// T02: 欠落（未登録品目・単価欠落・未登録講師・無効講師）― computeChecks_の純粋関数テスト
// =====================================================================
(function T02_missing() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructorMap = {
    T01: { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' },
    T19: { id: 'T19', name: '講師19', rank: 'C', email: 't19@example.com', status: '無効' }
  };
  var itemMap = {
    I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 3200, B: 2800, C: 2400 } },
    I30: { id: 'I30', name: '研修参加', unit: '回', rates: { A: '', B: '', C: 500 } }
  };
  var rows = [
    ['2026-09-01', 'T01', 'I30', 1, ''],   // 単価欠落（ランクAのI30単価が未設定）
    ['2026-09-02', 'T99', 'I01', 3, ''],   // 未登録講師
    ['2026-09-03', 'T01', 'I99', 2, ''],   // 未登録品目
    ['2026-09-04', 'T19', 'I01', 1, '']    // 無効講師
  ];
  var settings = { prevMonthRatioHigh: 2.0, prevMonthRatioLow: 0.3 };
  var result = c.computeChecks_(rows, instructorMap, itemMap, [], settings);
  var types = result.issues.map(function (i) { return i.type; });
  assert('T02: 単価欠落が検出される', types.indexOf('単価欠落') !== -1, JSON.stringify(types));
  assert('T02: 未登録講師が検出される', types.indexOf('未登録講師') !== -1, JSON.stringify(types));
  assert('T02: 未登録品目が検出される', types.indexOf('未登録品目') !== -1, JSON.stringify(types));
  assert('T02: 無効講師が検出される', types.indexOf('無効講師') !== -1, JSON.stringify(types));
  assert('T02: すべて深刻度エラーになる', result.issues.every(function (i) { return i.severity === 'エラー'; }), JSON.stringify(result.issues));
})();

// =====================================================================
// T03: 重複行
// =====================================================================
(function T03_duplicate() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructorMap = { T01: { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' } };
  var itemMap = { I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 3200, B: 2800, C: 2400 } } };
  var rows = [
    ['2026-09-01', 'T01', 'I01', 3, ''],
    ['2026-09-01', 'T01', 'I01', 7, ''] // 同日・同講師・同品目の重複（件数が違っても重複として検出する）
  ];
  var result = c.computeChecks_(rows, instructorMap, itemMap, [], { prevMonthRatioHigh: 2.0, prevMonthRatioLow: 0.3 });
  var dupIssues = result.issues.filter(function (i) { return i.type === '重複行'; });
  assert('T03: 重複行が2件とも検出される', dupIssues.length === 2, JSON.stringify(dupIssues));
})();

// =====================================================================
// T04: 負数
// =====================================================================
(function T04_negative() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructorMap = { T01: { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' } };
  var itemMap = { I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 3200, B: 2800, C: 2400 } } };
  var rows = [
    ['2026-09-01', 'T01', 'I01', -5, ''],
    ['2026-09-02', 'T01', 'I01', 'abc', ''] // 非数値（数量不正）も併せて確認
  ];
  var result = c.computeChecks_(rows, instructorMap, itemMap, [], { prevMonthRatioHigh: 2.0, prevMonthRatioLow: 0.3 });
  var types = result.issues.map(function (i) { return i.type; });
  assert('T04: 負数が検出される', types.indexOf('負数') !== -1, JSON.stringify(types));
  assert('T04: 非数値（数量不正）が検出される', types.indexOf('数量不正') !== -1, JSON.stringify(types));
})();

// =====================================================================
// T05: 端数処理・税区分
// =====================================================================
(function T05_rounding() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructor = { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' };
  var itemMap = { I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 333, B: 0, C: 0 } } };
  var rows = [['2026-09-01', 'T01', 'I01', 1, '']]; // 小計333円、税10%なら36.3円

  var roundUp = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '課税', taxRate: 0.1, roundingMethod: '切上げ' });
  var roundDown = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '課税', taxRate: 0.1, roundingMethod: '切下げ' });
  var roundNearest = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '課税', taxRate: 0.1, roundingMethod: '四捨五入' });
  var noTax = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '非課税', taxRate: 0.1, roundingMethod: '四捨五入' });

  assert('T05: 切上げで367円になる(333+33.3=366.3を切上げ)', roundUp.total === 367, JSON.stringify(roundUp));
  assert('T05: 切下げで366円になる(366.3を切下げ)', roundDown.total === 366, JSON.stringify(roundDown));
  assert('T05: 四捨五入で366円になる', roundNearest.total === 366, JSON.stringify(roundNearest));
  assert('T05: 非課税では税が0円で小計のまま333円になる', noTax.total === 333 && noTax.tax === 0, JSON.stringify(noTax));
})();

// =====================================================================
// T06: 前月比異常
// =====================================================================
(function T06_prevMonthAnomaly() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructorMap = { T01: { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' } };
  var itemMap = { I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 3200, B: 2800, C: 2400 } } };
  var prevRows = [['2026-08-01', 'T01', 'I01', 10, '']];
  var curRowsHigh = [['2026-09-01', 'T01', 'I01', 30, '']]; // 3倍 -> 閾値2.0以上で異常
  var curRowsLow = [['2026-09-01', 'T01', 'I01', 1, '']]; // 0.1倍 -> 閾値0.3以下で異常
  var settings = { prevMonthRatioHigh: 2.0, prevMonthRatioLow: 0.3 };

  var resultHigh = c.computeChecks_(curRowsHigh, instructorMap, itemMap, prevRows, settings);
  var resultLow = c.computeChecks_(curRowsLow, instructorMap, itemMap, prevRows, settings);
  assert('T06: 前月比が高すぎる場合に異常として警告される', resultHigh.issues.some(function (i) { return i.type === '前月比異常' && i.severity === '注意'; }), JSON.stringify(resultHigh.issues));
  assert('T06: 前月比が低すぎる場合に異常として警告される', resultLow.issues.some(function (i) { return i.type === '前月比異常'; }), JSON.stringify(resultLow.issues));
})();

// =====================================================================
// T07: 大量件数（200名）
// =====================================================================
(function T07_largeVolume() {
  var loaded = loadContext();
  var c = loaded.context;
  var ss = loaded.world.spreadsheet;

  var instructorRows = [['講師ID', '表示名', 'ランク', 'メールアドレス', '契約状態']];
  var perfRows = [['日付', '講師ID', '品目ID', '件数', 'メモ']];
  for (var i = 1; i <= 200; i++) {
    var id = 'L' + ('000' + i).slice(-3);
    var rank = ['A', 'B', 'C'][i % 3];
    instructorRows.push([id, '大量講師' + i, rank, 'bulk' + i + '@example.com', '有効']);
    perfRows.push(['2026-09-10', id, 'I01', (i % 5) + 1, '']);
  }
  var itemRows = [['品目ID', '品目名', '単位', '単価A', '単価B', '単価C'], ['I01', '個別指導60分', '回', 3200, 2800, 2400]];
  var settingsRows = parseCsv('settings.csv');

  loadSheetFromRows(ss, '設定', settingsRows, []);
  setSettingValue_(loaded, '対象年月', '2026-09');
  loadSheetFromRows(ss, 'マスタ_講師', instructorRows, []);
  loadSheetFromRows(ss, 'マスタ_品目', itemRows, [3, 4, 5]);
  loadSheetFromRows(ss, '実績_202609', perfRows, [3]);

  var checkState = c.runCheck();
  assert('T07: 200名分でもチェックがエラー0件で完了する', checkState.errorCount === 0, JSON.stringify(checkState));
  var calcState = c.runCalculate();
  assert('T07: 200名分の計算がCALCULATEDになる', calcState.step === 'CALCULATED', JSON.stringify(calcState));
  var calcRows = ss.getSheetByName('計算結果_202609').__dumpRows().slice(1);
  assert('T07: 計算結果が200行になる', calcRows.length === 200, '行数=' + calcRows.length);
})();

// =====================================================================
// T08: 再実行で二重作成しない（冪等性：PDF・下書き）
// =====================================================================
(function T08_idempotent() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');

  c.runCheck();
  c.runCalculate();
  c.runPdf();
  c.runDraft();

  var pdfCountAfterFirst = countPdfFiles_(loaded);
  var draftCountAfterFirst = Object.keys(loaded.world.drafts).length;
  var pdfMapAfterFirst = JSON.parse(JSON.stringify(c.getIdempotencyMap_('PDF_IDS_', '202608')));
  var draftMapAfterFirst = JSON.parse(JSON.stringify(c.getIdempotencyMap_('DRAFT_IDS_', '202608')));

  // 金額を何も変えずに再実行する（人がもう一度メニューを押した想定）
  c.runPdf();
  c.runDraft();

  var pdfCountAfterSecond = countPdfFiles_(loaded);
  var draftCountAfterSecond = Object.keys(loaded.world.drafts).length;
  var pdfMapAfterSecond = c.getIdempotencyMap_('PDF_IDS_', '202608');
  var draftMapAfterSecond = c.getIdempotencyMap_('DRAFT_IDS_', '202608');

  assert('T08: 再実行してもPDFファイル数が増えない', pdfCountAfterSecond === pdfCountAfterFirst, pdfCountAfterFirst + ' -> ' + pdfCountAfterSecond);
  assert('T08: 再実行しても下書き数が増えない', draftCountAfterSecond === draftCountAfterFirst, draftCountAfterFirst + ' -> ' + draftCountAfterSecond);
  assert('T08: PDFのファイルIDが再実行後も同じ(T01)', pdfMapAfterSecond.T01.fileId === pdfMapAfterFirst.T01.fileId);
  assert('T08: 下書きのIDが再実行後も同じ(T01)', draftMapAfterSecond.T01.draftId === draftMapAfterFirst.T01.draftId);

  // 再計算で金額が変わった場合は、PDF・下書きが「置換」されIDが変わることを確認する
  // （計算結果シートの金額を直接書き換えて「再計算で金額が変わった」状態を作る）
  var calcSheet = loaded.world.spreadsheet.getSheetByName('計算結果_202608');
  var calcRows = calcSheet.__dumpRows();
  calcRows[1][6] = calcRows[1][6] + 99999; // 講師T01(1行目)の合計金額を変更
  calcSheet.getRange(2, 7).setValue(calcRows[1][6]);
  c.finalizeCalculation_('202608', []);

  var stateAfterChange = c.getMonthState_('202608');
  assert('T08: 金額変更後はneedsRedraftがtrueになる', stateAfterChange.needsRedraft === true, JSON.stringify(stateAfterChange));

  c.runPdf();
  var pdfMapAfterReplace = c.getIdempotencyMap_('PDF_IDS_', '202608');
  assert('T08: 金額変更後はPDFのファイルIDが置換される(T01)', pdfMapAfterReplace.T01.fileId !== pdfMapAfterFirst.T01.fileId);

  c.runDraft();
  var draftMapAfterReplace = c.getIdempotencyMap_('DRAFT_IDS_', '202608');
  assert('T08: 金額変更後は下書きのIDが置換される(T01)', draftMapAfterReplace.T01.draftId !== draftMapAfterFirst.T01.draftId);
  assert('T08: 下書き再作成後はneedsRedraftがfalseに戻る', c.getMonthState_('202608').needsRedraft === false);
})();

// =====================================================================
// T09: 6分制限からの再開
// =====================================================================
(function T09_resumeBatch() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();

  // 実時間を待たずに「6分制限に近づいた」を再現する：呼ぶたびに時計が1000ms進む偽の時計に
  // 差し替え、制限を1500msにすることで「1人処理するたびに時間切れで中断する」状態を作る。
  var fakeNow = 0;
  c.nowMs_ = function () { fakeNow += 1000; return fakeNow; };
  c.BATCH_TIME_LIMIT_MS = 1500;

  var partialState = c.runCalculate();
  assert('T09: 時間切れで計算が中断し、CALCULATEDにならない', partialState.step !== 'CALCULATED', JSON.stringify(partialState));
  var ownTriggers = loaded.world.triggers.filter(function (t) { return t.getHandlerFunction() === 'resumeBatch'; });
  assert('T09: 再開用トリガー(resumeBatch)が1件登録される', ownTriggers.length === 1, '件数=' + ownTriggers.length);

  var progress = JSON.parse(loaded.world.scriptProperties['PROGRESS_202608_CALC']);
  assert('T09: 進捗が1件目を処理した時点で保存されている', progress.nextIndex === 1, 'nextIndex=' + progress.nextIndex);

  var resumeCount = 0;
  while (loaded.world.triggers.some(function (t) { return t.getHandlerFunction() === 'resumeBatch'; }) && resumeCount < 25) {
    c.resumeBatch();
    resumeCount += 1;
  }
  var finalState = c.getMonthState_('202608');
  assert('T09: 再開を繰り返すと最終的にCALCULATEDになる', finalState.step === 'CALCULATED', JSON.stringify(finalState) + ' resumeCount=' + resumeCount);

  var calcSheet = loaded.world.spreadsheet.getSheetByName('計算結果_202608');
  var idCounts = countByColumn_(calcSheet, 0);
  var duplicated = Object.keys(idCounts).filter(function (id) { return idCounts[id] > 1; });
  assert('T09: 中断・再開を経ても計算結果に重複行ができない', duplicated.length === 0, JSON.stringify(idCounts));
  assert('T09: 計算結果が19行になる(欠落・重複なし)', Object.keys(idCounts).length === 19, '件数=' + Object.keys(idCounts).length);
})();

// =====================================================================
// T10: 確定済み月の保護（管理者のみ再計算可）
// =====================================================================
(function T10_finalizedProtection() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  c.runPdf();
  c.runDraft();

  loaded.world.currentUserEmail = 'staff@example.com';
  var threwForStaff = false;
  try { c.runFinalize(); } catch (err) { threwForStaff = /管理者のみ/.test(err.message); }
  assert('T10: 担当者（管理者以外）は確定できない', threwForStaff);

  loaded.world.currentUserEmail = 'admin@example.com';
  c.runFinalize();
  assert('T10: 確定後は対象月がFINALIZED済みと判定される', c.isMonthFinalized_('202608') === true);

  loaded.world.currentUserEmail = 'staff@example.com';
  var threwForRecalc = false;
  try { c.runCalculate(); } catch (err) { threwForRecalc = /確定済み/.test(err.message); }
  assert('T10: 確定済み月は担当者が再計算できない', threwForRecalc);

  loaded.world.currentUserEmail = 'admin@example.com';
  var unlockedOk = true;
  try { c.runCalculate(); } catch (err) { unlockedOk = false; }
  assert('T10: 確定済み月でも管理者なら保護を解除して再計算できる', unlockedOk);
  assert('T10: 管理者の再計算後はFINALIZED状態が解除される', c.isMonthFinalized_('202608') === false);
})();

// =====================================================================
// T11: 送信系APIが呼ばれないこと
// =====================================================================
(function T11_noSendFunctionCalled() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  assert('T11: GmailAppはモックに存在しない(呼べば即ReferenceErrorになる構成)', typeof c.GmailApp === 'undefined');
  assert('T11: DriveAppはモックに存在しない(呼べば即ReferenceErrorになる構成)', typeof c.DriveApp === 'undefined');
  assert('T11: MailAppはモックに存在しない(呼べば即ReferenceErrorになる構成)', typeof c.MailApp === 'undefined');
  assert('T11: UrlFetchAppはモックに存在しない(呼べば即ReferenceErrorになる構成)', typeof c.UrlFetchApp === 'undefined');

  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  c.runPdf();
  c.runDraft();
  assert('T11: GmailApp・DriveApp等が未実装のままチェック〜下書き作成までのフルフローが例外なく完走する（呼んでいれば必ずここで例外になる）', true);
})();

// =====================================================================
// T12: メールアドレス未設定講師のスキップ（T01で確認済みだが単体でも確認する）
// =====================================================================
(function T12_missingEmailSkip() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  c.runPdf();
  c.runDraft();
  var draftMap = c.getIdempotencyMap_('DRAFT_IDS_', '202608');
  assert('T12: メール未設定のT18には下書きが作られない', !draftMap.T18);
  var calcMap = c.readCalcResultsMap_('202608');
  assert('T12: T18自体は計算結果には含まれる(計算自体はできる)', !!calcMap.T18);
})();

// =====================================================================
// T13: サイドバーのボタン有効/無効（順番外は無効）
// =====================================================================
(function T13_sidebarAvailability() {
  var loaded = loadContext();
  var c = loaded.context;
  var none = c.computeStepAvailability_({ step: 'NONE', errorCount: 0, checkedAt: '', calculatedAt: '', pdfAt: '', draftAt: '', needsRedraft: false });
  assert('T13: 初期状態ではCHECKのみ有効', none.CHECK === true && none.CALC === false && none.PDF === false && none.DRAFT === false && none.FINALIZE === false, JSON.stringify(none));

  var afterCheckWithErrors = c.computeStepAvailability_({ step: 'CHECKED_NG', errorCount: 2, checkedAt: 'x', calculatedAt: '', pdfAt: '', draftAt: '', needsRedraft: false });
  assert('T13: チェックでエラーがあればCALCは無効のまま', afterCheckWithErrors.CALC === false, JSON.stringify(afterCheckWithErrors));

  var afterCheckOk = c.computeStepAvailability_({ step: 'CHECKED_OK', errorCount: 0, checkedAt: 'x', calculatedAt: '', pdfAt: '', draftAt: '', needsRedraft: false });
  assert('T13: チェックがOKならCALCが有効になる', afterCheckOk.CALC === true && afterCheckOk.PDF === false, JSON.stringify(afterCheckOk));

  var afterDraftNeedsRedraft = c.computeStepAvailability_({ step: 'DRAFTED', errorCount: 0, checkedAt: 'x', calculatedAt: 'x', pdfAt: 'x', draftAt: 'x', needsRedraft: true });
  assert('T13: 再計算で金額が変わった状態ではFINALIZEが無効になる', afterDraftNeedsRedraft.FINALIZE === false, JSON.stringify(afterDraftNeedsRedraft));
})();

// =====================================================================
// T14: 端数処理後も合計=小計+税が常に成り立つこと（3方式＋非課税）
// =====================================================================
(function T14_roundingInvariant() {
  var loaded = loadContext();
  var c = loaded.context;
  var instructor = { id: 'T01', name: '講師01', rank: 'A', email: 't01@example.com', status: '有効' };
  // 単価を敢えて非整数にし、小計自体の端数処理も検証する（小計301.5円、税10%なら30.15円）
  var itemMap = { I01: { id: 'I01', name: '個別指導60分', unit: '回', rates: { A: 100.5, B: 0, C: 0 } } };
  var rows = [['2026-09-01', 'T01', 'I01', 3, '']];

  ['切上げ', '切下げ', '四捨五入'].forEach(function (method) {
    var result = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '課税', taxRate: 0.1, roundingMethod: method });
    assert('T14: ' + method + 'でも合計=小計+税が成り立つ', result.total === result.subtotal + result.tax, JSON.stringify(result));
  });

  var noTax = c.computeInstructorResult_(instructor, rows, itemMap, { taxCategory: '非課税', taxRate: 0.1, roundingMethod: '切上げ' });
  assert('T14: 非課税でも合計=小計+税(税0円)が成り立つ', noTax.total === noTax.subtotal + noTax.tax && noTax.tax === 0, JSON.stringify(noTax));
})();

// =====================================================================
// T15: 日次の作成上限に達すると処理を中断し、進捗を保存してエラーログに理由を書く
// =====================================================================
(function T15_dailyLimitStop() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  setSettingValue_(loaded, 'Drive作成日次上限', 5);
  c.runCheck();
  c.runCalculate();

  var pdfState = c.runPdf();
  assert('T15: 日次上限に達するとPDF_DONEにならず中断する', pdfState.step !== 'PDF_DONE', JSON.stringify(pdfState));
  assert('T15: 日次上限（5件）までしかPDFが作られない', countPdfFiles_(loaded) === 5, '件数=' + countPdfFiles_(loaded));

  var progress = c.getBatchProgress_('202608', 'PDF');
  assert('T15: 中断した位置（5件目まで処理済み）が進捗として保存される', progress && progress.nextIndex === 5, JSON.stringify(progress));

  var errorLog = c.getActiveSs_().getSheetByName('エラーログ');
  var hasLimitLog = errorLog && errorLog.__dumpRows().some(function (r) { return String(r[4] || '').indexOf('日次上限') !== -1; });
  assert('T15: 日次上限で中断した理由がエラーログに記録される', !!hasLimitLog);
})();

// =====================================================================
// T16: 日次上限で中断した翌日に再実行しても、二重作成せず続きから再開する
// =====================================================================
(function T16_resumeNextDayNoDuplicate() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  setSettingValue_(loaded, 'Drive作成日次上限', 5);
  c.runCheck();
  c.runCalculate();
  c.runPdf(); // 5件作成して中断（T15と同条件）

  var pdfMapDay1 = JSON.parse(JSON.stringify(c.getIdempotencyMap_('PDF_IDS_', '202608')));
  assert('T16: 1日目は5件のPDFが記録される', Object.keys(pdfMapDay1).length === 5, JSON.stringify(pdfMapDay1));

  // 「翌日」になったことを模す（日次カウントのキーは日付を含むため、日付が変われば0から数え直される）
  c.todayStr_ = function () { return '2026-08-03'; };
  var pdfStateDay2 = c.runPdf();
  assert('T16: 翌日の再開でもまだ19件に達していなければPDF_DONEにならない', pdfStateDay2.step !== 'PDF_DONE', JSON.stringify(pdfStateDay2));
  assert('T16: 翌日の再開で、同日の上限（5件）分だけ追加作成される（重複なし）', countPdfFiles_(loaded) === 10, '件数=' + countPdfFiles_(loaded));

  var pdfMapDay2 = c.getIdempotencyMap_('PDF_IDS_', '202608');
  Object.keys(pdfMapDay1).forEach(function (id) {
    assert('T16: 1日目に作成済みの講師(' + id + ')のPDFは翌日再実行しても同じファイルIDのまま', pdfMapDay2[id].fileId === pdfMapDay1[id].fileId);
  });

  // 上限を十分大きくして「翌日以降」を繰り返し、最終的に全員分が重複なく作られることを確認する
  c.todayStr_ = function () { return '2026-08-04'; };
  setSettingValue_(loaded, 'Drive作成日次上限', 400);
  var finalPdfState = c.runPdf();
  assert('T16: 十分な上限のもとで続きを再開すると最終的にPDF_DONEになる', finalPdfState.step === 'PDF_DONE', JSON.stringify(finalPdfState));
  assert('T16: 最終的にPDFが19件になる（重複なし）', countPdfFiles_(loaded) === 19, '件数=' + countPdfFiles_(loaded));
})();

// =====================================================================
// T17: 実行前バックアップは同じファイル内のバックアップシートに値を書き写す
// =====================================================================
(function T17_backupSheetsInSameFile() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  c.runPdf();
  c.runDraft();
  loaded.world.currentUserEmail = 'admin@example.com';
  c.runFinalize();
  var names = loaded.world.spreadsheet.getSheets().map(function (sh) { return sh.getName(); });
  var bk = names.filter(function (n) { return n.indexOf('BK_finalize_202608_') === 0; });
  assert('T17: 確定前に計算結果・明細・合計の3枚のバックアップシートが同じファイル内に作られる', bk.length === 3, JSON.stringify(bk));
  var calcBk = bk.filter(function (n) { return /計算結果_202608$/.test(n); })[0];
  var orig = loaded.world.spreadsheet.getSheetByName('計算結果_202608').__dumpRows().length;
  var copy = calcBk ? loaded.world.spreadsheet.getSheetByName(calcBk).__dumpRows().length : -1;
  assert('T17: バックアップシートの行数が元の計算結果シートと一致する', copy === orig, 'orig=' + orig + ' copy=' + copy);
  var driveCopies = Object.keys(loaded.world.files).filter(function (id) { return /^backup_/.test(loaded.world.files[id].name || ''); });
  assert('T17: Driveに別ファイルとしてのバックアップは作られない', driveCopies.length === 0, JSON.stringify(driveCopies));
})();

// =====================================================================
// T18: バックアップに失敗しても計算・確定は止まらず、エラーログに残る
// =====================================================================
(function T18_backupFailureDoesNotStop() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  var ss = loaded.world.spreadsheet;
  var origInsert = ss.insertSheet;
  ss.insertSheet = function (n) {
    if (String(n).indexOf('BK_') === 0) throw new Error('模擬：バックアップシートを作れない');
    return origInsert.apply(ss, arguments);
  };
  c.runCheck();
  var calcState = c.runCalculate();
  assert('T18: バックアップ失敗時も計算が完了しCALCULATEDになる', calcState.step === 'CALCULATED', JSON.stringify(calcState));
  c.runPdf();
  c.runDraft();
  loaded.world.currentUserEmail = 'admin@example.com';
  var finalState = c.runFinalize();
  assert('T18: バックアップ失敗時も確定が完了しFINALIZEDになる', finalState.step === 'FINALIZED', JSON.stringify(finalState));
  var errorLog = ss.getSheetByName('エラーログ');
  var bkLogs = errorLog ? errorLog.__dumpRows().filter(function (r) { return String(r[4] || '').indexOf('バックアップに失敗') !== -1; }) : [];
  assert('T18: バックアップ失敗がエラーログに記録される（初回計算時は元シートがまだ無いので確定時の1回分）', bkLogs.length === 1, '件数=' + bkLogs.length);
})();

// =====================================================================
// T19: 初期設定（setupInitial）が例外なく完了する
// =====================================================================
(function T19_setupInitial() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  var ok = true, msg = '';
  try { c.setupInitial(); } catch (err) { ok = false; msg = err.message; }
  assert('T19: setupInitialが例外なく完了する', ok, msg);
})();

// =====================================================================
// T20: 同じ処理のバックアップシートは最新の1回分だけ残る（他の処理の分は消さない）
// =====================================================================
(function T20_pruneOldBackups() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  var ss = loaded.world.spreadsheet;
  // 1回目：計算結果があるので3枚できる。古い回に見せかけるため名前の日時を過去にする
  var first = c.backupSpreadsheet_('calc_202608', '202608');
  first.forEach(function (n) { ss.getSheetByName(n).setName(n.replace(/^BK_calc_202608_\d{8}_\d{6}_/, 'BK_calc_202608_20000101_000000_')); });
  var other = c.backupSpreadsheet_('finalize_202608', '202608');
  var second = c.backupSpreadsheet_('calc_202608', '202608');
  var names = ss.getSheets().map(function (sh) { return sh.getName(); });
  var calcBk = names.filter(function (n) { return n.indexOf('BK_calc_202608_') === 0; });
  assert('T20: 同じ処理のバックアップは最新の1回分（3枚）だけ残る', calcBk.length === 3 && calcBk.every(function (n) { return second.indexOf(n) !== -1; }), JSON.stringify(calcBk));
  var finBk = names.filter(function (n) { return n.indexOf('BK_finalize_202608_') === 0; });
  assert('T20: 別の処理（確定）のバックアップは消されない', finBk.length === other.length && other.length === 3, JSON.stringify(finBk));
})();

// =====================================================================
// T21: バックアップが3枚そろわなかったときは、古いバックアップを消さない
// =====================================================================
(function T21_noPruneOnPartialBackup() {
  var loaded = freshContextWithFixtures();
  var c = loaded.context;
  setSettingValue_(loaded, '対象年月', '2026-08');
  c.runCheck();
  c.runCalculate();
  var ss = loaded.world.spreadsheet;
  var first = c.backupSpreadsheet_('calc_202608', '202608');
  var oldNames = first.map(function (n) {
    var o = n.replace(/^BK_calc_202608_\d{8}_\d{6}_/, 'BK_calc_202608_20000101_000000_');
    ss.getSheetByName(n).setName(o);
    return o;
  });
  var origInsert = ss.insertSheet;
  ss.insertSheet = function (n) {
    if (String(n).indexOf('BK_') === 0 && String(n).indexOf('明細_') !== -1) throw new Error('模擬：明細のバックアップだけ失敗');
    return origInsert.apply(ss, arguments);
  };
  c.backupSpreadsheet_('calc_202608', '202608');
  ss.insertSheet = origInsert;
  var names = ss.getSheets().map(function (sh) { return sh.getName(); });
  var oldLeft = oldNames.filter(function (n) { return names.indexOf(n) !== -1; });
  assert('T21: 一部のバックアップが失敗したときは、前回の3枚がすべて残る', oldLeft.length === 3, JSON.stringify(oldLeft));
  var errorLog = ss.getSheetByName('エラーログ');
  var logged = errorLog && errorLog.__dumpRows().some(function (r) { return String(r[4] || '').indexOf('バックアップに失敗') !== -1; });
  assert('T21: 失敗はエラーログに記録される', !!logged);
})();

// =====================================================================
// T22: 日付はスクリプトのタイムゾーンで読む（東京以外でも年月・日がずれない）
// =====================================================================
(function T22_scriptTimeZone() {
  var loaded = loadContext();
  var c = loaded.context;
  // シート側の日付と同じく、GASの中の Date で作る
  function d(ms) { return vm.runInContext('new Date(' + ms + ')', c); }
  // 2026-09-01 00:00 JST は、UTCでは 2026-08-31 15:00、ニューヨークでは 2026-08-31 11:00
  var jstMidnight = d(Date.UTC(2026, 7, 31, 15, 0, 0));
  loaded.world.scriptTimeZone = 'Asia/Tokyo';
  assert('T22: 東京のスクリプトでは、日本時間0時の日付が9月として読まれる', c.normalizeMonth_(jstMidnight) === '202609', c.normalizeMonth_(jstMidnight));
  assert('T22: 東京のスクリプトでは、tz_() が Asia/Tokyo を返す', c.tz_() === 'Asia/Tokyo');
  // ニューヨークの会社がシートに 2026-09-01 と入れた値（現地0時）
  var nyMidnight = d(Date.UTC(2026, 8, 1, 4, 0, 0));
  loaded.world.scriptTimeZone = 'America/New_York';
  assert('T22: ニューヨークのスクリプトでは、現地0時の日付が9月のまま読まれる', c.normalizeMonth_(nyMidnight) === '202609', c.normalizeMonth_(nyMidnight));
  assert('T22: ニューヨークのスクリプトでは、月末の現地23時が翌月にずれない', c.normalizeMonth_(d(Date.UTC(2026, 9, 1, 3, 0, 0))) === '202609');
  assert('T22: todayStr_ が yyyy-MM-dd の形で返る', /^\d{4}-\d{2}-\d{2}$/.test(c.todayStr_()), c.todayStr_());
})();

console.log('');
var okCount = results.filter(function (r) { return r.ok; }).length;
var ngCount = results.length - okCount;
console.log(okCount + ' OK / ' + ngCount + ' NG (total ' + results.length + ')');
if (ngCount > 0) {
  console.error('一部の検証がNGです。上記の[NG]行を確認してください。');
  process.exitCode = 1;
}
