/**
 * 報酬計算。実績（件数）×単価マスタ（ランク別）で講師ごとの報酬を計算する。
 * 端数処理・税区分は「設定」シートで切替える。
 * computeInstructorResult_ はGASサービスに依存しない純粋関数（test/run.jsから直接検証できる）。
 */

/**
 * @param {Object} instructor {id,name,rank,email,status}
 * @param {Array[]} perfRows この講師・対象月の実績行（日付,講師ID,品目ID,件数,メモ）
 * @param {Object} itemMap 品目ID -> {id,name,unit,rates:{A,B,C}}
 * @param {Object} settings getSettingsWithDefaults_() の戻り値
 */
function computeInstructorResult_(instructor, perfRows, itemMap, settings) {
  var qtyTotal = 0;
  var subtotal = 0;
  var itemAgg = {};

  perfRows.forEach(function (r) {
    var itemId = String(r[COL_PERF.ITEM_ID - 1] || '').trim();
    var item = itemMap[itemId];
    if (!item) return; // チェック済み運用では通常発生しない（未登録品目はチェックでエラーになる）
    var qty = toNumber_(r[COL_PERF.QTY - 1]);
    var rate = toNumber_(item.rates[instructor.rank]);
    var amount = qty * rate;
    qtyTotal += qty;
    subtotal += amount;
    if (!itemAgg[itemId]) itemAgg[itemId] = { itemId: itemId, itemName: item.name, unitPrice: rate, qty: 0, amount: 0 };
    itemAgg[itemId].qty += qty;
    itemAgg[itemId].amount += amount;
  });

  var lines = Object.keys(itemAgg).sort().map(function (id) { return itemAgg[id]; });
  var tax = settings.taxCategory === '課税' ? subtotal * settings.taxRate : 0;
  // 小計・税をそれぞれ設定の端数処理方式で丸め、丸めた値をそのまま足して合計にする。
  // （生の小計+税を丸めて合計にすると、表示される小計・税の合計と合計が一致しなくなるため）
  var roundedSubtotal = roundAmount_(subtotal, settings.roundingMethod);
  var roundedTax = roundAmount_(tax, settings.roundingMethod);
  var total = roundedSubtotal + roundedTax;

  return {
    instructorId: instructor.id,
    name: instructor.name,
    rank: instructor.rank,
    qtyTotal: qtyTotal,
    subtotal: roundedSubtotal,
    tax: roundedTax,
    total: total,
    lines: lines
  };
}

function groupPerfByInstructor_(rows) {
  var map = {};
  rows.forEach(function (r) {
    var id = String(r[COL_PERF.INSTRUCTOR_ID - 1] || '').trim();
    if (!map[id]) map[id] = [];
    map[id].push(r);
  });
  return map;
}

/** 確定済みの月を管理者が再計算する場合は、保護を外してから進める（履歴は監査ログに残す） */
function ensureNotFinalizedOrUnlock_(yyyyMM, settings) {
  if (!isMonthFinalized_(yyyyMM)) return;
  if (!isAdmin_(settings)) {
    throw new Error('対象月「' + formatMonthLabel_(yyyyMM) + '」は確定済みです。再計算は管理者のみ可能です。');
  }
  unprotectMonthSheets_(yyyyMM);
  unmarkMonthFinalized_(yyyyMM);
  appendAuditLog_('確定月の再計算(管理者)', yyyyMM, '保護を解除して再計算を許可');
}

/** メニュー「②計算」から呼ばれる */
function runCalculate(yyyyMMOverride) {
  return withScriptLock_(function () {
    var settings = getSettingsWithDefaults_();
    var yyyyMM = yyyyMMOverride || settings.targetMonth;
    var state = getMonthState_(yyyyMM);
    if (state.errorCount > 0 || !state.checkedAt) {
      throw new Error('先に①チェックを実行し、エラー0件にしてください。');
    }
    ensureNotFinalizedOrUnlock_(yyyyMM, settings);

    var isFreshStart = !getBatchProgress_(yyyyMM, 'CALC');
    if (isFreshStart) {
      backupSpreadsheet_('calc_' + yyyyMM, yyyyMM);
      getOrResetOwnedSheet_(SHEET_PREFIX_CALC + yyyyMM, ['講師ID', '表示名', 'ランク', '件数合計', '小計', '消費税', '合計', 'ステータス']);
      getOrResetOwnedSheet_(SHEET_PREFIX_DETAIL + yyyyMM, ['講師ID', '品目ID', '品目名', '単価', '件数', '金額']);
    }

    var calcSheet = getSheetOrThrow_(SHEET_PREFIX_CALC + yyyyMM);
    var detailSheet = getSheetOrThrow_(SHEET_PREFIX_DETAIL + yyyyMM);
    var instructorMap = readInstructorMap_();
    var activeIds = Object.keys(instructorMap)
      .filter(function (id) { return instructorMap[id].status === INSTRUCTOR_STATUS_ACTIVE; })
      .sort();
    var itemMap = readItemMap_();
    var perfByInstructor = groupPerfByInstructor_(readPerformanceRows_(yyyyMM));

    var batchResult = runBatchOverInstructors_(yyyyMM, 'CALC', activeIds, function (instructorId) {
      var instructor = instructorMap[instructorId];
      var result = computeInstructorResult_(instructor, perfByInstructor[instructorId] || [], itemMap, settings);
      if (result.lines.length) {
        var detailRows = result.lines.map(function (l) {
          return [instructor.id, l.itemId, l.itemName, l.unitPrice, l.qty, l.amount];
        });
        writeRows_(detailSheet, detailSheet.getLastRow() + 1, 1, detailRows);
      }
      calcSheet.appendRow([instructor.id, instructor.name, instructor.rank, result.qtyTotal, result.subtotal, result.tax, result.total, ROW_STATUS.CALCULATED]);
      upsertSummaryRow_(yyyyMM, instructor, result.total, ROW_STATUS.CALCULATED);
    });

    if (!batchResult.done) {
      return getMonthState_(yyyyMM); // 続きはresumeBatchトリガーが実行する
    }
    finalizeCalculation_(yyyyMM, batchResult.failedIds);
    return getMonthState_(yyyyMM);
  });
}

function finalizeCalculation_(yyyyMM, failedIds) {
  var calcSheet = getSheetOrThrow_(SHEET_PREFIX_CALC + yyyyMM);
  var rows = getRowsBelowHeader_(calcSheet);
  var hashMap = {};
  rows.forEach(function (r) {
    hashMap[String(r[COL_CALC.INSTRUCTOR_ID - 1])] = simpleHash_({ total: r[COL_CALC.TOTAL - 1], qty: r[COL_CALC.QTY_TOTAL - 1] });
  });
  saveIdempotencyMap_('CALC_HASH_', yyyyMM, hashMap);

  var pdfMap = getIdempotencyMap_(PROP_PDF_IDS_PREFIX, yyyyMM);
  var needsRedraft = Object.keys(pdfMap).some(function (id) { return pdfMap[id].hash !== hashMap[id]; });

  saveMonthState_(yyyyMM, { step: MONTH_STATE.CALCULATED, calculatedAt: nowStr_(), needsRedraft: needsRedraft });
  appendAuditLog_('計算実行', yyyyMM, rows.length + '名、失敗' + (failedIds ? failedIds.length : 0) + '件');
  refreshDashboard_();
}

function getCalcHash_(yyyyMM) {
  return getIdempotencyMap_('CALC_HASH_', yyyyMM);
}

/** 「合計」シートの対象月・講師の行を更新する。無ければ追加する（振込チェック欄は人が付けるので保持する） */
function upsertSummaryRow_(yyyyMM, instructor, total, status) {
  var sheet = getOrCreateAppendSheet_(SHEET_SUMMARY, ['年月', '講師ID', '表示名', '報酬合計', '振込チェック', 'ステータス', '更新日時']);
  var rows = getRowsBelowHeader_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL_SUMMARY.MONTH - 1]) === yyyyMM && String(rows[i][COL_SUMMARY.INSTRUCTOR_ID - 1]) === instructor.id) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, COL_SUMMARY.NAME).setValue(instructor.name);
      sheet.getRange(rowNum, COL_SUMMARY.TOTAL).setValue(total);
      sheet.getRange(rowNum, COL_SUMMARY.STATUS).setValue(status);
      sheet.getRange(rowNum, COL_SUMMARY.UPDATED_AT).setValue(nowStr_());
      return;
    }
  }
  sheet.appendRow([yyyyMM, instructor.id, instructor.name, total, false, status, nowStr_()]);
}

function updateSummaryStatus_(yyyyMM, instructorId, status) {
  var sheet = getOrCreateAppendSheet_(SHEET_SUMMARY, ['年月', '講師ID', '表示名', '報酬合計', '振込チェック', 'ステータス', '更新日時']);
  var rows = getRowsBelowHeader_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL_SUMMARY.MONTH - 1]) === yyyyMM && String(rows[i][COL_SUMMARY.INSTRUCTOR_ID - 1]) === instructorId) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, COL_SUMMARY.STATUS).setValue(status);
      sheet.getRange(rowNum, COL_SUMMARY.UPDATED_AT).setValue(nowStr_());
      return;
    }
  }
}
