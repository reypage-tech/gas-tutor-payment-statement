/**
 * 事前チェック（ドライラン）。実績データをマスタと突き合わせ、
 * 未登録講師・単価欠落・負数・重複行・前月比異常値を「チェック結果」シートに一覧する。
 * エラー（深刻度=エラー）が1件でもあれば、計算以降のボタンは押せない（state.errorCountで判定）。
 *
 * computeChecks_ はGASのサービスに依存しない純粋関数にしてあり、node側のテスト
 * （test/run.js）からGASの実行環境なしに直接呼び出して検証できる。
 */

/**
 * @param {Array[]} rows 実績シートの全行（日付,講師ID,品目ID,件数,メモ）
 * @param {Object} instructorMap 講師ID -> {id,name,rank,email,status}
 * @param {Object} itemMap 品目ID -> {id,name,unit,rates:{A,B,C}}
 * @param {Array[]} prevRows 前月の実績シートの全行（無ければ空配列）
 * @param {Object} settings getSettingsWithDefaults_() の戻り値
 * @returns {{issues: Array, instructorTotals: Object}}
 */
function computeChecks_(rows, instructorMap, itemMap, prevRows, settings) {
  var issues = [];
  var dupKeyCount = {};
  rows.forEach(function (r) {
    var key = [r[COL_PERF.DATE - 1], r[COL_PERF.INSTRUCTOR_ID - 1], r[COL_PERF.ITEM_ID - 1]].join('\t');
    dupKeyCount[key] = (dupKeyCount[key] || 0) + 1;
  });

  rows.forEach(function (r, idx) {
    var rowNo = idx + 2; // ヘッダ行(1行目)を除いた実際のシート行番号
    var instructorId = String(r[COL_PERF.INSTRUCTOR_ID - 1] || '').trim();
    var itemId = String(r[COL_PERF.ITEM_ID - 1] || '').trim();
    var qtyRaw = r[COL_PERF.QTY - 1];
    var instructor = instructorMap[instructorId];
    var item = itemMap[itemId];

    if (!instructor) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.UNKNOWN_INSTRUCTOR, instructorId, itemId,
        'マスタ_講師に存在しない講師ID「' + instructorId + '」です'));
    } else if (instructor.status !== INSTRUCTOR_STATUS_ACTIVE) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.INACTIVE_INSTRUCTOR, instructorId, itemId,
        '講師「' + instructorId + '」は契約状態が「' + instructor.status + '」です'));
    }

    if (!item) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.UNKNOWN_ITEM, instructorId, itemId,
        'マスタ_品目に存在しない品目ID「' + itemId + '」です'));
    } else if (instructor && instructor.status === INSTRUCTOR_STATUS_ACTIVE) {
      var rate = item.rates[instructor.rank];
      if (isBlank_(rate) || isNaN(Number(rate))) {
        issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.MISSING_RATE, instructorId, itemId,
          '品目「' + itemId + '」にランク' + instructor.rank + 'の単価が設定されていません'));
      }
    }

    var qtyNum = Number(qtyRaw);
    if (qtyRaw === '' || qtyRaw === null || isNaN(qtyNum)) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.INVALID_QTY, instructorId, itemId,
        '件数が数値ではありません（値: ' + qtyRaw + '）'));
    } else if (qtyNum < 0) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.NEGATIVE_QTY, instructorId, itemId,
        '件数が負数です（値: ' + qtyNum + '）'));
    }

    var key = [r[COL_PERF.DATE - 1], instructorId, itemId].join('\t');
    if (dupKeyCount[key] > 1) {
      issues.push(makeIssue_(rowNo, SEVERITY_ERROR, CHECK_TYPE.DUPLICATE_ROW, instructorId, itemId,
        '同じ日付・講師・品目の行が' + dupKeyCount[key] + '件あります（重複の可能性）'));
    }
  });

  var curTotals = sumQtyByInstructor_(rows);
  var prevTotals = sumQtyByInstructor_(prevRows);
  Object.keys(curTotals).forEach(function (instructorId) {
    var cur = curTotals[instructorId];
    var prev = prevTotals[instructorId] || 0;
    if (prev <= 0) return; // 前月実績が無い（新規講師等）は比較対象にしない
    var ratio = cur / prev;
    if (ratio >= settings.prevMonthRatioHigh || ratio <= settings.prevMonthRatioLow) {
      issues.push(makeIssue_('-', SEVERITY_WARNING, CHECK_TYPE.PREV_MONTH_ANOMALY, instructorId, '',
        '前月比 ' + ratio.toFixed(2) + ' 倍（前月' + prev + '件、今月' + cur + '件）。要確認'));
    }
  });

  return { issues: issues, instructorTotals: curTotals };
}

function makeIssue_(rowNo, severity, type, instructorId, itemId, detail) {
  return { rowNo: rowNo, severity: severity, type: type, instructorId: instructorId, itemId: itemId, detail: detail };
}

function sumQtyByInstructor_(rows) {
  var totals = {};
  rows.forEach(function (r) {
    var instructorId = String(r[COL_PERF.INSTRUCTOR_ID - 1] || '').trim();
    if (!instructorId) return;
    var qty = toNumber_(r[COL_PERF.QTY - 1]);
    totals[instructorId] = (totals[instructorId] || 0) + qty;
  });
  return totals;
}

// ===== GAS側の入出力（シート読み書き・状態保存） =====

/** メニュー「①チェック」から呼ばれる */
function runCheck() {
  return withScriptLock_(function () {
    var settings = getSettingsWithDefaults_();
    var yyyyMM = settings.targetMonth;
    var instructorMap = readInstructorMap_();
    var itemMap = readItemMap_();
    var rows = readPerformanceRows_(yyyyMM);
    var prevRows = readPerformanceRows_(previousMonth_(yyyyMM));

    var result = computeChecks_(rows, instructorMap, itemMap, prevRows, settings);
    writeCheckSheet_(yyyyMM, result.issues);

    var errorCount = result.issues.filter(function (i) { return i.severity === SEVERITY_ERROR; }).length;
    var warningCount = result.issues.filter(function (i) { return i.severity === SEVERITY_WARNING; }).length;
    saveMonthState_(yyyyMM, {
      step: errorCount === 0 ? MONTH_STATE.CHECKED_OK : MONTH_STATE.CHECKED_NG,
      checkedAt: nowStr_(),
      errorCount: errorCount,
      warningCount: warningCount
    });
    appendAuditLog_('チェック実行', yyyyMM, 'エラー' + errorCount + '件・注意' + warningCount + '件');
    return getMonthState_(yyyyMM);
  });
}

function writeCheckSheet_(yyyyMM, issues) {
  var sheet = getOrResetOwnedSheet_(SHEET_PREFIX_CHECK + yyyyMM, ['行番号', '深刻度', '種別', '講師ID', '品目ID', '詳細']);
  if (issues.length) {
    var rows = issues.map(function (i) { return [i.rowNo, i.severity, i.type, i.instructorId, i.itemId, i.detail]; });
    writeRows_(sheet, 2, 1, rows);
  }
  applyCheckConditionalFormat_(sheet, issues.length);
}

/** 深刻度列の値に応じて赤（エラー）・黄（注意）の条件付き書式を付ける */
function applyCheckConditionalFormat_(sheet, issueCount) {
  var numRows = Math.max(issueCount, 1);
  var errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(SEVERITY_ERROR)
    .setBackground('#f4c7c3')
    .setRanges([sheet.getRange(2, 2, numRows, 1)])
    .build();
  var warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(SEVERITY_WARNING)
    .setBackground('#fce8b2')
    .setRanges([sheet.getRange(2, 2, numRows, 1)])
    .build();
  sheet.setConditionalFormatRules([errorRule, warningRule]);
}
