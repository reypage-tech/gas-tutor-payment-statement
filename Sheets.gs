/**
 * シートアクセス共通処理。
 * 原データ（実績・マスタ・設定）は読むだけで、書き込まない（古物台帳案件の
 * 「マスタ非破損」要件に対応）。script が作る中間シート（計算結果・明細・チェック結果・
 * 合計・エラーログ・監査ログ・ダッシュボード）だけを insertSheet/clearContents する。
 */

function getActiveSs_() {
  return SpreadsheetApp.getActive();
}

function getSheetOrThrow_(name) {
  var sh = getActiveSs_().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

/** 見出し行を除いた全行を返す（空シートなら空配列） */
function getRowsBelowHeader_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

/** 「設定」シート（項目|値|説明）を { 項目: 値 } の連想配列で返す */
function readSettings_() {
  var sheet = getSheetOrThrow_(SHEET_SETTINGS);
  var rows = getRowsBelowHeader_(sheet);
  var settings = {};
  rows.forEach(function (r) {
    var key = String(r[0] || '').trim();
    if (key) settings[key] = r[1];
  });
  return settings;
}

function getSettingsWithDefaults_() {
  var s = readSettings_();
  return {
    targetMonth: normalizeMonth_(s['対象年月']),
    roundingMethod: s['端数処理'] || '四捨五入',
    taxCategory: s['課税区分'] || '非課税',
    taxRate: toNumber_(s['消費税率']) || 0.1,
    prevMonthRatioHigh: toNumber_(s['前月比異常閾値倍率']) || 2.0,
    prevMonthRatioLow: s['前月比異常閾値下限'] === undefined || s['前月比異常閾値下限'] === '' ? 0.3 : toNumber_(s['前月比異常閾値下限']),
    pdfFolderId: String(s['PDF保存先フォルダID'] || '').trim(),
    dailyPdfLimit: toNumber_(s['Drive作成日次上限']) || DEFAULT_DAILY_PDF_LIMIT,
    dailyDraftLimit: toNumber_(s['Gmail下書き日次上限']) || DEFAULT_DAILY_DRAFT_LIMIT,
    adminEmails: String(s['管理者メールアドレス'] || ''),
    staffEmails: String(s['担当者メールアドレス'] || ''),
    raw: s
  };
}

function readInstructors_() {
  var sheet = getSheetOrThrow_(SHEET_MASTER_INSTRUCTORS);
  return getRowsBelowHeader_(sheet).filter(function (r) { return String(r[COL_INSTR.ID - 1] || '').trim() !== ''; });
}

function readInstructorMap_() {
  var map = {};
  readInstructors_().forEach(function (r) {
    var id = String(r[COL_INSTR.ID - 1]).trim();
    map[id] = {
      id: id,
      name: r[COL_INSTR.NAME - 1],
      rank: String(r[COL_INSTR.RANK - 1] || '').trim(),
      email: String(r[COL_INSTR.EMAIL - 1] || '').trim(),
      status: String(r[COL_INSTR.STATUS - 1] || '').trim()
    };
  });
  return map;
}

function readItems_() {
  var sheet = getSheetOrThrow_(SHEET_MASTER_ITEMS);
  return getRowsBelowHeader_(sheet).filter(function (r) { return String(r[COL_ITEM.ID - 1] || '').trim() !== ''; });
}

function readItemMap_() {
  var map = {};
  readItems_().forEach(function (r) {
    var id = String(r[COL_ITEM.ID - 1]).trim();
    map[id] = {
      id: id,
      name: r[COL_ITEM.NAME - 1],
      unit: r[COL_ITEM.UNIT - 1],
      rates: { A: r[COL_ITEM.RATE_A - 1], B: r[COL_ITEM.RATE_B - 1], C: r[COL_ITEM.RATE_C - 1] }
    };
  });
  return map;
}

/** 実績_<yyyyMM> シートの全行。存在しなければ空配列（前月データが無い場合に使う） */
function readPerformanceRows_(yyyyMM) {
  var sheet = getActiveSs_().getSheetByName(SHEET_PREFIX_PERFORMANCE + yyyyMM);
  if (!sheet) return [];
  return getRowsBelowHeader_(sheet);
}

/** script が作り直す中間シートを取得する。無ければ作成し、あれば全消去して返す（冪等な作り直し） */
function getOrResetOwnedSheet_(name, headerRow) {
  var ss = getActiveSs_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormat();
    sheet.getConditionalFormatRules && sheet.setConditionalFormatRules([]);
  }
  if (headerRow && headerRow.length) {
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
  }
  return sheet;
}

/** 追記専用シート（エラーログ・監査ログ）。無ければヘッダ付きで作る。あれば何もしない */
function getOrCreateAppendSheet_(name, headerRow) {
  var ss = getActiveSs_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
  }
  return sheet;
}

function appendErrorLog_(functionName, instructorId, targetMonth, detail) {
  var sheet = getOrCreateAppendSheet_(SHEET_ERROR_LOG, ['時刻', '関数', '講師ID', '対象月', '内容']);
  sheet.appendRow([nowStr_(), functionName, instructorId || '', targetMonth || '', String(detail)]);
}

function appendAuditLog_(operation, targetMonth, detail) {
  var sheet = getOrCreateAppendSheet_(SHEET_AUDIT_LOG, ['時刻', '操作者', '操作', '対象月', '詳細']);
  var email = Session.getEffectiveUser().getEmail();
  sheet.appendRow([nowStr_(), email, operation, targetMonth || '', String(detail)]);
}

/** 書き込みを1回にまとめて値を流し込む（getValues/setValues一括。行数が増えても遅くならない） */
function writeRows_(sheet, startRow, startCol, rows) {
  if (!rows.length) return;
  sheet.getRange(startRow, startCol, rows.length, rows[0].length).setValues(rows);
}

/** 「計算結果_yyyyMM」シートを 講師ID -> {name,rank,qtyTotal,subtotal,tax,total} の連想配列で返す */
function readCalcResultsMap_(yyyyMM) {
  var sheet = getSheetOrThrow_(SHEET_PREFIX_CALC + yyyyMM);
  var map = {};
  getRowsBelowHeader_(sheet).forEach(function (r) {
    var id = String(r[COL_CALC.INSTRUCTOR_ID - 1]);
    map[id] = {
      name: r[COL_CALC.NAME - 1],
      rank: r[COL_CALC.RANK - 1],
      qtyTotal: r[COL_CALC.QTY_TOTAL - 1],
      subtotal: r[COL_CALC.SUBTOTAL - 1],
      tax: r[COL_CALC.TAX - 1],
      total: r[COL_CALC.TOTAL - 1]
    };
  });
  return map;
}

/** 「明細_yyyyMM」シートを 講師ID -> [{itemId,itemName,unitPrice,qty,amount}] の連想配列で返す */
function readDetailLinesMap_(yyyyMM) {
  var sheet = getSheetOrThrow_(SHEET_PREFIX_DETAIL + yyyyMM);
  var map = {};
  getRowsBelowHeader_(sheet).forEach(function (r) {
    var id = String(r[COL_DETAIL.INSTRUCTOR_ID - 1]);
    if (!map[id]) map[id] = [];
    map[id].push({
      itemId: r[COL_DETAIL.ITEM_ID - 1],
      itemName: r[COL_DETAIL.ITEM_NAME - 1],
      unitPrice: r[COL_DETAIL.UNIT_PRICE - 1],
      qty: r[COL_DETAIL.QTY - 1],
      amount: r[COL_DETAIL.AMOUNT - 1]
    });
  });
  return map;
}

/**
 * 実行前バックアップ：計算結果・明細・合計シートの「値」を、同じファイル内の非表示の
 * バックアップシート（BK_処理名_日時_元シート名）へ書き写す。
 * 別ファイルへのコピーはしない（drive.file の範囲外になるため）。このファイル自体の権限
 * （spreadsheets.currentonly）だけで動く。
 * バックアップに失敗しても本処理は止めず、エラーログに記録して続行する。
 * 戻り値：作成したバックアップシート名の配列（失敗時は空配列）。
 */
function backupSpreadsheet_(label, yyyyMM) {
  var created = [];
  try {
    var ss = getActiveSs_();
    var stamp = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd_HHmmss');
    var targets = [SHEET_PREFIX_CALC + yyyyMM, SHEET_PREFIX_DETAIL + yyyyMM, SHEET_SUMMARY];
    targets.forEach(function (name) {
      var src = ss.getSheetByName(name);
      if (!src || src.getLastRow() === 0 || src.getLastColumn() === 0) return;
      var values = src.getRange(1, 1, src.getLastRow(), src.getLastColumn()).getValues();
      var bkName = BACKUP_SHEET_PREFIX + label + '_' + stamp + '_' + name;
      var bk = ss.getSheetByName(bkName) || ss.insertSheet(bkName);
      bk.getRange(1, 1, values.length, values[0].length).setValues(values);
      if (typeof bk.hideSheet === 'function') bk.hideSheet();
      created.push(bkName);
    });
    // 対象シートがすべてそろって作れたときだけ、同じ処理の古いバックアップを消す
    var expected = targets.filter(function (name) {
      var src = ss.getSheetByName(name);
      return src && src.getLastRow() > 0 && src.getLastColumn() > 0;
    }).length;
    if (created.length > 0 && created.length === expected && expected === targets.length) {
      pruneOldBackupSheets_(ss, label, stamp);
    }
  } catch (err) {
    try {
      appendErrorLog_('backupSpreadsheet_', '', yyyyMM, 'バックアップに失敗（処理は続行）: ' + (err && err.message ? err.message : err));
    } catch (ignore) {}
  }
  return created;
}

/**
 * 同じ処理（label、例：calc_202608）のバックアップシートは最新の1回分だけ残し、
 * それより古い回のシートを削除する。削除に失敗しても本処理は止めない（呼び出し元のtry/catch）。
 */
function pruneOldBackupSheets_(ss, label, keepStamp) {
  var prefix = BACKUP_SHEET_PREFIX + label + '_';
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (n.indexOf(prefix) !== 0) return;
    if (n.indexOf(prefix + keepStamp + '_') === 0) return;
    ss.deleteSheet(sh);
  });
}

