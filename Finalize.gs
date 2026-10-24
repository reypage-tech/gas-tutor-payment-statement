/**
 * 確定ロック。確定月は「計算結果」「明細」シートを保護し、管理者以外は再計算できないようにする。
 * 確定前に必ずバックアップを取り、確定操作は監査ログに残す。
 */

function protectMonthSheets_(yyyyMM, settings) {
  var admins = parseEmails_(settings.adminEmails);
  [SHEET_PREFIX_CALC + yyyyMM, SHEET_PREFIX_DETAIL + yyyyMM].forEach(function (name) {
    var sheet = getActiveSs_().getSheetByName(name);
    if (!sheet) return;
    var protection = sheet.protect().setDescription('確定済み(' + yyyyMM + ')：管理者のみ編集可');
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    protection.removeEditors(protection.getEditors());
    if (admins.length) protection.addEditors(admins);
  });
}

function unprotectMonthSheets_(yyyyMM) {
  [SHEET_PREFIX_CALC + yyyyMM, SHEET_PREFIX_DETAIL + yyyyMM].forEach(function (name) {
    var sheet = getActiveSs_().getSheetByName(name);
    if (!sheet) return;
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  });
}

function setAllSummaryStatusForMonth_(yyyyMM, status) {
  var sheet = getOrCreateAppendSheet_(SHEET_SUMMARY, ['年月', '講師ID', '表示名', '報酬合計', '振込チェック', 'ステータス', '更新日時']);
  var rows = getRowsBelowHeader_(sheet);
  rows.forEach(function (r, idx) {
    if (String(r[COL_SUMMARY.MONTH - 1]) !== yyyyMM) return;
    var rowNum = idx + 2;
    sheet.getRange(rowNum, COL_SUMMARY.STATUS).setValue(status);
    sheet.getRange(rowNum, COL_SUMMARY.UPDATED_AT).setValue(nowStr_());
  });
}

/** メニュー「⑤確定」から呼ばれる。管理者のみ実行できる */
function runFinalize(yyyyMMOverride) {
  return withScriptLock_(function () {
    var settings = getSettingsWithDefaults_();
    var yyyyMM = yyyyMMOverride || settings.targetMonth;
    var state = getMonthState_(yyyyMM);

    if (!isAdmin_(settings)) throw new Error('確定は管理者のみ可能です。');
    if (state.errorCount > 0) throw new Error('チェックのエラーが残っています。確定できません。');
    if (!state.draftAt) throw new Error('先に④下書き作成を実行してください。');
    if (state.needsRedraft) throw new Error('再計算により金額が変わっています。④下書き作成をやり直してから確定してください。');
    if (isMonthFinalized_(yyyyMM)) throw new Error('対象月「' + formatMonthLabel_(yyyyMM) + '」は既に確定済みです。');

    backupSpreadsheet_('finalize_' + yyyyMM, yyyyMM);
    protectMonthSheets_(yyyyMM, settings);
    markMonthFinalized_(yyyyMM);
    setAllSummaryStatusForMonth_(yyyyMM, ROW_STATUS.FINALIZED);

    var email = Session.getEffectiveUser().getEmail();
    saveMonthState_(yyyyMM, { step: MONTH_STATE.FINALIZED, finalizedAt: nowStr_(), finalizedBy: email });
    appendAuditLog_('確定', yyyyMM, email + ' が確定しました');
    refreshDashboard_();
    return getMonthState_(yyyyMM);
  });
}
