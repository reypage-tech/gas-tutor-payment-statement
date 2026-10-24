/**
 * 講師別PDF（A4縦1枚、架空ロゴ枠「Sample School」）。
 *
 * PDF化の方式について：
 * Googleスライドは新規作成時の既定サイズ（16:9相当）からページサイズを変更する手段が無く、
 * A4縦にできない。一方Googleドキュメントは Body#setPageWidth/setPageHeight で明示的に
 * A4（595.3pt×841.9pt）を指定できるため、本見本ではGoogleドキュメントの雛形をコピーして
 * 値を差し込み、Drive高度なサービス（Drive.Files.export）でPDF化する方式を採用する。
 * ドキュメントの本文編集はDocumentApp（documentsスコープ）を使い、ファイルのコピー・
 * アップロード・ゴミ箱移動はDrive.Files.*（drive.fileスコープ）を使う。
 */

var PAGE_WIDTH_A4_PT = 595.3;
var PAGE_HEIGHT_A4_PT = 841.9;

var PLACEHOLDER = {
  TARGET_MONTH: '##TARGET_MONTH##',
  INSTRUCTOR_NAME: '##INSTRUCTOR_NAME##',
  INSTRUCTOR_ID: '##INSTRUCTOR_ID##',
  RANK: '##RANK##',
  ISSUED_DATE: '##ISSUED_DATE##',
  QTY_TOTAL: '##QTY_TOTAL##',
  SUBTOTAL: '##SUBTOTAL##',
  TAX: '##TAX##',
  TOTAL: '##TOTAL##'
};

/** 講師別PDFのファイル名。Draft.gsの添付ファイル名もこれと揃える */
function buildPdfFileName_(yyyyMM, instructor) {
  return yyyyMM + '_' + instructor.id + '_' + instructor.name + '_報酬明細.pdf';
}

/**
 * 雛形ドキュメントのIDを返す。ScriptPropertiesに保存済みのIDがあればそれを使い
 * （削除済み等で開けない場合は作り直す）、無ければ新規作成して保存する。
 */
function getOrCreateTemplateDocId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_TEMPLATE_DOC_ID);
  if (id) {
    try {
      DocumentApp.openById(id);
      return id;
    } catch (err) {
      // 削除済み等。作り直す。
    }
  }
  var newId = buildTemplateDocument_();
  props.setProperty(PROP_TEMPLATE_DOC_ID, newId);
  return newId;
}

function buildTemplateDocument_() {
  var doc = DocumentApp.create('報酬明細テンプレート（見本）');
  var body = doc.getBody();
  body.setPageWidth(PAGE_WIDTH_A4_PT);
  body.setPageHeight(PAGE_HEIGHT_A4_PT);
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(42).setMarginRight(42);

  var logo = body.appendParagraph(SCHOOL_DISPLAY_NAME);
  logo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('月次 報酬・支払明細（見本／架空データ）').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('対象月：' + PLACEHOLDER.TARGET_MONTH);
  body.appendParagraph('講師：' + PLACEHOLDER.INSTRUCTOR_NAME + '（ID: ' + PLACEHOLDER.INSTRUCTOR_ID + ' / ランク ' + PLACEHOLDER.RANK + '）');
  body.appendParagraph('発行日：' + PLACEHOLDER.ISSUED_DATE);
  body.appendParagraph('');
  body.appendTable([['品目', '単価', '件数', '金額']]);
  body.appendParagraph('');
  body.appendParagraph('件数合計：' + PLACEHOLDER.QTY_TOTAL + '件　　小計：' + PLACEHOLDER.SUBTOTAL + '円　　消費税：' + PLACEHOLDER.TAX + '円');
  body.appendParagraph('報酬合計：' + PLACEHOLDER.TOTAL + '円').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('本書はGoogle Apps Script見本により自動生成されたものです（架空データ）。送信前に内容をご確認ください。');
  doc.saveAndClose();
  return doc.getId();
}

/**
 * 講師1名分のPDFを生成し、指定フォルダに保存する。生成後、作業用に作ったDocumentのコピーは
 * ゴミ箱に移動する（PDFだけを残す）。
 */
function buildPdfForInstructor_(yyyyMM, instructor, calcRow, lines, settings) {
  var templateId = getOrCreateTemplateDocId_();
  var tempName = 'tmp_reward_pdf_' + instructor.id + '_' + yyyyMM;
  var tempFileId = Drive.Files.copy({ name: tempName }, templateId).id;
  var doc = DocumentApp.openById(tempFileId);
  var body = doc.getBody();

  body.replaceText(PLACEHOLDER.TARGET_MONTH, formatMonthLabel_(yyyyMM));
  body.replaceText(PLACEHOLDER.INSTRUCTOR_NAME, String(instructor.name));
  body.replaceText(PLACEHOLDER.INSTRUCTOR_ID, String(instructor.id));
  body.replaceText(PLACEHOLDER.RANK, String(instructor.rank));
  body.replaceText(PLACEHOLDER.ISSUED_DATE, todayStr_());
  body.replaceText(PLACEHOLDER.QTY_TOTAL, String(calcRow.qtyTotal));
  body.replaceText(PLACEHOLDER.SUBTOTAL, Number(calcRow.subtotal).toLocaleString('ja-JP'));
  body.replaceText(PLACEHOLDER.TAX, Number(calcRow.tax).toLocaleString('ja-JP'));
  body.replaceText(PLACEHOLDER.TOTAL, Number(calcRow.total).toLocaleString('ja-JP'));

  var table = body.getTables()[0];
  if (lines.length) {
    lines.forEach(function (l) {
      var row = table.appendTableRow();
      row.appendTableCell(String(l.itemName));
      row.appendTableCell(yen_(l.unitPrice));
      row.appendTableCell(String(l.qty));
      row.appendTableCell(yen_(l.amount));
    });
  } else {
    var emptyRow = table.appendTableRow();
    emptyRow.appendTableCell('（対象期間の実績なし）');
    emptyRow.appendTableCell('-');
    emptyRow.appendTableCell('0');
    emptyRow.appendTableCell('0円');
  }
  doc.saveAndClose();

  var pdfBlob = Drive.Files.export(tempFileId, 'application/pdf');
  var fileName = buildPdfFileName_(yyyyMM, instructor);
  pdfBlob.setName(fileName);
  var monthFolderId = getOrCreateMonthFolderId_(yyyyMM);
  var pdfFile = Drive.Files.create({ name: fileName, parents: [monthFolderId] }, pdfBlob);
  Drive.Files.update({ trashed: true }, tempFileId);

  return { fileId: pdfFile.id, url: driveFileUrl_(pdfFile.id), fileName: fileName };
}

/** メニュー「③PDF作成」から呼ばれる */
function runPdf(yyyyMMOverride) {
  return withScriptLock_(function () {
    var settings = getSettingsWithDefaults_();
    var yyyyMM = yyyyMMOverride || settings.targetMonth;
    var state = getMonthState_(yyyyMM);
    if (!state.calculatedAt) throw new Error('先に②計算を実行してください。');
    if (state.errorCount > 0) throw new Error('チェックのエラーが残っています。①チェックをやり直してください。');
    ensureNotFinalizedOrUnlock_(yyyyMM, settings);

    var calcMap = readCalcResultsMap_(yyyyMM);
    var detailMap = readDetailLinesMap_(yyyyMM);
    var instructorMap = readInstructorMap_();
    var ids = Object.keys(calcMap).sort();
    var hashMap = getCalcHash_(yyyyMM);
    var pdfMap = getIdempotencyMap_(PROP_PDF_IDS_PREFIX, yyyyMM);

    var batchResult = runBatchOverInstructors_(yyyyMM, 'PDF', ids, function (instructorId) {
      var currentHash = hashMap[instructorId];
      var existing = pdfMap[instructorId];
      if (existing && existing.hash === currentHash && driveFileExists_(existing.fileId)) {
        return; // 金額に変更がなく、既存PDFも生きている場合は再作成しない（冪等）
      }
      assertDailyLimitNotReached_('PDF', settings.dailyPdfLimit);
      if (existing && existing.fileId && driveFileExists_(existing.fileId)) {
        Drive.Files.update({ trashed: true }, existing.fileId); // 置換：古いPDFはゴミ箱へ
      }
      var instructor = instructorMap[instructorId];
      var pdf = buildPdfForInstructor_(yyyyMM, instructor, calcMap[instructorId], detailMap[instructorId] || [], settings);
      incrementDailyCount_('PDF');
      pdfMap[instructorId] = { fileId: pdf.fileId, url: pdf.url, hash: currentHash };
      saveIdempotencyMap_(PROP_PDF_IDS_PREFIX, yyyyMM, pdfMap);
      updateSummaryStatus_(yyyyMM, instructorId, ROW_STATUS.PDF_DONE);
    });

    if (!batchResult.done) return getMonthState_(yyyyMM);
    saveMonthState_(yyyyMM, { step: MONTH_STATE.PDF_DONE, pdfAt: nowStr_() });
    appendAuditLog_('PDF作成', yyyyMM, ids.length + '名、失敗' + batchResult.failedIds.length + '件');
    return getMonthState_(yyyyMM);
  });
}
