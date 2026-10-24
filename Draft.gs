/**
 * Gmail下書きの一括作成。Gmail高度なサービス（Gmail.Users.Drafts.create/get/delete、
 * Gmail API v1）を使い、RFC 2822形式のメッセージ（UTF-8、件名はMIMEエンコード、
 * PDF添付はmultipart/mixed）をbase64urlにしてraw に入れて下書きを作成する。
 *
 * gmail.compose スコープは下書きの作成・編集・削除だけでなく下書きの送信も許可するため、
 * 権限の構造だけでは送信を防げない。送信しないことは、このプロジェクトのコード内に
 * 送信（Gmail.Users.Messages.send・Gmail.Users.Drafts.send等）の呼び出しが0件であること
 * （rgで確認できる）と、作成された下書きを人が1通ずつ確認してから送る運用によって担保している。
 */

function buildDraftSubject_(yyyyMM, instructor) {
  return '【' + formatMonthLabel_(yyyyMM) + '】報酬・支払明細のご案内（' + instructor.name + '様）';
}

function buildDraftBody_(yyyyMM, instructor, calcRow) {
  return [
    instructor.name + '様',
    '',
    SCHOOL_DISPLAY_NAME + '事務局です。' + formatMonthLabel_(yyyyMM) + 'の報酬・支払明細をお送りします。',
    '',
    '件数合計：' + calcRow.qtyTotal + '件',
    '小計　　：' + yen_(calcRow.subtotal),
    '消費税　：' + yen_(calcRow.tax),
    '報酬合計：' + yen_(calcRow.total),
    '',
    '明細はPDFを添付しております。内容をご確認ください。',
    '',
    '※このメールはApps Scriptの見本により下書きとして作成されたものです。送信前に担当者が内容を確認します。',
    '（架空データのサンプルです）'
  ].join('\n');
}

/** RFC 2047のエンコードドワード（件名など非ASCIIヘッダのMIMEエンコード） */
function mimeEncodeHeader_(text) {
  return '=?UTF-8?B?' + Utilities.base64Encode(text, Utilities.Charset.UTF_8) + '?=';
}

/** base64文字列を76文字ごとに折り返す（MIMEの行長規約） */
function wrapBase64Lines_(base64) {
  var lines = [];
  for (var i = 0; i < base64.length; i += 76) lines.push(base64.slice(i, i + 76));
  return lines.join('\r\n');
}

/** To・件名・本文・PDF添付を持つ multipart/mixed のRFC 2822メッセージ文字列を組み立てる */
function buildMimeMessage_(to, subject, bodyText, attachment) {
  var boundary = 'boundary_' + Utilities.getUuid().replace(/-/g, '');
  var lines = [
    'To: ' + to,
    'Subject: ' + mimeEncodeHeader_(subject),
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/plain; charset="UTF-8"',
    'Content-Transfer-Encoding: base64',
    '',
    wrapBase64Lines_(Utilities.base64Encode(bodyText, Utilities.Charset.UTF_8)),
    '--' + boundary,
    'Content-Type: ' + attachment.mimeType + '; name="' + attachment.name + '"',
    'Content-Disposition: attachment; filename="' + attachment.name + '"',
    'Content-Transfer-Encoding: base64',
    '',
    wrapBase64Lines_(Utilities.base64Encode(attachment.blob.getBytes())),
    '--' + boundary + '--',
    ''
  ];
  return lines.join('\r\n');
}

/** メッセージ全体をbase64url化する（Gmail APIのraw形式） */
function buildRawMessage_(to, subject, bodyText, attachment) {
  var mime = buildMimeMessage_(to, subject, bodyText, attachment);
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes());
}

function createGmailDraft_(to, subject, bodyText, pdfBlob, pdfFileName) {
  var raw = buildRawMessage_(to, subject, bodyText, { mimeType: 'application/pdf', name: pdfFileName, blob: pdfBlob });
  var draft = Gmail.Users.Drafts.create({ message: { raw: raw } }, 'me');
  return draft.id;
}

function gmailDraftExists_(draftId) {
  try {
    Gmail.Users.Drafts.get('me', draftId);
    return true;
  } catch (err) {
    return false;
  }
}

function deleteGmailDraft_(draftId) {
  Gmail.Users.Drafts.delete('me', draftId);
}

/** メニュー「④下書き作成」から呼ばれる */
function runDraft(yyyyMMOverride) {
  return withScriptLock_(function () {
    var settings = getSettingsWithDefaults_();
    var yyyyMM = yyyyMMOverride || settings.targetMonth;
    var state = getMonthState_(yyyyMM);
    if (!state.pdfAt) throw new Error('先に③PDF作成を実行してください。');
    if (state.errorCount > 0) throw new Error('チェックのエラーが残っています。①チェックをやり直してください。');
    ensureNotFinalizedOrUnlock_(yyyyMM, settings);

    var calcMap = readCalcResultsMap_(yyyyMM);
    var instructorMap = readInstructorMap_();
    var hashMap = getCalcHash_(yyyyMM);
    var pdfMap = getIdempotencyMap_(PROP_PDF_IDS_PREFIX, yyyyMM);
    var draftMap = getIdempotencyMap_(PROP_DRAFT_IDS_PREFIX, yyyyMM);

    Object.keys(calcMap).forEach(function (id) {
      if (!instructorMap[id] || !instructorMap[id].email) {
        appendErrorLog_('下書き作成', id, yyyyMM, 'メールアドレス未設定のため下書きをスキップしました');
      }
    });
    var ids = Object.keys(calcMap)
      .filter(function (id) { return instructorMap[id] && instructorMap[id].email; })
      .sort();

    var batchResult = runBatchOverInstructors_(yyyyMM, 'DRAFT', ids, function (instructorId) {
      var pdfRecord = pdfMap[instructorId];
      if (!pdfRecord || !driveFileExists_(pdfRecord.fileId)) {
        throw new Error('PDFが未作成、または見つかりません。先に③PDF作成を実行してください。');
      }
      var currentHash = hashMap[instructorId];
      var existingDraft = draftMap[instructorId];
      if (existingDraft && existingDraft.hash === currentHash && gmailDraftExists_(existingDraft.draftId)) {
        return; // 金額に変更がなく、既存下書きも残っている場合は再作成しない（冪等）
      }
      assertDailyLimitNotReached_('DRAFT', settings.dailyDraftLimit);
      if (existingDraft && existingDraft.draftId && gmailDraftExists_(existingDraft.draftId)) {
        deleteGmailDraft_(existingDraft.draftId); // 置換：古い下書きは削除してから作り直す
      }
      var instructor = instructorMap[instructorId];
      var calcRow = calcMap[instructorId];
      var fileName = buildPdfFileName_(yyyyMM, instructor);
      var pdfBlob = Drive.Files.get(pdfRecord.fileId, { alt: 'media' });
      pdfBlob.setName(fileName);
      var draftId = createGmailDraft_(
        instructor.email,
        buildDraftSubject_(yyyyMM, instructor),
        buildDraftBody_(yyyyMM, instructor, calcRow),
        pdfBlob,
        fileName
      );
      incrementDailyCount_('DRAFT');
      draftMap[instructorId] = { draftId: draftId, hash: currentHash };
      saveIdempotencyMap_(PROP_DRAFT_IDS_PREFIX, yyyyMM, draftMap);
      updateSummaryStatus_(yyyyMM, instructorId, ROW_STATUS.DRAFTED);
    });

    if (!batchResult.done) return getMonthState_(yyyyMM);
    saveMonthState_(yyyyMM, { step: MONTH_STATE.DRAFTED, draftAt: nowStr_(), needsRedraft: false });
    appendAuditLog_('下書き作成', yyyyMM, ids.length + '名、失敗' + batchResult.failedIds.length + '件');
    return getMonthState_(yyyyMM);
  });
}
