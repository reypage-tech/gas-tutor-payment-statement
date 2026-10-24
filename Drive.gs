/**
 * Driveフォルダ・ファイルの取得・作成。Drive高度なサービス（Drive API v3、Drive.Files.*）
 * だけを使い、drive.file スコープの範囲（このスクリプトが作成したファイル・フォルダ）だけで
 * 動くようにする。設定シートにフォルダIDが指定されていればそれを使い、空欄なら自動作成して
 * IDをPropertiesServiceに保存する（人がIDを手入力する必要はない）。
 * PDFフォルダは共有しない設定をデフォルトとする（既定で「自分のみ」。drive.fileスコープの
 * 範囲内で作成するため、他ユーザーと共有するには別途手動で共有設定が必要）。
 */

var DRIVE_MIME_FOLDER = 'application/vnd.google-apps.folder';

/** Drive APIクエリ文字列に埋め込む値の ' と \ をエスケープする */
function escapeDriveQueryValue_(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function driveFileExists_(fileId) {
  try {
    var f = Drive.Files.get(fileId, { fields: 'id,trashed' });
    return !f.trashed;
  } catch (err) {
    return false;
  }
}

function driveAssertFolderExists_(folderId) {
  if (!driveFileExists_(folderId)) {
    throw new Error('指定されたフォルダが見つかりません（削除された、またはIDが不正です）: ' + folderId);
  }
}

function driveCreateFolder_(name, parentId) {
  var resource = { name: name, mimeType: DRIVE_MIME_FOLDER };
  if (parentId) resource.parents = [parentId];
  return Drive.Files.create(resource).id;
}

/** 指定フォルダ（parentId）直下にある同名フォルダを探す。無ければnull */
function driveFindFolderByName_(name, parentId) {
  var q = "mimeType='" + DRIVE_MIME_FOLDER + "' and trashed=false and name='" + escapeDriveQueryValue_(name) + "'" +
    (parentId ? " and '" + parentId + "' in parents" : '');
  var res = Drive.Files.list({ q: q, fields: 'files(id,name)' });
  return (res.files && res.files.length) ? res.files[0].id : null;
}

function getOrCreatePdfRootFolderId_() {
  var settings = getSettingsWithDefaults_();
  if (settings.pdfFolderId) {
    driveAssertFolderExists_(settings.pdfFolderId);
    return settings.pdfFolderId;
  }
  var props = PropertiesService.getScriptProperties();
  var savedId = props.getProperty(PROP_PDF_ROOT_FOLDER_ID);
  if (savedId && driveFileExists_(savedId)) return savedId;
  var newId = driveCreateFolder_(PDF_ROOT_FOLDER_NAME, null);
  props.setProperty(PROP_PDF_ROOT_FOLDER_ID, newId);
  return newId;
}

function getOrCreateMonthFolderId_(yyyyMM) {
  var rootId = getOrCreatePdfRootFolderId_();
  var name = yyyyMM.slice(0, 4) + '-' + yyyyMM.slice(4, 6);
  var existing = driveFindFolderByName_(name, rootId);
  if (existing) return existing;
  return driveCreateFolder_(name, rootId);
}

function driveFileUrl_(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/view';
}
