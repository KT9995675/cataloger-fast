/**
 * §13.5 — создать Google-файл в каталоге.
 * Создаётся через Drive API, переносится во владение CONTROLLER_EMAIL.
 * Доступно любому пользователю каталога с правом редактор на целевую папку
 * (Управляющему — всегда).
 *
 * @param {{
 *   targetFolderId: string,
 *   name: string,
 *   fileType: string
 * }} input
 * @returns {{
 *   ok: true,
 *   catalogId: string,
 *   fileId: string,
 *   displayName: string,
 *   fileType: string,
 *   url: string,
 *   item: Object
 * }}
 */
function createCatalogFile(input) {
  assertCatalogReady_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var name = String(input.name || '').trim();
  var fileType = String(input.fileType || '').trim().toLowerCase();

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (!name) {
    throw catalogError_('INVALID_INPUT', 'Имя файла не может быть пустым.');
  }

  var mimeType = resolveGoogleFileMimeType_(fileType);
  if (!mimeType) {
    throw catalogError_('INVALID_INPUT', 'Неизвестный тип файла: ' + fileType);
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  if (!engine.foldersById[targetFolderId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Target folder not found: ' + targetFolderId);
  }
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, targetFolderId);

  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  if (!controllerEmail) {
    throw catalogError_('CATALOG_NOT_CONFIGURED', 'CONTROLLER_EMAIL is missing.');
  }

  var catalogRootFolderId = getCatalogRootFolderId_();
  var driveFile = createGoogleNativeFileInFolder_(name, mimeType, catalogRootFolderId);
  transferDriveFileToController_(driveFile, controllerEmail, userEmail);

  var catalogId = Utilities.getUuid();
  var displayName = driveFile.getName();
  var fileId = driveFile.getId();
  var sizeBytes = 0;
  try {
    sizeBytes = driveFile.getSize();
  } catch (e) {
    sizeBytes = 0;
  }
  var modifiedAt = driveFile.getLastUpdated();

  appendCatalogFileRow_({
    catalogId: catalogId,
    folderId: targetFolderId,
    fileId: fileId,
    displayName: displayName,
    sizeBytes: sizeBytes,
    driveModifiedAt: modifiedAt,
    sourceFileId: '',
    mimeType: mimeType
  });

  appendExplicitUserAclRow_('file', catalogId, userEmail, 'editor');

  var acl = getEffectiveAclDisplayFromEngine_(engine, 'folder', targetFolderId);
  var editors = (acl.editors || []).slice();
  if (
    editors
      .map(function (e) {
        return String(e).toLowerCase();
      })
      .indexOf(userEmail.toLowerCase()) < 0
  ) {
    editors.push(userEmail);
  }

  return {
    ok: true,
    catalogId: catalogId,
    fileId: fileId,
    displayName: displayName,
    fileType: fileType,
    url: buildCatalogFileOpenUrl_(fileId, mimeType, driveFile.getUrl()),
    item: {
      kind: 'file',
      id: catalogId,
      name: displayName,
      mimeType: mimeType,
      sizeBytes: sizeBytes,
      modifiedAt: formatCatalogDate_(modifiedAt),
      approved: false,
      approvedBy: '',
      isSystem: false,
      editors: editors,
      commenters: acl.commenters || [],
      readers: acl.readers || []
    }
  };
}

/**
 * @returns {Array<{ id: string, label: string }>}
 */
function listCreatableGoogleFileTypes() {
  return [
    { id: 'document', label: 'Google Документ' },
    { id: 'spreadsheet', label: 'Google Таблица' },
    { id: 'presentation', label: 'Google Презентация' },
    { id: 'form', label: 'Google Форма' },
    { id: 'drawing', label: 'Google Рисунок' },
    { id: 'vid', label: 'Google Vids' },
    { id: 'site', label: 'Google Сайт' },
    { id: 'map', label: 'Google Мои карты' },
    { id: 'jam', label: 'Google Jamboard' }
  ];
}

/**
 * @param {string} fileType
 * @returns {(string|null)}
 */
function resolveGoogleFileMimeType_(fileType) {
  var map = {
    document: 'application/vnd.google-apps.document',
    spreadsheet: 'application/vnd.google-apps.spreadsheet',
    presentation: 'application/vnd.google-apps.presentation',
    form: 'application/vnd.google-apps.form',
    drawing: 'application/vnd.google-apps.drawing',
    vid: 'application/vnd.google-apps.vid',
    site: 'application/vnd.google-apps.site',
    map: 'application/vnd.google-apps.map',
    jam: 'application/vnd.google-apps.jam'
  };
  return map[fileType] || null;
}

/**
 * @param {string} name
 * @param {string} mimeType
 * @param {string} parentFolderId
 * @returns {GoogleAppsScript.Drive.File}
 */
function createGoogleNativeFileInFolder_(name, mimeType, parentFolderId) {
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?supportsAllDrives=true', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify({
      name: name,
      mimeType: mimeType,
      parents: [parentFolderId]
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw catalogError_(
      'DRIVE_CREATE_FAILED',
      'Не удалось создать файл на Drive (' + code + '): ' + body
    );
  }

  var parsed = JSON.parse(body);
  if (!parsed.id) {
    throw catalogError_('DRIVE_CREATE_FAILED', 'Drive API не вернул id файла.');
  }

  return DriveApp.getFileById(parsed.id);
}

/**
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} controllerEmail
 * @param {string} creatorEmail
 */
function transferDriveFileToController_(file, controllerEmail, creatorEmail) {
  if (isDriveFileOwnedByEmail_(file, controllerEmail)) {
    ensureDriveEditor_(file, creatorEmail);
    return;
  }

  try {
    ensureDriveEditor_(file, controllerEmail);
    file.setOwner(controllerEmail);
  } catch (e) {
    throw catalogError_(
      'OWNERSHIP_TRANSFER_FAILED',
      'Не удалось передать файл владельцу каталога. ' +
        'Обычно это работает в Google Workspace одного домена. Детали: ' +
        (e.message || String(e))
    );
  }

  ensureDriveEditor_(file, creatorEmail);
}

/**
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} email
 */
function ensureDriveEditor_(file, email) {
  if (!email) {
    return;
  }
  try {
    file.addEditor(email);
  } catch (e) {
    // already editor / domain policy
  }
}

/**
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {string} email
 * @param {'none'|'reader'|'commenter'|'editor'} permissionLevel
 */
function appendExplicitUserAclRow_(objectType, objectId, email, permissionLevel) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    return;
  }
  sheet.appendRow([
    Utilities.getUuid(),
    objectType,
    objectId,
    'user',
    email,
    permissionLevel
  ]);
}
