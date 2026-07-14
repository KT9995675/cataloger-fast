/**
 * §9.1, §9.2, §9.6 — импорт одного файла с Google Drive в каталог.
 * Запись в `Files` создаётся сразу; копия/перемещение на Drive — синхронно (v1).
 *
 * @param {{
 *   targetFolderId: string,
 *   driveFileUrl: string,
 *   mode?: 'copy'|'move'
 * }} input
 * @returns {{
 *   ok: true,
 *   catalogId: string,
 *   displayName: string,
 *   fileId: string,
 *   mode: 'copy'|'move'
 * }}
 */
function importDriveFile(input) {
  assertCatalogReady_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var driveFileUrl = String(input.driveFileUrl || '').trim();
  var mode = String(input.mode || 'copy').trim().toLowerCase();

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (!driveFileUrl) {
    throw catalogError_('INVALID_INPUT', 'driveFileUrl is required.');
  }
  if (mode !== 'copy' && mode !== 'move') {
    throw catalogError_('INVALID_INPUT', 'mode must be copy or move.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var engine = createAclEngine_();
  if (!engine.foldersById[targetFolderId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Target folder not found: ' + targetFolderId);
  }
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, targetFolderId);

  var sourceFileId = parseDriveFileId_(driveFileUrl);
  var sourceFile;
  try {
    sourceFile = DriveApp.getFileById(sourceFileId);
  } catch (e) {
    throw catalogError_('INVALID_FILE', 'Drive file not found or not accessible.');
  }

  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var ownedByController = isDriveFileOwnedByEmail_(sourceFile, controllerEmail);
  if (!ownedByController && mode === 'move') {
    throw catalogError_('COPY_ONLY', 'Files not owned by the catalog controller can only be copied.');
  }

  var catalogRootFolderId = getCatalogRootFolderId_();
  var catalogRootFolder = DriveApp.getFolderById(catalogRootFolderId);
  var catalogFile = placeFileInCatalogRoot_(sourceFile, catalogRootFolder, mode);
  ensureUsersFromDriveFile_(sourceFile, userEmail);

  var catalogId = Utilities.getUuid();
  appendCatalogFileRow_({
    catalogId: catalogId,
    folderId: targetFolderId,
    fileId: catalogFile.getId(),
    displayName: sourceFile.getName(),
    sizeBytes: catalogFile.getSize(),
    driveModifiedAt: catalogFile.getLastUpdated(),
    sourceFileId: mode === 'copy' ? sourceFileId : '',
    mimeType: getDriveFileMimeType_(catalogFile)
  });

  return {
    ok: true,
    catalogId: catalogId,
    displayName: sourceFile.getName(),
    fileId: catalogFile.getId(),
    mode: mode
  };
}

/**
 * @returns {string}
 */
function getCatalogRootFolderId_() {
  var id = PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_ROOT_FOLDER_ID_);
  if (!id) {
    throw catalogError_('CATALOG_NOT_CONFIGURED', 'CATALOG_ROOT_FOLDER_ID is missing.');
  }
  return id;
}

/**
 * @param {string} urlOrId
 * @returns {string}
 */
function parseDriveFileId_(urlOrId) {
  var value = String(urlOrId || '').trim();
  var fileMatch = value.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) {
    return fileMatch[1];
  }
  var gdocMatch = value.match(/\/(?:document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]+)/);
  if (gdocMatch) {
    return gdocMatch[1];
  }
  var idMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch) {
    return idMatch[1];
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) {
    return value;
  }
  throw catalogError_('INVALID_FILE_URL', 'Cannot parse Drive file id from URL.');
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} email
 * @returns {boolean}
 */
function isDriveFileOwnedByEmail_(driveFile, email) {
  if (!email) {
    return false;
  }
  try {
    var owner = driveFile.getOwner();
    if (!owner) {
      return false;
    }
    var ownerEmail = owner.getEmail();
    return !!ownerEmail && ownerEmail.toLowerCase() === email.toLowerCase();
  } catch (e) {
    return false;
  }
}

/**
 * @param {GoogleAppsScript.Drive.File} sourceFile
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @param {'copy'|'move'} mode
 * @returns {GoogleAppsScript.Drive.File}
 */
function placeFileInCatalogRoot_(sourceFile, catalogRootFolder, mode) {
  if (mode === 'move') {
    sourceFile.moveTo(catalogRootFolder);
    return sourceFile;
  }
  return sourceFile.makeCopy(sourceFile.getName(), catalogRootFolder);
}

/**
 * §9.4 — добавляет пользователей с доступом к исходному файлу в `Users`.
 *
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} addedBy
 */
function ensureUsersFromDriveFile_(driveFile, addedBy) {
  var emails = collectDriveFileParticipantEmails_(driveFile);
  if (!emails.length) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    return;
  }

  var values = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < values.length; i++) {
    var email = String(values[i][0] || '').trim().toLowerCase();
    if (email) {
      existing[email] = true;
    }
  }

  var now = new Date();
  emails.forEach(function (email) {
    var key = email.toLowerCase();
    if (existing[key]) {
      return;
    }
    sheet.appendRow([email, 'user', now, addedBy || '']);
    existing[key] = true;
  });
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @returns {string[]}
 */
function collectDriveFileParticipantEmails_(driveFile) {
  var emails = {};

  function rememberUsers(users) {
    if (!users) {
      return;
    }

    if (typeof users.hasNext === 'function') {
      while (users.hasNext()) {
        rememberUser_(users.next());
      }
      return;
    }

    if (Array.isArray(users)) {
      for (var i = 0; i < users.length; i++) {
        rememberUser_(users[i]);
      }
    }
  }

  function rememberUser_(user) {
    if (!user) {
      return;
    }
    var email = user.getEmail();
    if (email) {
      emails[email.toLowerCase()] = email;
    }
  }

  rememberUsers(driveFile.getEditors());
  rememberUsers(driveFile.getViewers());
  try {
    rememberUsers(driveFile.getCommenters());
  } catch (e) {
    // commenters may be unavailable for some mime types
  }

  return Object.keys(emails).map(function (key) {
    return emails[key];
  });
}

/**
 * @param {{
 *   catalogId: string,
 *   folderId: string,
 *   fileId: string,
 *   displayName: string,
 *   sizeBytes: number,
 *   driveModifiedAt: Date,
 *   sourceFileId: string,
 *   mimeType?: string,
 *   approved?: boolean,
 *   approvedBy?: string,
 *   approvedAt?: (Date|string),
 *   status?: string,
 *   lastError?: string
 * }} row
 */
function appendCatalogFileRow_(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }

  ensureCatalogSchemaUpToDate_();

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  while (headers.length && !headers[headers.length - 1]) {
    headers.pop();
  }
  if (!headers.length) {
    throw catalogError_('SCHEMA_MISMATCH', 'Files sheet has no headers.');
  }

  var byHeader = {
    catalog_id: row.catalogId,
    folder_id: row.folderId,
    file_id: row.fileId,
    display_name: row.displayName,
    size_bytes: row.sizeBytes,
    drive_modified_at: row.driveModifiedAt,
    approved: row.approved === true,
    approved_by: row.approvedBy || '',
    approved_at: row.approvedAt || '',
    status: row.status || 'ready',
    last_error: row.lastError || '',
    source_file_id: row.sourceFileId || '',
    mime_type: row.mimeType || ''
  };

  var line = [];
  for (var c = 0; c < headers.length; c++) {
    var key = headers[c];
    line.push(Object.prototype.hasOwnProperty.call(byHeader, key) ? byHeader[key] : '');
  }
  sheet.appendRow(line);
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @returns {string}
 */
function getDriveFileMimeType_(driveFile) {
  try {
    return String(driveFile.getMimeType() || '');
  } catch (e) {
    return '';
  }
}
