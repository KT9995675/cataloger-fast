/**
 * §4.2, §4.3, §12 — открытие файла каталога с проверкой прав и временным доступом на Drive.
 *
 * @param {string} catalogId `Files.catalog_id`
 * @returns {{
 *   ok: true,
 *   url: string,
 *   displayName: string,
 *   mimeType: string,
 *   permissionLevel: ('reader'|'commenter'|'editor')
 * }|{
 *   ok: false,
 *   code: 'NO_ACCESS',
 *   message: string,
 *   editors: string[]
 * }}
 */
function openCatalogFile(catalogId) {
  assertCatalogReady_();

  catalogId = String(catalogId || '').trim();
  if (!catalogId) {
    throw catalogError_('INVALID_INPUT', 'catalogId is required.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertActiveCatalogUser_(userEmail);

  var engine = createAclEngine_();
  var file = engine.filesByCatalogId[catalogId];
  if (!file) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
  }

  var driveShortcutId = String(file.shortcut_of_drive_file_id || '').trim();
  if (driveShortcutId) {
    return openExternalFileShortcut_(catalogId, file, driveShortcutId);
  }

  var blueShortcut = String(file.shortcut_of_catalog_id || '').trim();
  if (blueShortcut) {
    catalogId = resolveFileShortcutTargetCatalogId_(engine.filesByCatalogId, catalogId);
    file = engine.filesByCatalogId[catalogId];
    if (!file) {
      throw catalogError_('FILE_NOT_FOUND', 'Файл-цель ярлыка не найден.');
    }
  }

  var driveFileId = String(file.file_id || '').trim();
  if (!driveFileId) {
    throw catalogError_('FILE_NOT_READY', 'Файл ещё не загружен на Drive.');
  }

  var status = String(file.status || 'ready').trim().toLowerCase();
  if (status === 'pending') {
    throw catalogError_('FILE_NOT_READY', 'Файл ещё загружается на Drive.');
  }
  if (status === 'failed') {
    throw catalogError_('FILE_IMPORT_FAILED', 'Импорт файла завершился с ошибкой.');
  }

  var controllerEmail = PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var permission = 'none';
  if (controllerEmail && controllerEmail.toLowerCase() === userEmail.toLowerCase()) {
    permission = 'editor';
  } else {
    permission = getEffectivePermissionForUserFromEngine_(engine, 'file', catalogId, userEmail);
  }

  if (!canOpenWithPermission_(permission)) {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Запросите права у редактора данной папки.',
      editors: getNoAccessEditorsFromEngine_(engine, catalogId)
    };
  }

  var driveMeta = fetchDriveFileOpenMeta_(driveFileId);
  var mimeType =
    driveMeta.mimeType ||
    String(file.mime_type || '') ||
    getDriveFileMimeTypeSafeById_(driveFileId);
  var displayName = String(file.display_name || driveMeta.name || driveFileId);
  var openUrl = buildCatalogFileOpenUrl_(driveFileId, mimeType, driveMeta.webViewLink);

  try {
    var driveFile = DriveApp.getFileById(driveFileId);
    var isDriveOwner = isDriveFileOwner_(driveFile, userEmail);
    if (!isDriveOwner) {
      try {
        grantTemporaryDriveAccess_(driveFile, userEmail, permission);
      } catch (grantErr) {
        // уже есть доступ / политика — всё равно открываем по URL
      }
    }
    var modifiedAt = null;
    try {
      modifiedAt = driveFile.getLastUpdated();
    } catch (e2) {
      modifiedAt = new Date();
    }
    updateFileDriveMeta_(catalogId, modifiedAt, mimeType);
  } catch (e) {
    if (mimeType) {
      updateFileDriveMeta_(catalogId, new Date(), mimeType);
    }
  }

  return {
    ok: true,
    url: openUrl,
    displayName: displayName,
    mimeType: mimeType || '',
    permissionLevel: permission
  };
}

/**
 * §22 — открытие красного ярлыка файла (Drive OAuth пользователя).
 *
 * @param {string} shortcutCatalogId
 * @param {Object} fileRow
 * @param {string} driveFileId
 * @returns {Object}
 */
function openExternalFileShortcut_(shortcutCatalogId, fileRow, driveFileId) {
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(driveFileId) +
      '?fields=' +
      encodeURIComponent('id,name,mimeType,webViewLink') +
      '&supportsAllDrives=true',
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );
  var code = response.getResponseCode();
  var body = {};
  try {
    body = JSON.parse(response.getContentText() || '{}');
  } catch (eParse) {
    body = {};
  }
  if (code === 404) {
    deleteCatalogFileShortcutRows_([shortcutCatalogId]);
    bumpCatalogRev_();
    return {
      ok: false,
      code: 'GONE',
      message: 'Больше нет такого файла, удаляю ярлык',
      deletedShortcutId: shortcutCatalogId,
      parentFolderId: fileRow.folder_id ? String(fileRow.folder_id) : null
    };
  }
  if (code === 401 || code === 403 || code < 200 || code >= 300) {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа',
      editors: []
    };
  }
  var mimeType = String(body.mimeType || fileRow.mime_type || '');
  return {
    ok: true,
    url: buildCatalogFileOpenUrl_(driveFileId, mimeType, body.webViewLink || ''),
    displayName: String(body.name || fileRow.display_name || driveFileId),
    mimeType: mimeType,
    permissionLevel: 'reader'
  };
}

/**
 * Канонический URL редактора Google-файла — надёжнее, чем DriveApp.getUrl() для Docs.
 *
 * @param {string} fileId
 * @param {string} mimeType
 * @param {string=} webViewLink
 * @returns {string}
 */
function buildCatalogFileOpenUrl_(fileId, mimeType, webViewLink) {
  var mime = String(mimeType || '').toLowerCase();
  var id = encodeURIComponent(String(fileId || '').trim());
  var byMime = {
    'application/vnd.google-apps.document': 'https://docs.google.com/document/d/' + id + '/edit',
    'application/vnd.google-apps.spreadsheet': 'https://docs.google.com/spreadsheets/d/' + id + '/edit',
    'application/vnd.google-apps.presentation': 'https://docs.google.com/presentation/d/' + id + '/edit',
    'application/vnd.google-apps.form': 'https://docs.google.com/forms/d/' + id + '/edit',
    'application/vnd.google-apps.drawing': 'https://docs.google.com/drawings/d/' + id + '/edit',
    'application/vnd.google-apps.site': 'https://sites.google.com/d/' + id + '/edit',
    'application/vnd.google-apps.jam': 'https://jamboard.google.com/d/' + id + '/edit',
    'application/vnd.google-apps.map': 'https://www.google.com/maps/d/edit?mid=' + id,
    'application/vnd.google-apps.vid': 'https://vids.google.com/document/d/' + id + '/edit'
  };
  if (byMime[mime]) {
    return byMime[mime];
  }
  if (webViewLink) {
    return String(webViewLink);
  }
  return 'https://drive.google.com/file/d/' + id + '/view';
}

/**
 * @param {string} fileId
 * @returns {{ mimeType: string, webViewLink: string, name: string }}
 */
function fetchDriveFileOpenMeta_(fileId) {
  var empty = { mimeType: '', webViewLink: '', name: '' };
  try {
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(fileId) +
        '?fields=mimeType,webViewLink,name&supportsAllDrives=true',
      {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return empty;
    }
    var parsed = JSON.parse(response.getContentText() || '{}');
    return {
      mimeType: String(parsed.mimeType || ''),
      webViewLink: String(parsed.webViewLink || ''),
      name: String(parsed.name || '')
    };
  } catch (e) {
    return empty;
  }
}

/**
 * @param {string} fileId
 * @returns {string}
 */
function getDriveFileMimeTypeSafeById_(fileId) {
  try {
    return getDriveFileMimeType_(DriveApp.getFileById(fileId));
  } catch (e) {
    return '';
  }
}

/**
 * @param {string} email
 */
function assertActiveCatalogUser_(email) {
  var users = readSheetRecords_('Users');
  var normalized = email.toLowerCase();
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].email || '').toLowerCase() === normalized) {
      return;
    }
  }
  throw catalogError_('NOT_IN_CATALOG', 'Ваш аккаунт не добавлен в каталог.');
}

/**
 * @param {('none'|'reader'|'commenter'|'editor')} permission
 * @returns {boolean}
 */
function canOpenWithPermission_(permission) {
  return permission === 'reader' || permission === 'commenter' || permission === 'editor';
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} email
 * @returns {boolean}
 */
function isDriveFileOwner_(driveFile, email) {
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
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} email
 * @param {('reader'|'commenter'|'editor')} permission
 */
function grantTemporaryDriveAccess_(driveFile, email, permission) {
  if (permission === 'editor') {
    driveFile.addEditor(email);
    return;
  }
  if (permission === 'commenter') {
    driveFile.addCommenter(email);
    return;
  }
  driveFile.addViewer(email);
}

/**
 * @param {string} catalogId
 * @param {Date} modifiedAt
 * @param {string=} mimeType
 */
function updateFileDriveMeta_(catalogId, modifiedAt, mimeType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    return;
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var catalogCol = headers.indexOf('catalog_id');
  var modifiedCol = headers.indexOf('drive_modified_at');
  var mimeCol = headers.indexOf('mime_type');
  if (catalogCol < 0 || modifiedCol < 0) {
    return;
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][catalogCol]) === catalogId) {
      sheet.getRange(i + 1, modifiedCol + 1).setValue(modifiedAt);
      if (mimeCol >= 0 && mimeType) {
        var currentMime = String(values[i][mimeCol] || '').trim();
        if (!currentMime || currentMime !== mimeType) {
          sheet.getRange(i + 1, mimeCol + 1).setValue(mimeType);
        }
      }
      return;
    }
  }
}

/** @deprecated use updateFileDriveMeta_ */
function updateFileDriveModifiedAt_(catalogId, modifiedAt) {
  updateFileDriveMeta_(catalogId, modifiedAt, '');
}
