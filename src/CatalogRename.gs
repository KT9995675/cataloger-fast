/**
 * §13.1 — переименование файла или папки (Shift+F6).
 * Папка: только лист Tree (быстрый путь). Файл: Tree/Files + имя на Drive.
 *
 * @param {{
 *   kind: 'folder'|'file',
 *   id: string,
 *   newName: string
 * }} input
 * @returns {{ ok: true, kind: 'folder'|'file', id: string, name: string }}
 */
function renameCatalogItem(input) {
  assertCatalogReadyLight_();

  input = input || {};
  var kind = String(input.kind || '').trim();
  var id = String(input.id || '').trim();
  var newName = String(input.newName || '').trim();

  if (kind !== 'folder' && kind !== 'file') {
    throw catalogError_('INVALID_INPUT', 'kind must be folder or file.');
  }
  if (!id) {
    throw catalogError_('INVALID_INPUT', 'id is required.');
  }
  if (!newName) {
    throw catalogError_('INVALID_INPUT', 'Новое имя не может быть пустым.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  if (kind === 'folder') {
    return renameCatalogFolderFast_(userEmail, loginRole, id, newName);
  }

  var engine = createAclEngine_();
  return renameCatalogFile_(engine, userEmail, loginRole, id, newName);
}

/**
 * Быстрый rename папки: без чтения Files; ACL-движок только Tree+ACL+Users+Groups.
 *
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} folderId
 * @param {string} newName
 * @returns {{ ok: true, kind: 'folder', id: string, name: string }}
 */
function renameCatalogFolderFast_(userEmail, loginRole, folderId, newName) {
  var virtualRootFolderId = getVirtualRootFolderId_();
  if (folderId === virtualRootFolderId) {
    throw catalogError_('NOT_ALLOWED', 'Нельзя переименовать корневую папку каталога.');
  }
  if (folderId === '__TRASH__') {
    throw catalogError_('NOT_ALLOWED', 'Нельзя переименовать системную папку.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + folderId);
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var folderIdCol = headers.indexOf('folder_id');
  var parentCol = headers.indexOf('parent_folder_id');
  var nameCol = headers.indexOf('name');
  var systemCol = headers.indexOf('is_system');
  if (folderIdCol < 0 || nameCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Tree sheet headers are invalid.');
  }

  var rowIndex = -1;
  var parentFolderId = '';
  var isSystem = false;
  var currentName = '';
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][folderIdCol]) === folderId) {
      rowIndex = i;
      parentFolderId = parentCol >= 0 ? String(values[i][parentCol] || '') : '';
      isSystem = systemCol >= 0 && parseBoolean_(values[i][systemCol]);
      currentName = String(values[i][nameCol] || '');
      break;
    }
  }
  if (rowIndex < 0) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + folderId);
  }
  if (isSystem) {
    throw catalogError_('NOT_ALLOWED', 'Нельзя переименовать системную папку.');
  }

  // Права: лёгкий движок без Files
  var engine = buildAclEngineFromRows_(
    readSheetRecords_('Tree'),
    [],
    readSheetRecords_('ACL'),
    readSheetRecords_('GroupMembers'),
    readSheetRecords_('Users'),
    readSheetRecords_('Groups')
  );
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, folderId);
  if (parentFolderId) {
    assertEditorOnFolderForMove_(engine, userEmail, loginRole, parentFolderId);
  }

  if (currentName === newName) {
    return { ok: true, kind: 'folder', id: folderId, name: newName };
  }

  sheet.getRange(rowIndex + 1, nameCol + 1).setValue(newName);
  return { ok: true, kind: 'folder', id: folderId, name: newName };
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} catalogId
 * @param {string} newName
 * @returns {{ ok: true, kind: 'file', id: string, name: string }}
 */
function renameCatalogFile_(engine, userEmail, loginRole, catalogId, newName) {
  var file = engine.filesByCatalogId[catalogId];
  if (!file) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
  }

  var parentFolderId = String(file.folder_id || '');
  if (!parentFolderId) {
    throw catalogError_('INVALID_STATE', 'File has no parent folder.');
  }
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, parentFolderId);

  if (String(file.display_name || '') === newName) {
    return { ok: true, kind: 'file', id: catalogId, name: newName };
  }

  var fileId = String(file.file_id || '').trim();
  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setName(newName);
    } catch (eDrive) {
      throw catalogError_(
        'DRIVE_RENAME_FAILED',
        (eDrive && eDrive.message) || 'Не удалось переименовать файл на Drive.'
      );
    }
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('catalog_id');
  var nameCol = headers.indexOf('display_name');
  if (idCol < 0 || nameCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Files sheet headers are invalid.');
  }
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === catalogId) {
      sheet.getRange(i + 1, nameCol + 1).setValue(newName);
      return { ok: true, kind: 'file', id: catalogId, name: newName };
    }
  }
  throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
}
