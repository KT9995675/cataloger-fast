/**
 * §13.1 — переименование файла или папки (Shift+F6).
 * Имя меняется в каталоге и на Drive (для файлов).
 *
 * @param {{
 *   kind: 'folder'|'file',
 *   id: string,
 *   newName: string
 * }} input
 * @returns {{ ok: true, kind: 'folder'|'file', id: string, name: string }}
 */
function renameCatalogItem(input) {
  assertCatalogReady_();

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

  var engine = createAclEngine_();
  var virtualRootFolderId = getVirtualRootFolderId_();

  if (kind === 'folder') {
    return renameCatalogFolder_(engine, userEmail, loginRole, id, newName, virtualRootFolderId);
  }
  return renameCatalogFile_(engine, userEmail, loginRole, id, newName);
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} folderId
 * @param {string} newName
 * @param {string} virtualRootFolderId
 * @returns {{ ok: true, kind: 'folder', id: string, name: string }}
 */
function renameCatalogFolder_(engine, userEmail, loginRole, folderId, newName, virtualRootFolderId) {
  var folder = engine.foldersById[folderId];
  if (!folder) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + folderId);
  }
  if (folderId === virtualRootFolderId) {
    throw catalogError_('NOT_ALLOWED', 'Нельзя переименовать корневую папку каталога.');
  }
  if (folderId === '__TRASH__' || parseBoolean_(folder.is_system)) {
    throw catalogError_('NOT_ALLOWED', 'Нельзя переименовать системную папку.');
  }

  assertEditorOnFolderForMove_(engine, userEmail, loginRole, folderId);
  if (folder.parent_folder_id) {
    assertEditorOnFolderForMove_(engine, userEmail, loginRole, String(folder.parent_folder_id));
  }

  if (String(folder.name || '') === newName) {
    return { ok: true, kind: 'folder', id: folderId, name: newName };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var folderIdCol = headers.indexOf('folder_id');
  var nameCol = headers.indexOf('name');
  if (folderIdCol < 0 || nameCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Tree sheet headers are invalid.');
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][folderIdCol]) === folderId) {
      sheet.getRange(i + 1, nameCol + 1).setValue(newName);
      return { ok: true, kind: 'folder', id: folderId, name: newName };
    }
  }

  throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + folderId);
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

  var driveFileId = String(file.file_id || '').trim();
  if (!driveFileId) {
    throw catalogError_('INVALID_STATE', 'File has no Drive id.');
  }

  var driveFile;
  try {
    driveFile = DriveApp.getFileById(driveFileId);
  } catch (e) {
    throw catalogError_('DRIVE_ERROR', 'Drive file not found or not accessible.');
  }

  if (String(file.display_name || '') !== newName) {
    driveFile.setName(newName);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var catalogCol = headers.indexOf('catalog_id');
  var nameCol = headers.indexOf('display_name');
  var modifiedCol = headers.indexOf('drive_modified_at');
  if (catalogCol < 0 || nameCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Files sheet headers are invalid.');
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][catalogCol]) === catalogId) {
      sheet.getRange(i + 1, nameCol + 1).setValue(newName);
      if (modifiedCol >= 0) {
        try {
          sheet.getRange(i + 1, modifiedCol + 1).setValue(driveFile.getLastUpdated());
        } catch (e2) {
          // optional column update
        }
      }
      return { ok: true, kind: 'file', id: catalogId, name: newName };
    }
  }

  throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
}
