/**
 * §9.13 — импорт с локального диска (браузер → Drive + Tree/Files).
 * Только загрузка; структура папок — виртуальная в Tree.
 */

/** @const {number} Порог предупреждения на клиенте (байт); сервер не режет. */
var LOCAL_IMPORT_WARN_BYTES_ = 10 * 1024 * 1024;

/**
 * Создаёт виртуальные папки по относительным путям под targetFolderId.
 *
 * @param {{
 *   targetFolderId: string,
 *   folderPaths: string[]
 * }} input
 * @returns {{
 *   ok: true,
 *   pathToFolderId: Object.<string, string>,
 *   folderCount: number
 * }}
 */
function prepareLocalImportTree(input) {
  assertCatalogReady_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var folderPaths = Array.isArray(input.folderPaths) ? input.folderPaths : [];

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
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

  var normalized = normalizeLocalFolderPaths_(folderPaths);
  normalized.sort(function (a, b) {
    return a.split('/').length - b.split('/').length || (a < b ? -1 : a > b ? 1 : 0);
  });

  var treeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!treeSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  var pathToFolderId = {};
  var now = new Date();
  var created = 0;

  normalized.forEach(function (path) {
    if (pathToFolderId[path]) {
      return;
    }
    var parts = path.split('/');
    var parentId = targetFolderId;
    var accum = '';
    for (var i = 0; i < parts.length; i++) {
      accum = accum ? accum + '/' + parts[i] : parts[i];
      if (pathToFolderId[accum]) {
        parentId = pathToFolderId[accum];
        continue;
      }
      var folderId = Utilities.getUuid();
      treeSheet.appendRow([folderId, parentId, parts[i], now, false]);
      pathToFolderId[accum] = folderId;
      parentId = folderId;
      created++;
    }
  });

  return {
    ok: true,
    pathToFolderId: pathToFolderId,
    folderCount: created
  };
}

/**
 * Загружает один локальный файл в плоскую папку каталога на Drive + Files.
 *
 * @param {{
 *   parentFolderId: string,
 *   fileName: string,
 *   mimeType?: string,
 *   base64Data: string,
 *   sizeBytes?: number
 * }} input
 * @returns {{
 *   ok: true,
 *   catalogId: string,
 *   displayName: string,
 *   fileId: string,
 *   sizeBytes: number
 * }}
 */
function importLocalFile(input) {
  assertCatalogReady_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var fileName = String(input.fileName || '').trim();
  var mimeType = String(input.mimeType || 'application/octet-stream').trim();
  var base64Data = String(input.base64Data || '').trim();
  var sizeBytes = parseNumber_(input.sizeBytes);

  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!fileName) {
    throw catalogError_('INVALID_INPUT', 'fileName is required.');
  }
  if (!base64Data) {
    throw catalogError_('INVALID_INPUT', 'base64Data is required.');
  }

  var comma = base64Data.indexOf(',');
  if (base64Data.indexOf('base64') >= 0 && comma >= 0) {
    base64Data = base64Data.substring(comma + 1);
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var engine = createAclEngine_();
  if (!engine.foldersById[parentFolderId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Parent folder not found: ' + parentFolderId);
  }
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, parentFolderId);

  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Data);
  } catch (e) {
    throw catalogError_('INVALID_INPUT', 'Не удалось разобрать содержимое файла.');
  }

  if (!sizeBytes || sizeBytes < 0) {
    sizeBytes = bytes.length;
  }

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
  var driveFile = catalogRootFolder.createFile(blob);

  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  if (controllerEmail) {
    transferDriveFileToController_(driveFile, controllerEmail, userEmail);
  }

  var catalogId = Utilities.getUuid();
  var resolvedMime = '';
  try {
    resolvedMime = String(driveFile.getMimeType() || mimeType || '');
  } catch (eMime) {
    resolvedMime = mimeType || '';
  }

  appendCatalogFileRow_({
    catalogId: catalogId,
    folderId: parentFolderId,
    fileId: driveFile.getId(),
    displayName: fileName,
    sizeBytes: sizeBytes,
    driveModifiedAt: driveFile.getLastUpdated(),
    sourceFileId: '',
    mimeType: resolvedMime
  });

  return {
    ok: true,
    catalogId: catalogId,
    displayName: fileName,
    fileId: driveFile.getId(),
    sizeBytes: sizeBytes
  };
}

/**
 * @param {string[]} folderPaths
 * @returns {string[]}
 */
function normalizeLocalFolderPaths_(folderPaths) {
  var seen = {};
  var out = [];
  for (var i = 0; i < folderPaths.length; i++) {
    var raw = String(folderPaths[i] || '')
      .replace(/\\/g, '/')
      .replace(/^\/+|\/+$/g, '');
    if (!raw) {
      continue;
    }
    var parts = raw.split('/').filter(function (p) {
      return p && p !== '.' && p !== '..';
    });
    if (!parts.length) {
      continue;
    }
    var path = parts.join('/');
    if (seen[path]) {
      continue;
    }
    seen[path] = true;
    out.push(path);
  }
  return out;
}
