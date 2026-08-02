/** @const {string} */
var TRASH_FOLDER_ID_EMPTY_ = '__TRASH__';

/**
 * §13.4 — очистить корзину через Jobs (`empty_trash`):
 * файлы → корзина Google Drive; строки Tree/Files/ACL — в воркере чанками.
 * Только роль входа **Управляющий**.
 *
 * @returns {{
 *   ok: true,
 *   queued: boolean,
 *   jobId: (string|undefined),
 *   fileCount: number,
 *   folderCount: number
 * }}
 */
function emptyCatalogTrash() {
  assertCatalogReady_();

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var engine = createAclEngine_();
  if (!engine.foldersById[TRASH_FOLDER_ID_EMPTY_]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Папка корзины не найдена.');
  }

  var items = collectEmptyTrashFileItemsFromSheet_();
  var folderIdsToDelete = collectEmptyTrashFolderIdsFromSheet_();

  if (!items.length && !folderIdsToDelete.length) {
    return {
      ok: true,
      queued: false,
      fileCount: 0,
      folderCount: 0
    };
  }

  var payload = {
    items: items,
    folderIds: folderIdsToDelete,
    foldersDone: false,
    driveErrors: 0
  };

  setCatalogJobsPaused_(false);
  var jobId = enqueueCatalogJob_('empty_trash', payload, userEmail, '');
  kickCatalogJobsProcessing_();

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    fileCount: items.length,
    folderCount: folderIdsToDelete.length
  };
}

/* <!-- OLD: emptyCatalogTrash — синхронный wipe Drive+Sheets без Jobs / прогресса N/M --> */

/**
 * Содержимое корзины для confirm в UI (без удаления).
 *
 * @returns {{ fileCount: number, folderCount: number }}
 */
function getCatalogTrashStats() {
  assertCatalogReady_();
  var engine = createAclEngine_();
  if (!engine.foldersById[TRASH_FOLDER_ID_EMPTY_]) {
    return { fileCount: 0, folderCount: 0 };
  }
  var trashFolderIds = collectTrashSubtreeFolderIds_(engine);
  var fileCount = 0;
  Object.keys(engine.filesByCatalogId).forEach(function (catalogId) {
    var folderId = String(engine.filesByCatalogId[catalogId].folder_id || '');
    if (trashFolderIds[folderId]) {
      fileCount += 1;
    }
  });
  var folderCount = Object.keys(trashFolderIds).filter(function (id) {
    return id !== TRASH_FOLDER_ID_EMPTY_;
  }).length;
  return { fileCount: fileCount, folderCount: folderCount };
}

/**
 * @param {string} userEmail
 */
function assertIsCatalogController_(userEmail) {
  var loginRole = getLoginRoleForUser_(userEmail);
  if (loginRole === 'controller') {
    return;
  }
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  if (
    controllerEmail &&
    String(userEmail).toLowerCase() === String(controllerEmail).toLowerCase()
  ) {
    return;
  }
  if (isSpreadsheetOwnerEmail_(userEmail)) {
    return;
  }
  throw catalogError_('NOT_ALLOWED', 'Действие доступно только Управляющему.');
}

/**
 * @param {Object} engine
 * @returns {Object.<string, boolean>}
 */
function collectTrashSubtreeFolderIds_(engine) {
  var folderIds = {};
  folderIds[TRASH_FOLDER_ID_EMPTY_] = true;
  var queue = [TRASH_FOLDER_ID_EMPTY_];
  while (queue.length) {
    var currentId = queue.shift();
    Object.keys(engine.foldersById).forEach(function (childId) {
      var child = engine.foldersById[childId];
      if (String(child.parent_folder_id || '') === currentId && !folderIds[childId]) {
        folderIds[childId] = true;
        queue.push(childId);
      }
    });
  }
  return folderIds;
}

/**
 * Перенос файла в корзину Google Drive (`trashed: true`).
 * Безвозвратный DELETE не используем — Google обычно стирает из своей корзины ~через 30 дней.
 *
 * @param {string} driveFileId
 */
function moveDriveFileToTrash_(driveFileId) {
  var result = moveDriveFilesToTrashBatch_([driveFileId]);
  if (result.errors > 0) {
    throw catalogError_(
      'DRIVE_TRASH_FAILED',
      'Не удалось переместить файл в корзину Drive'
    );
  }
}

/**
 * Пакетный перенос в корзину Drive (`trashed: true`) через DriveApp.
 * Без UrlFetch — иначе дневная квота urlfetch валит clear/empty_trash (B1).
 *
 * @param {string[]} driveFileIds
 * @returns {{
 *   ok: true,
 *   count: number,
 *   errors: number,
 *   okIds: string[],
 *   failedIds: string[]
 * }}
 */
function moveDriveFilesToTrashBatch_(driveFileIds) {
  var ids = [];
  var seen = {};
  (driveFileIds || []).forEach(function (id) {
    id = String(id || '').trim();
    if (!id || seen[id]) {
      return;
    }
    seen[id] = true;
    ids.push(id);
  });
  if (!ids.length) {
    return { ok: true, count: 0, errors: 0, okIds: [], failedIds: [] };
  }

  var errors = 0;
  var okIds = [];
  var failedIds = [];
  for (var i = 0; i < ids.length; i++) {
    var fileId = ids[i];
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      okIds.push(fileId);
    } catch (eTrash) {
      var msg = String((eTrash && eTrash.message) || eTrash || '');
      // Уже в корзине / нет файла — успех (как HTTP 404 раньше).
      if (
        /not found|could be found|не найден|trashed|корзин|does not exist/i.test(
          msg
        )
      ) {
        okIds.push(fileId);
      } else {
        errors += 1;
        failedIds.push(fileId);
      }
    }
  }

  return {
    ok: true,
    count: ids.length,
    errors: errors,
    okIds: okIds,
    failedIds: failedIds
  };
}

/* OLD: UrlFetchApp.fetchAll patch trashed:true — дневная квота urlfetch валила clear/empty_trash. */

/* <!-- OLD: permanentlyDeleteDriveFile_ — Drive API DELETE (безвозвратно); заменено на moveDriveFileToTrash_ --> */

/**
 * @param {string[]} catalogIds
 */
function removeCatalogFileRows_(catalogIds) {
  if (!catalogIds || !catalogIds.length) {
    return;
  }
  var remove = {};
  catalogIds.forEach(function (id) {
    remove[String(id)] = true;
  });
  rewriteSheetRemovingRows_('Files', 'catalog_id', remove);
}

/**
 * @param {string[]} folderIds
 */
function removeCatalogTreeRows_(folderIds) {
  if (!folderIds || !folderIds.length) {
    return;
  }
  var remove = {};
  folderIds.forEach(function (id) {
    remove[String(id)] = true;
  });
  rewriteSheetRemovingRows_('Tree', 'folder_id', remove);
}

/**
 * @param {string[]} fileCatalogIds
 * @param {string[]} folderIds
 */
function removeAclForTrashObjects_(fileCatalogIds, folderIds) {
  var targetKeys = {};
  (fileCatalogIds || []).forEach(function (id) {
    targetKeys['file:' + id] = true;
  });
  (folderIds || []).forEach(function (id) {
    targetKeys['folder:' + id] = true;
  });
  if (!Object.keys(targetKeys).length) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
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
  var objectTypeCol = headers.indexOf('object_type');
  var objectIdCol = headers.indexOf('object_id');
  if (objectTypeCol < 0 || objectIdCol < 0) {
    return;
  }

  var kept = [headers];
  for (var i = 1; i < values.length; i++) {
    var key =
      String(values[i][objectTypeCol] || '').trim() +
      ':' +
      String(values[i][objectIdCol] || '').trim();
    if (!targetKeys[key]) {
      kept.push(values[i]);
    }
  }

  sheet.clearContents();
  if (kept.length) {
    sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
  }
}

/**
 * @param {string} sheetName
 * @param {string} idHeader
 * @param {Object.<string, boolean>} removeIds
 */
function rewriteSheetRemovingRows_(sheetName, idHeader, removeIds) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ' + sheetName);
  }
  var values = sheet.getDataRange().getValues();
  if (!values.length) {
    return;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf(idHeader);
  if (idCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', sheetName + ' missing column ' + idHeader);
  }

  var kept = [headers];
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][idCol] || '').trim();
    if (!removeIds[id]) {
      kept.push(values[i]);
    }
  }

  sheet.clearContents();
  if (kept.length) {
    sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
  }
}
