/**
 * §24.3 — очистить каталог (Управляющий).
 * Файлы → корзина Drive (пакетно); затем wipe метаданных.
 */

/**
 * @param {{
 *   keepTree?: boolean,
 *   keepAcl?: boolean,
 *   keepUsers?: boolean,
 *   keepGroups?: boolean
 * }=} input
 * @returns {{
 *   ok: true,
 *   queued: boolean,
 *   jobId: (string|undefined),
 *   fileCount: number,
 *   fullReset: boolean
 * }}
 */
function clearCatalog(input) {
  assertCatalogReady_();
  input = input || {};

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var keep = {
    tree: !!input.keepTree,
    acl: !!input.keepAcl,
    users: !!input.keepUsers,
    groups: !!input.keepGroups
  };
  var fullReset = !keep.tree && !keep.acl && !keep.users && !keep.groups;

  var driveFileIds = collectAllCatalogDriveFileIdsForClear_();

  if (!driveFileIds.length) {
    applyClearCatalogWipe_(keep, fullReset);
    return {
      ok: true,
      queued: false,
      fileCount: 0,
      fullReset: fullReset,
      needsFirstLaunch: fullReset
    };
  }

  var jobId = enqueueCatalogJob_(
    'clear_catalog',
    {
      // Без списка id в payload (лимит ячейки / зависания). Воркер каждый раз
      // заново собирает оставшиеся файлы с Drive/Files.
      initialFileCount: driveFileIds.length,
      keep: keep,
      fullReset: fullReset,
      driveErrors: 0,
      wipeDone: false
    },
    userEmail,
    ''
  );
  kickCatalogJobsProcessing_();

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    fileCount: driveFileIds.length,
    fullReset: fullReset
  };
}

/**
 * Все file_id из Files + файлы прямо в плоской папке каталога (сироты).
 *
 * @returns {string[]}
 */
function collectAllCatalogDriveFileIdsForClear_() {
  var seen = {};
  try {
    var fileRows = readSheetRecords_('Files');
    fileRows.forEach(function (row) {
      var id = String(row.file_id || '').trim();
      if (id) {
        seen[id] = true;
      }
    });
  } catch (eFiles) {
    // ignore
  }

  try {
    var rootId = getCatalogRootFolderId_();
    var folder = DriveApp.getFolderById(rootId);
    var it = folder.getFiles();
    var guard = 0;
    while (it.hasNext() && guard < 20000) {
      guard += 1;
      seen[it.next().getId()] = true;
    }
  } catch (eList) {
    // нет доступа / нет папки — всё равно чистим по Files
  }

  return Object.keys(seen);
}

/**
 * @param {{ tree: boolean, acl: boolean, users: boolean, groups: boolean }} keep
 * @param {boolean} fullReset
 */
function applyClearCatalogWipe_(keep, fullReset) {
  if (fullReset) {
    clearCatalogDataForFirstLaunch_();
    return;
  }

  clearCatalogSheetDataRows_('Files');
  clearCatalogSheetDataRows_('Jobs');

  if (!keep.tree) {
    clearCatalogSheetDataRows_('Tree');
    recreateEmptyCatalogTreeAfterClear_();
  }
  if (!keep.acl) {
    clearCatalogSheetDataRows_('ACL');
  }
  if (!keep.users) {
    clearCatalogSheetDataRows_('Users');
    ensureControllerUserAfterPartialClear_();
  }
  if (!keep.groups) {
    clearCatalogSheetDataRows_('Groups');
    clearCatalogSheetDataRows_('GroupMembers');
  }

  bumpCatalogRev_();
}

/**
 * Очистить строки данных листа, заголовок по схеме оставить.
 *
 * @param {string} sheetName
 */
function clearCatalogSheetDataRows_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schema = getCatalogSheetSchema_();
  var headers = schema[sheetName];
  if (!headers) {
    return;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return;
  }
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), headers.length, 1);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
}

/**
 * После удаления Tree — снова корень + ## Корзина (каталог остаётся initialized).
 */
function recreateEmptyCatalogTreeAfterClear_() {
  var props = PropertiesService.getDocumentProperties();
  var virtualRootFolderId =
    props.getProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_) || Utilities.getUuid();
  props.setProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_, virtualRootFolderId);
  var rootName = 'Каталог';
  writeVirtualTreeBootstrap_(virtualRootFolderId, rootName, '__TRASH__', new Date());
}

/**
 * Если Users очищен — вернуть строку Управляющего.
 */
function ensureControllerUserAfterPartialClear_() {
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  if (!controllerEmail) {
    return;
  }
  writeControllerUser_(controllerEmail, new Date());
}

/**
 * Добить зависшую/текущую очистку: trash оставшихся + wipe.
 * Вызывается из воркера и из cancel.
 *
 * @param {Object=} payload
 * @returns {{ trashed: number, errors: number, wiped: boolean, needsFirstLaunch: boolean }}
 */
function finishClearCatalogNow_(payload) {
  payload = payload || {};
  var keep = payload.keep || {
    tree: false,
    acl: false,
    users: false,
    groups: false
  };
  var fullReset = !!payload.fullReset;

  var ids = collectAllCatalogDriveFileIdsForClear_();
  // Старый payload со списком items — добрать неdone.
  (payload.items || []).forEach(function (it) {
    if (it && !it.done && it.driveFileId) {
      ids.push(String(it.driveFileId));
    }
  });
  var seen = {};
  var unique = [];
  ids.forEach(function (id) {
    id = String(id || '').trim();
    if (!id || seen[id]) {
      return;
    }
    seen[id] = true;
    unique.push(id);
  });

  var errors = 0;
  var trashed = 0;
  if (unique.length) {
    var result = moveDriveFilesToTrashBatch_(unique);
    errors = Number(result.errors) || 0;
    trashed = Math.max(0, (Number(result.count) || 0) - errors);
  }

  applyClearCatalogWipe_(
    {
      tree: !!keep.tree,
      acl: !!keep.acl,
      users: !!keep.users,
      groups: !!keep.groups
    },
    fullReset
  );

  return {
    trashed: trashed,
    errors: errors,
    wiped: true,
    needsFirstLaunch: fullReset
  };
}

/**
 * Воркер Jobs `clear_catalog`.
 * Каждый тик: чанк DriveApp trash → убрать file_id из Files → прогресс;
 * id, которые Drive не даёт в корзину (права и т.п.) — в skip, иначе вечный хвост 900/N;
 * когда collect∖skip пуст — wipe.
 *
 * @param {Object} job
 * @returns {number}
 */
function processClearCatalogJobChunk_(job) {
  var payload = parseJobPayload_(job);
  var keep = payload.keep || {};
  var fullReset = !!payload.fullReset;
  var initialTotal =
    Number(payload.initialFileCount) ||
    ((payload.items && payload.items.length) || 0);
  var skipMap = payload.skipDriveIds || {};

  var remaining = collectAllCatalogDriveFileIdsForClear_();
  (payload.items || []).forEach(function (it) {
    if (it && !it.done && it.driveFileId) {
      remaining.push(String(it.driveFileId));
    }
  });
  var seen = {};
  remaining = remaining.filter(function (id) {
    id = String(id || '').trim();
    if (!id || seen[id] || skipMap[id]) {
      return false;
    }
    seen[id] = true;
    return true;
  });

  if (!initialTotal) {
    initialTotal = remaining.length;
    payload.initialFileCount = initialTotal;
  }

  if (remaining.length) {
    var chunkSize =
      typeof JOBS_CHUNK_SIZE_ === 'number' && JOBS_CHUNK_SIZE_ > 0
        ? JOBS_CHUNK_SIZE_
        : 50;
    var chunk = remaining.slice(0, chunkSize);
    var result = moveDriveFilesToTrashBatch_(chunk);
    payload.driveErrors =
      (Number(payload.driveErrors) || 0) + (Number(result.errors) || 0);

    (result.failedIds || []).forEach(function (id) {
      skipMap[String(id)] = true;
    });
    payload.skipDriveIds = skipMap;

    // Сжимаем Files — иначе collect снова вернёт те же id и wipe не наступит.
    var removeMap = {};
    chunk.forEach(function (id) {
      removeMap[String(id)] = true;
    });
    rewriteSheetRemovingRows_('Files', 'file_id', removeMap);

    (payload.items || []).forEach(function (it) {
      if (it && it.driveFileId && removeMap[String(it.driveFileId)]) {
        it.done = true;
      }
    });

    var still = collectAllCatalogDriveFileIdsForClear_().filter(function (id) {
      return !skipMap[String(id)];
    });
    var skipCount = Object.keys(skipMap).length;
    var doneApprox = Math.max(0, initialTotal - still.length);
    if (still.length) {
      var msg =
        'Очистка каталога: ' + doneApprox + '/' + initialTotal;
      if (skipCount) {
        msg += ' (пропуск: ' + skipCount + ')';
      }
      saveJobPayloadProgress_(
        job.job_id,
        payload,
        Math.min(90, (doneApprox / initialTotal) * 90),
        msg,
        false
      );
      return chunk.length;
    }
  }

  if (!payload.wipeDone) {
    applyClearCatalogWipe_(
      {
        tree: !!keep.tree,
        acl: !!keep.acl,
        users: !!keep.users,
        groups: !!keep.groups
      },
      fullReset
    );
    payload.wipeDone = true;
  }

  return 1;
}

/**
 * Ручное добивание зависшей очистки (Управляющий).
 *
 * @returns {{
 *   ok: true,
 *   recovered: boolean,
 *   trashed: number,
 *   needsFirstLaunch: boolean,
 *   message: string
 * }}
 */
function recoverStuckClearCatalog() {
  assertCatalogReady_();
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw catalogError_(
      'JOBS_BUSY',
      'Сейчас занят воркер. Повторите через несколько секунд.'
    );
  }

  try {
    var job = findClearCatalogJobAny_();
    var payload = job ? parseJobPayload_(job) : {};
    var result = finishClearCatalogNow_(payload);
    if (job) {
      try {
        cancelJobsByChainOrId_('', String(job.job_id));
      } catch (eCancel) {
        // wipe мог уже стереть Jobs
      }
    }
    bumpCatalogRev_();
    return {
      ok: true,
      recovered: true,
      trashed: result.trashed,
      needsFirstLaunch: result.needsFirstLaunch,
      message:
        'Очистка завершена' +
        (result.trashed ? ' (в корзину Drive: ' + result.trashed + ')' : '')
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * @returns {(Object|null)}
 */
function findClearCatalogJobAny_() {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].job_type || '') !== 'clear_catalog') {
      continue;
    }
    var st = String(rows[i].status || '').toLowerCase();
    if (st === 'pending' || st === 'running') {
      return rows[i];
    }
  }
  return null;
}
