/**
 * §9.6 — очередь фоновых Jobs + планировщик (time-driven trigger).
 * Импорт Drive / локальный / очистка корзины / копирование F5: через Jobs;
 * одновременно только одна активная очередь.
 */

/** @const {string} */
var JOBS_PROCESS_HANDLER_ = 'processCatalogJobs';

/** @const {number} Файлов за один прогон воркера (acl+copy за один проход). */
var JOBS_CHUNK_SIZE_ = 20;

/** @const {number} Жёсткий потолок файлов в одном импорте Drive (payload). */
var IMPORT_DRIVE_JOB_MAX_FILES_ = 500;

/**
 * Статус очереди для UI (poll).
 *
 * @returns {{
 *   ok: true,
 *   busy: boolean,
 *   activeJob: (Object|null),
 *   progress: number,
 *   progressMessage: string
 * }}
 */
function getCatalogJobsStatus() {
  assertCatalogReadyLight_();
  var active = findActiveCatalogJob_();
  if (!active) {
    return {
      ok: true,
      busy: false,
      activeJob: null,
      progress: 100,
      progressMessage: ''
    };
  }
  return {
    ok: true,
    busy: true,
    activeJob: {
      jobId: active.job_id,
      jobType: active.job_type,
      status: active.status,
      progress: parseNumber_(active.progress),
      progressMessage: String(active.progress_message || ''),
      lastError: String(active.last_error || '')
    },
    progress: parseNumber_(active.progress),
    progressMessage: String(active.progress_message || '')
  };
}

/**
 * Воркер планировщика / one-shot kick.
 * @returns {{ ok: true, processed: number, done: boolean, busy: boolean }}
 */
function processCatalogJobs() {
  assertCatalogReadyLight_();
  ensureCatalogJobsTrigger_();

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { ok: true, processed: 0, done: false, busy: true };
  }

  try {
    var job = findActiveCatalogJob_();
    if (!job) {
      job = findNextPendingCatalogJob_();
      if (!job) {
        return { ok: true, processed: 0, done: true, busy: false };
      }
      markCatalogJobRunning_(job.job_id);
      job = getCatalogJobById_(job.job_id) || job;
    }

    var processed = 0;
    var type = String(job.job_type || '');
    if (type === 'import_drive') {
      processed = processImportDriveJobChunk_(job);
    } else if (type === 'import_upload') {
      // Локальный upload кормит клиент; воркер только обновляет сообщение / завершает пустые.
      processed = processImportUploadJobIdle_(job);
    } else if (type === 'empty_trash') {
      processed = processEmptyTrashJobChunk_(job);
    } else if (type === 'copy_catalog') {
      processed = processCopyCatalogJobChunk_(job);
    } else {
      failCatalogJob_(job.job_id, 'Неизвестный тип задачи: ' + type);
      return { ok: true, processed: 0, done: false, busy: false };
    }

    job = getCatalogJobById_(job.job_id);
    var stillActive = job && (job.status === 'running' || job.status === 'pending');
    // Дальше подхватит минутный триггер; kick только при enqueue (иначе лимит one-shot triggers).
    return {
      ok: true,
      processed: processed,
      done: !stillActive,
      busy: !!stillActive
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * @throws при активной очереди
 */
function assertNoActiveCatalogJobs_() {
  if (findActiveCatalogJob_()) {
    throw catalogError_(
      'JOBS_BUSY',
      'Идут фоновые процессы. Дождитесь завершения текущей очереди и повторите.'
    );
  }
}

/**
 * @returns {boolean}
 */
function hasActiveCatalogJobs_() {
  return !!findActiveCatalogJob_();
}

/**
 * Создаёт time-driven trigger раз в 1 минуту (если ещё нет).
 */
function ensureCatalogJobsTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === JOBS_PROCESS_HANDLER_) {
      return;
    }
  }
  ScriptApp.newTrigger(JOBS_PROCESS_HANDLER_).timeBased().everyMinutes(1).create();
}

/**
 * One-shot kick ~через 1 с (дополнительно к минутному триггеру).
 */
function kickCatalogJobsProcessing_() {
  ensureCatalogJobsTrigger_();
  try {
    ScriptApp.newTrigger(JOBS_PROCESS_HANDLER_).timeBased().after(1000).create();
  } catch (e) {
    // лимит триггеров — минутный всё равно подхватит
  }
}

/**
 * @param {string} jobType
 * @param {Object} payload
 * @param {string} createdBy
 * @param {string=} catalogId
 * @returns {string} jobId
 */
function enqueueCatalogJob_(jobType, payload, createdBy, catalogId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jobs');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Jobs');
  }
  var jobId = Utilities.getUuid();
  var now = new Date();
  var progressMessage = buildInitialJobProgressMessage_(jobType, payload);
  sheet.appendRow([
    jobId,
    jobType,
    'pending',
    catalogId || '',
    JSON.stringify(payload || {}),
    0,
    progressMessage,
    now,
    '',
    '',
    '',
    createdBy || ''
  ]);
  return jobId;
}

/**
 * @param {string} jobType
 * @param {Object=} payload
 * @returns {string}
 */
function buildInitialJobProgressMessage_(jobType, payload) {
  var items = (payload && payload.items) || [];
  var n = items.length;
  var t = String(jobType || '');
  if (t === 'import_upload') {
    return 'Загрузка: 0/' + n;
  }
  if (t === 'empty_trash') {
    var total = n || ((payload && payload.folderIds && payload.folderIds.length) ? 1 : 0);
    return 'Очистка: 0/' + total;
  }
  if (t === 'copy_catalog') {
    return 'Копирование: 0/' + n;
  }
  return 'Импорт: 0/' + n;
}

/**
 * @returns {(Object|null)} row map
 */
function findActiveCatalogJob_() {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    var st = String(rows[i].status || '').toLowerCase();
    if (st === 'running' || st === 'pending') {
      return rows[i];
    }
  }
  return null;
}

/**
 * @returns {(Object|null)}
 */
function findNextPendingCatalogJob_() {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].status || '').toLowerCase() === 'pending') {
      return rows[i];
    }
  }
  return null;
}

/**
 * @param {string} jobId
 * @returns {(Object|null)}
 */
function getCatalogJobById_(jobId) {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].job_id) === String(jobId)) {
      return rows[i];
    }
  }
  return null;
}

/**
 * @param {string} jobId
 */
function markCatalogJobRunning_(jobId) {
  patchCatalogJobRow_(jobId, {
    status: 'running',
    started_at: new Date()
  });
}

/**
 * @param {string} jobId
 * @param {string} message
 */
function failCatalogJob_(jobId, message) {
  patchCatalogJobRow_(jobId, {
    status: 'failed',
    completed_at: new Date(),
    last_error: message || 'Ошибка',
    progress_message: message || 'Ошибка',
    progress: parseNumber_((getCatalogJobById_(jobId) || {}).progress)
  });
}

/**
 * @param {string} jobId
 * @param {Object} fields
 */
function patchCatalogJobRow_(jobId, fields) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jobs');
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
  var idCol = headers.indexOf('job_id');
  if (idCol < 0) {
    return;
  }
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) !== String(jobId)) {
      continue;
    }
    Object.keys(fields).forEach(function (key) {
      var c = headers.indexOf(key);
      if (c >= 0) {
        values[r][c] = fields[key];
      }
    });
    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);
    return;
  }
}

/**
 * @param {Object} job
 * @returns {Object}
 */
function parseJobPayload_(job) {
  try {
    return JSON.parse(String(job.payload_json || '{}')) || {};
  } catch (e) {
    return {};
  }
}

/**
 * @param {string} jobId
 * @param {Object} payload
 * @param {number} progress
 * @param {string} message
 * @param {boolean=} markDone
 */
function saveJobPayloadProgress_(jobId, payload, progress, message, markDone) {
  var fields = {
    payload_json: JSON.stringify(payload),
    progress: Math.max(0, Math.min(100, Math.round(progress))),
    progress_message: message || ''
  };
  if (markDone) {
    fields.status = 'done';
    fields.completed_at = new Date();
  }
  patchCatalogJobRow_(jobId, fields);
  if (markDone) {
    bumpCatalogRev_();
  }
}

/**
 * Чанк F5: Drive makeCopy для pending-файлов; Files обновляются batch.
 *
 * @param {Object} job
 * @returns {number} processed items
 */
function processCopyCatalogJobChunk_(job) {
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  if (!items.length) {
    saveJobPayloadProgress_(job.job_id, payload, 100, 'Копирование: 0/0', true);
    return 0;
  }

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var batch = beginFilesUpdateBatch_();
  var processed = 0;

  for (var i = 0; i < items.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var item = items[i];
    if (item.done) {
      continue;
    }
    try {
      var sourceDriveFile = DriveApp.getFileById(String(item.sourceDriveFileId || ''));
      var driveCopy = sourceDriveFile.makeCopy(
        String(item.displayName || sourceDriveFile.getName()),
        catalogRootFolder
      );
      patchFilesBatchRow_(batch, item.catalogId, {
        fileId: driveCopy.getId(),
        sizeBytes: driveCopy.getSize(),
        driveModifiedAt: driveCopy.getLastUpdated(),
        mimeType: getDriveFileMimeType_(driveCopy),
        status: 'ready',
        lastError: ''
      });
    } catch (eCopy) {
      patchFilesBatchRow_(batch, item.catalogId, {
        status: 'failed',
        lastError: String((eCopy && eCopy.message) || eCopy || 'Ошибка копирования')
      });
    }
    item.done = true;
    processed += 1;
  }

  commitFilesUpdateBatch_(batch);

  var doneCount = 0;
  items.forEach(function (it) {
    if (it.done) {
      doneCount += 1;
    }
  });

  var message = 'Копирование: ' + doneCount + '/' + items.length;
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    (doneCount / items.length) * 100,
    message,
    doneCount >= items.length
  );
  return processed;
}

/**
 * Чанк очистки корзины: Drive trashed + удаление строк Files/ACL; папки Tree — в конце.
 *
 * @param {Object} job
 * @returns {number} processed items
 */
function processEmptyTrashJobChunk_(job) {
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  var folderIds = payload.folderIds || [];
  var progressTotal = items.length || (folderIds.length ? 1 : 0);

  if (!progressTotal) {
    saveJobPayloadProgress_(job.job_id, payload, 100, 'Очистка: 0/0', true);
    return 0;
  }

  var processed = 0;
  var chunkCatalogIds = [];
  var driveErrors = Number(payload.driveErrors) || 0;

  for (var i = 0; i < items.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var item = items[i];
    if (item.done) {
      continue;
    }
    if (item.driveFileId) {
      try {
        moveDriveFileToTrash_(item.driveFileId);
      } catch (eDrive) {
        driveErrors += 1;
      }
    }
    item.done = true;
    chunkCatalogIds.push(item.catalogId);
    processed += 1;
  }

  payload.driveErrors = driveErrors;

  if (chunkCatalogIds.length) {
    removeCatalogFileRows_(chunkCatalogIds);
    removeAclForTrashObjects_(chunkCatalogIds, []);
  }

  var doneCount = 0;
  items.forEach(function (it) {
    if (it.done) {
      doneCount += 1;
    }
  });

  if (items.length && doneCount < items.length) {
    var midProgress = (doneCount / items.length) * (folderIds.length ? 95 : 100);
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      midProgress,
      'Очистка: ' + doneCount + '/' + items.length,
      false
    );
    return processed;
  }

  if (!payload.foldersDone) {
    if (folderIds.length) {
      removeCatalogTreeRows_(folderIds);
      removeAclForTrashObjects_([], folderIds);
    }
    payload.foldersDone = true;
  }

  var finalDone = items.length || 1;
  var finalTotal = items.length || 1;
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    100,
    'Очистка: ' + finalDone + '/' + finalTotal,
    true
  );
  return processed || 1;
}

/**
 * @param {Object} job
 * @returns {number} processed items
 */
function processImportDriveJobChunk_(job) {
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  if (!items.length) {
    saveJobPayloadProgress_(job.job_id, payload, 100, 'Готово', true);
    return 0;
  }

  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var createdBy = String(job.created_by || '');
  payload.phase = 'work';

  var workCtx = beginImportDriveWorkContext_(createdBy);
  var processed = 0;
  var i = 0;
  while (i < items.length && processed < JOBS_CHUNK_SIZE_) {
    var item = items[i];
    if (item.copyDone) {
      i++;
      continue;
    }
    try {
      processImportDriveItemOnce_(
        item,
        payload.mode || 'copy',
        controllerEmail,
        catalogRootFolder,
        workCtx
      );
      item.aclDone = true;
      item.copyDone = true;
      if (!item.error) {
        item.error = '';
      }
    } catch (eWork) {
      item.aclDone = true;
      item.copyDone = true;
      item.error = (eWork && eWork.message) || String(eWork);
      patchFilesBatchRow_(workCtx.filesBatch, item.catalogId, {
        status: 'failed',
        lastError: item.error
      });
    }
    processed++;
    i++;
  }
  commitImportDriveWorkContext_(workCtx);

  var left = 0;
  var failed = 0;
  var done = 0;
  items.forEach(function (it) {
    if (it.copyDone) {
      done++;
    } else {
      left++;
    }
    if (it.error) {
      failed++;
    }
  });

  payload.items = items;
  var total = items.length;
  var progress = total ? Math.round((done / total) * 100) : 100;
  if (left === 0) {
    var msg =
      failed > 0
        ? 'Импорт завершён с ошибками: ' + failed + ' из ' + total
        : 'Импорт завершён: ' + total + ' файл.';
    saveJobPayloadProgress_(job.job_id, payload, 100, msg, true);
  } else {
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      progress,
      'Импорт: ' + done + '/' + total
    );
  }
  return processed;
}

/**
 * Контекст чанка: один снимок Files + пакет Users/ACL.
 * @param {string} createdBy
 * @returns {Object}
 */
function beginImportDriveWorkContext_(createdBy) {
  return {
    createdBy: createdBy || '',
    filesBatch: beginFilesUpdateBatch_(),
    usersBatch: beginUsersEnsureBatch_(),
    aclBatch: beginAclAppendBatch_(),
    aclTouchedIds: []
  };
}

/**
 * @param {Object} ctx
 */
function commitImportDriveWorkContext_(ctx) {
  if (!ctx) {
    return;
  }
  commitUsersEnsureBatch_(ctx.usersBatch);
  commitAclAppendBatch_(ctx.aclBatch);
  commitFilesUpdateBatch_(ctx.filesBatch);
  if (ctx.aclTouchedIds && ctx.aclTouchedIds.length) {
    var engine = createAclEngine_();
    var seen = {};
    ctx.aclTouchedIds.forEach(function (catalogId) {
      if (!catalogId || seen[catalogId]) {
        return;
      }
      seen[catalogId] = true;
      var entries = buildAclEntriesFromObject_(engine, 'file', catalogId);
      syncAclCacheForObjects_(
        [{ objectType: 'file', objectId: catalogId }],
        entries,
        engine
      );
    });
  }
}

/**
 * Один проход: права источника + copy/move (один getFileById).
 *
 * @param {Object} item
 * @param {string} mode
 * @param {string} controllerEmail
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @param {Object} workCtx
 */
function processImportDriveItemOnce_(item, mode, controllerEmail, catalogRootFolder, workCtx) {
  var sourceFile = DriveApp.getFileById(String(item.sourceFileId));
  if (!item.aclDone) {
    applyDriveFileAclToCatalogFile_(
      sourceFile,
      String(item.catalogId),
      workCtx.createdBy,
      workCtx
    );
  }
  var appliedMode = resolveDriveImportPlaceMode_(mode, sourceFile, controllerEmail);
  var catalogFile = placeFileInCatalogRoot_(sourceFile, catalogRootFolder, appliedMode);
  patchFilesBatchRow_(workCtx.filesBatch, item.catalogId, {
    fileId: catalogFile.getId(),
    sizeBytes: catalogFile.getSize(),
    driveModifiedAt: catalogFile.getLastUpdated(),
    mimeType: getDriveFileMimeType_(catalogFile),
    sourceFileId: appliedMode === 'copy' ? String(item.sourceFileId) : '',
    status: 'ready',
    lastError: ''
  });
  item.appliedMode = appliedMode;
}

/**
 * @param {Object} item
 * @param {string} createdBy
 */
function applyImportDriveAclForItem_(item, createdBy) {
  var sourceFile = DriveApp.getFileById(String(item.sourceFileId));
  ensureUsersFromDriveFile_(sourceFile, createdBy);
  applyDriveFileAclToCatalogFile_(sourceFile, String(item.catalogId), createdBy);
}

/**
 * @param {Object} item
 * @param {string} mode
 * @param {string} controllerEmail
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 */
function completeImportDriveCopyForItem_(item, mode, controllerEmail, catalogRootFolder) {
  var sourceFile = DriveApp.getFileById(String(item.sourceFileId));
  var appliedMode = resolveDriveImportPlaceMode_(mode, sourceFile, controllerEmail);
  var catalogFile = placeFileInCatalogRoot_(sourceFile, catalogRootFolder, appliedMode);
  updateCatalogFileAfterDriveImport_(item.catalogId, {
    fileId: catalogFile.getId(),
    sizeBytes: catalogFile.getSize(),
    driveModifiedAt: catalogFile.getLastUpdated(),
    mimeType: getDriveFileMimeType_(catalogFile),
    sourceFileId: appliedMode === 'copy' ? String(item.sourceFileId) : '',
    status: 'ready',
    lastError: ''
  });
  item.appliedMode = appliedMode;
}

/**
 * Локальный job: клиент грузит файлы; воркер завершает, если всё уже done.
 * @param {Object} job
 * @returns {number}
 */
function processImportUploadJobIdle_(job) {
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  if (!items.length) {
    saveJobPayloadProgress_(job.job_id, payload, 100, 'Готово', true);
    return 0;
  }
  var done = 0;
  var failed = 0;
  items.forEach(function (it) {
    if (it.status === 'done') {
      done++;
    } else if (it.status === 'failed') {
      failed++;
      done++;
    }
  });
  var progress = Math.round((done / items.length) * 100);
  if (done >= items.length) {
    var msg =
      failed > 0
        ? 'Загрузка завершена с ошибками: ' + failed
        : 'Загрузка с компьютера завершена';
    saveJobPayloadProgress_(job.job_id, payload, 100, msg, true);
    return 0;
  }
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    progress,
    'Загрузка: ' + done + '/' + items.length
  );
  return 0;
}

/**
 * @returns {{
 *   sheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   values: Array<Array<*>>,
 *   headers: string[],
 *   idCol: number,
 *   dirty: boolean
 * }|null}
 */
function beginFilesUpdateBatch_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 1) {
    return null;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('catalog_id');
  if (idCol < 0) {
    return null;
  }
  return {
    sheet: sheet,
    values: values,
    headers: headers,
    idCol: idCol,
    dirty: false
  };
}

/**
 * @param {Object|null} batch
 * @param {string} catalogId
 * @param {Object} fields
 */
function patchFilesBatchRow_(batch, catalogId, fields) {
  if (!batch || !catalogId) {
    return;
  }
  var map = {
    file_id: fields.fileId,
    size_bytes: fields.sizeBytes,
    drive_modified_at: fields.driveModifiedAt,
    mime_type: fields.mimeType,
    source_file_id: fields.sourceFileId,
    status: fields.status,
    last_error: fields.lastError
  };
  for (var r = 1; r < batch.values.length; r++) {
    if (String(batch.values[r][batch.idCol]) !== String(catalogId)) {
      continue;
    }
    Object.keys(map).forEach(function (key) {
      var c = batch.headers.indexOf(key);
      if (c >= 0 && map[key] !== undefined) {
        batch.values[r][c] = map[key];
      }
    });
    batch.dirty = true;
    return;
  }
}

/**
 * @param {Object|null} batch
 */
function commitFilesUpdateBatch_(batch) {
  if (!batch || !batch.dirty) {
    return;
  }
  batch.sheet
    .getRange(1, 1, batch.values.length, batch.headers.length)
    .setValues(batch.values);
}

/**
 * @param {string} catalogId
 * @param {Object} fields
 */
function updateCatalogFileAfterDriveImport_(catalogId, fields) {
  var batch = beginFilesUpdateBatch_();
  patchFilesBatchRow_(batch, catalogId, fields);
  commitFilesUpdateBatch_(batch);
}

/**
 * @param {string} catalogId
 * @param {string} message
 */
function markCatalogFileFailed_(catalogId, message) {
  updateCatalogFileAfterDriveImport_(catalogId, {
    status: 'failed',
    lastError: message || 'Ошибка импорта'
  });
}

/**
 * @returns {{
 *   sheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   headers: string[],
 *   emailSet: Object.<string, boolean>,
 *   newRows: Array<Array<*>>
 * }|null}
 */
function beginUsersEnsureBatch_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    return null;
  }
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
  var emailCol = headers.indexOf('email');
  if (emailCol < 0) {
    return null;
  }
  var values = sheet.getDataRange().getValues();
  var emailSet = {};
  for (var i = 1; i < values.length; i++) {
    var em = String(values[i][emailCol] || '')
      .trim()
      .toLowerCase();
    if (em) {
      emailSet[em] = true;
    }
  }
  return {
    sheet: sheet,
    headers: headers,
    emailSet: emailSet,
    newRows: []
  };
}

/**
 * @param {Object|null} batch
 * @param {{ email: string, displayName?: string, addedBy?: string }} input
 */
function ensureUserInBatch_(batch, input) {
  if (!batch || !input) {
    return;
  }
  var email = String(input.email || '').trim();
  if (!email) {
    return;
  }
  var key = email.toLowerCase();
  if (batch.emailSet[key]) {
    return;
  }
  batch.emailSet[key] = true;
  var displayName = String(input.displayName || '').trim() || email;
  var byHeader = {
    email: email,
    login_role: 'user',
    added_at: new Date(),
    added_by: input.addedBy || '',
    display_name: displayName
  };
  var line = [];
  for (var c = 0; c < batch.headers.length; c++) {
    var h = batch.headers[c];
    line.push(Object.prototype.hasOwnProperty.call(byHeader, h) ? byHeader[h] : '');
  }
  batch.newRows.push(line);
}

/**
 * @param {Object|null} batch
 */
function commitUsersEnsureBatch_(batch) {
  if (!batch || !batch.newRows.length) {
    return;
  }
  var start = batch.sheet.getLastRow() + 1;
  batch.sheet
    .getRange(start, 1, batch.newRows.length, batch.headers.length)
    .setValues(batch.newRows);
}

/**
 * @returns {{ rows: Array<Array<*>> }}
 */
function beginAclAppendBatch_() {
  return { rows: [] };
}

/**
 * @param {Object} batch
 * @param {string} objectType
 * @param {string} objectId
 * @param {string} email
 * @param {string} permissionLevel
 */
function appendAclInBatch_(batch, objectType, objectId, email, permissionLevel) {
  if (!batch) {
    return;
  }
  batch.rows.push([
    Utilities.getUuid(),
    objectType,
    objectId,
    'user',
    email,
    permissionLevel
  ]);
}

/**
 * @param {Object} batch
 */
function commitAclAppendBatch_(batch) {
  if (!batch || !batch.rows.length) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    return;
  }
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, batch.rows.length, 6).setValues(batch.rows);
}

/**
 * ACL файла Drive → Users + явные права Files (один проход editors/viewers).
 *
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} catalogId
 * @param {string} addedBy
 * @param {Object=} workCtx пакетный контекст чанка (опционально)
 */
function applyDriveFileAclToCatalogFile_(driveFile, catalogId, addedBy, workCtx) {
  var levelsByEmail = {};
  var namesByEmail = {};
  var emailByKey = {};

  function remember(users, level) {
    if (!users) {
      return;
    }
    var list = [];
    if (typeof users.hasNext === 'function') {
      while (users.hasNext()) {
        list.push(users.next());
      }
    } else if (Array.isArray(users)) {
      list = users;
    }
    for (var i = 0; i < list.length; i++) {
      var user = list[i];
      if (!user) {
        continue;
      }
      var email = '';
      try {
        email = String(user.getEmail() || '').trim();
      } catch (e) {
        continue;
      }
      if (!email) {
        continue;
      }
      var key = email.toLowerCase();
      emailByKey[key] = email;
      namesByEmail[key] = resolveDriveUserDisplayName_(user) || email;
      levelsByEmail[key] = maxPermissionLevel_(levelsByEmail[key], level);
    }
  }

  remember(driveFile.getEditors(), 'editor');
  try {
    remember(driveFile.getCommenters(), 'commenter');
  } catch (eC) {
    // optional
  }
  remember(driveFile.getViewers(), 'reader');

  Object.keys(levelsByEmail).forEach(function (key) {
    var level = levelsByEmail[key];
    if (!level || level === 'none') {
      return;
    }
    var email = emailByKey[key] || key;
    var displayName = namesByEmail[key] || email;
    if (workCtx && workCtx.usersBatch) {
      ensureUserInBatch_(workCtx.usersBatch, {
        email: email,
        displayName: displayName,
        addedBy: addedBy || ''
      });
      appendAclInBatch_(workCtx.aclBatch, 'file', catalogId, email, level);
      if (workCtx.aclTouchedIds) {
        workCtx.aclTouchedIds.push(String(catalogId));
      }
      return;
    }
    appendOrEnsureUserRow_({
      email: email,
      loginRole: 'user',
      addedBy: addedBy || '',
      displayName: displayName
    });
    appendExplicitUserAclRow_('file', catalogId, email, level);
  });

  if (!(workCtx && workCtx.usersBatch)) {
    var engine = createAclEngine_();
    var entries = buildAclEntriesFromObject_(engine, 'file', catalogId);
    syncAclCacheForObjects_([{ objectType: 'file', objectId: catalogId }], entries, engine);
  }
}
