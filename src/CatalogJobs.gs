/**
 * §9.6 — очередь фоновых Jobs + планировщик (time-driven trigger).
 * Импорт Drive / локальный / очистка корзины / копирование F5: через Jobs;
 * одновременно только одна активная очередь.
 */

/** @const {string} */
var JOBS_PROCESS_HANDLER_ = 'processCatalogJobs';

/** @const {number} Файлов за один прогон воркера (= размер одной job import_drive). */
var JOBS_CHUNK_SIZE_ = 50;

/** @const {number} Максимум файлов в payload одной job import_drive (= JOBS_CHUNK_SIZE_). */
var IMPORT_DRIVE_JOB_MAX_FILES_ = 50;

/** @const {number} Максимум файлов за одну операцию импорта Drive (цепочка Jobs). */
var IMPORT_DRIVE_OPERATION_MAX_FILES_ = 5000;

/** @const {number} Оценка: секунд на файл (оверхед ACL/учёт/чанк). */
var IMPORT_ETA_SEC_PER_FILE_ = 3;

/** @const {number} Оценка: байт на 1 с доп. времени копирования содержимого. */
var IMPORT_ETA_BYTES_PER_SEC_ = 2 * 1024 * 1024;

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
      progressMessage: '',
      canCancel: false,
      remaining: 0,
      chainTotal: 0,
      estimateRemainingLabel: ''
    };
  }
  var payload = parseJobPayload_(active);
  var chainTotal =
    Number(payload.chainTotalFiles) ||
    ((payload.items && payload.items.length) || 0);
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var items = payload.items || [];
  var doneInJob = 0;
  items.forEach(function (it) {
    if (it && it.copyDone) {
      doneInJob++;
    }
  });
  var globalDone = chainBefore + doneInJob;
  var remaining = Math.max(0, chainTotal - globalDone);
  var eta = estimateImportDriveTimes_(remaining, 0);
  return {
    ok: true,
    busy: true,
    activeJob: {
      jobId: active.job_id,
      jobType: active.job_type,
      status: active.status,
      progress: parseNumber_(active.progress),
      progressMessage: String(active.progress_message || ''),
      lastError: String(active.last_error || ''),
      chainId: String(payload.chainId || ''),
      mode: String(payload.mode || '')
    },
    progress: parseNumber_(active.progress),
    progressMessage: String(active.progress_message || ''),
    canCancel: true,
    remaining: remaining,
    chainTotal: chainTotal,
    done: globalDone,
    estimateRemainingLabel: eta.estimateCopyLabel
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
      return { ok: true, processed: 0, done: true, busy: false };
    }
    if (String(job.status || '').toLowerCase() === 'pending') {
      markCatalogJobRunning_(job.job_id);
      job = getCatalogJobById_(job.job_id) || job;
    }
    if (String(job.status || '').toLowerCase() === 'cancelled') {
      var anyLeftAfterCancel = !!findActiveCatalogJob_();
      return {
        ok: true,
        processed: 0,
        done: !anyLeftAfterCancel,
        busy: anyLeftAfterCancel
      };
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
    // Дальше подхватит минутный триггер; kick только при enqueue (иначе лимит one-shot triggers).
    var anyLeft = !!findActiveCatalogJob_();
    return {
      ok: true,
      processed: processed,
      done: !anyLeft,
      busy: anyLeft
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
  if (t === 'import_drive') {
    var chainTotal = Number(payload && payload.chainTotalFiles) || n;
    var chainBefore = Number(payload && payload.chainDoneBefore) || 0;
    return 'Импорт: ' + chainBefore + '/' + chainTotal;
  }
  return 'Импорт: 0/' + n;
}

/**
 * Оценка времени импорта Drive (§9.6a).
 *
 * @param {number} fileCount
 * @param {number} totalBytes
 * @returns {{
 *   estimateCopySeconds: number,
 *   estimateMoveSeconds: number,
 *   estimateCopyLabel: string,
 *   estimateMoveLabel: string,
 *   jobParts: number
 * }}
 */
function estimateImportDriveTimes_(fileCount, totalBytes) {
  var files = Math.max(0, Math.floor(Number(fileCount) || 0));
  var bytes = Math.max(0, Number(totalBytes) || 0);
  var bytesTerm = bytes > 0 ? Math.ceil(bytes / IMPORT_ETA_BYTES_PER_SEC_) : 0;
  var copySec = Math.round(files * IMPORT_ETA_SEC_PER_FILE_ + bytesTerm);
  var moveSec = Math.round(files * (IMPORT_ETA_SEC_PER_FILE_ * 0.5) + bytesTerm * 0.3);
  return {
    estimateCopySeconds: copySec,
    estimateMoveSeconds: moveSec,
    estimateCopyLabel: formatImportEtaLabel_(copySec),
    estimateMoveLabel: formatImportEtaLabel_(moveSec),
    jobParts: files ? Math.ceil(files / IMPORT_DRIVE_JOB_MAX_FILES_) : 0
  };
}

/**
 * @param {number} seconds
 * @returns {string}
 */
function formatImportEtaLabel_(seconds) {
  var sec = Math.max(0, Number(seconds) || 0);
  if (sec < 60) {
    return '<1 мин';
  }
  var min = Math.round(sec / 60);
  if (min < 1) {
    min = 1;
  }
  return '~' + min + ' мин';
}

/**
 * Ставит одну или несколько Jobs import_drive (нарезка по 500).
 *
 * @param {string} mode
 * @param {string} targetFolderId
 * @param {Array} jobItems
 * @param {string} userEmail
 * @param {string=} scenario
 * @returns {{ jobId: string, chainId: string, fileCount: number, jobParts: number }}
 */
/**
 * Ставит одну или несколько Jobs import_drive (нарезка по IMPORT_DRIVE_JOB_MAX_FILES_).
 *
 * @param {string} mode
 * @param {string} targetFolderId
 * @param {Array} jobItems
 * @param {string} userEmail
 * @param {string=} scenario
 * @param {string=} chainIdOpt заранее сгенерированный chainId (для Files.import_chain_id)
 * @returns {{ jobId: string, chainId: string, fileCount: number, jobParts: number }}
 */
function enqueueImportDriveJobsChain_(mode, targetFolderId, jobItems, userEmail, scenario, chainIdOpt) {
  var items = jobItems || [];
  var total = items.length;
  if (!total) {
    throw catalogError_('INVALID_INPUT', 'Нечего импортировать.');
  }
  if (total > IMPORT_DRIVE_OPERATION_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов (' +
        total +
        '). Максимум за одну операцию: ' +
        IMPORT_DRIVE_OPERATION_MAX_FILES_ +
        '.'
    );
  }
  var chainId = String(chainIdOpt || '').trim() || Utilities.getUuid();
  var parts = Math.ceil(total / IMPORT_DRIVE_JOB_MAX_FILES_);
  var firstJobId = '';
  for (var p = 0; p < parts; p++) {
    var start = p * IMPORT_DRIVE_JOB_MAX_FILES_;
    var chunk = items.slice(start, start + IMPORT_DRIVE_JOB_MAX_FILES_);
    var jobId = enqueueCatalogJob_(
      'import_drive',
      {
        scenario: scenario || 'drive',
        mode: mode,
        targetFolderId: targetFolderId,
        phase: 'work',
        chainId: chainId,
        chainIndex: p,
        chainParts: parts,
        chainTotalFiles: total,
        chainDoneBefore: start,
        items: chunk
      },
      userEmail
    );
    if (!firstJobId) {
      firstJobId = jobId;
    }
  }
  return {
    jobId: firstJobId,
    chainId: chainId,
    fileCount: total,
    jobParts: parts
  };
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
  var fresh = getCatalogJobById_(job.job_id);
  if (fresh && String(fresh.status || '').toLowerCase() === 'cancelled') {
    return 0;
  }

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
  var copiedForAcl = [];
  var processed = 0;
  var i = 0;

  // Фаза 1: только Drive copy/move (без ACL).
  while (i < items.length && processed < JOBS_CHUNK_SIZE_) {
    var item = items[i];
    if (item.copyDone) {
      i++;
      continue;
    }
    try {
      var sourceFile = DriveApp.getFileById(String(item.sourceFileId));
      var appliedMode = resolveDriveImportPlaceMode_(
        payload.mode || 'copy',
        sourceFile,
        controllerEmail
      );
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
      item.copyDone = true;
      item.error = '';
      if (!item.aclDone) {
        copiedForAcl.push({ item: item, sourceFile: sourceFile });
      }
    } catch (eWork) {
      item.copyDone = true;
      item.aclDone = true;
      item.error = (eWork && eWork.message) || String(eWork);
      patchFilesBatchRow_(workCtx.filesBatch, item.catalogId, {
        status: 'failed',
        lastError: item.error
      });
    }
    processed++;
    i++;
  }

  // Фаза 2: batch Files.
  commitFilesUpdateBatch_(workCtx.filesBatch);
  workCtx.filesBatch = beginFilesUpdateBatch_();

  // Фаза 3: ACL батчем по скопированным в этом чанке.
  for (var a = 0; a < copiedForAcl.length; a++) {
    var pair = copiedForAcl[a];
    try {
      applyDriveFileAclToCatalogFile_(
        pair.sourceFile,
        String(pair.item.catalogId),
        createdBy,
        workCtx
      );
      pair.item.aclDone = true;
    } catch (eAcl) {
      pair.item.aclDone = true;
      if (!pair.item.error) {
        pair.item.error = (eAcl && eAcl.message) || String(eAcl);
      }
    }
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
  var chainTotal = Number(payload.chainTotalFiles) || total;
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var globalDone = chainBefore + done;
  var progress = chainTotal ? Math.round((globalDone / chainTotal) * 100) : 100;

  fresh = getCatalogJobById_(job.job_id);
  if (fresh && String(fresh.status || '').toLowerCase() === 'cancelled') {
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      progress,
      'Импорт прерван: ' + globalDone + '/' + chainTotal,
      false
    );
    patchCatalogJobRow_(job.job_id, {
      status: 'cancelled',
      completed_at: new Date(),
      progress_message: 'Импорт прерван: ' + globalDone + '/' + chainTotal
    });
    bumpCatalogRev_();
    return processed;
  }

  if (left === 0) {
    var isLastPart =
      !payload.chainParts ||
      Number(payload.chainIndex) >= Number(payload.chainParts) - 1;
    var msg;
    if (isLastPart) {
      msg =
        failed > 0
          ? 'Импорт завершён с ошибками: ' + failed + ' из ' + total
          : 'Импорт завершён: ' + chainTotal + ' файл.';
    } else {
      msg = 'Импорт: ' + globalDone + '/' + chainTotal;
    }
    saveJobPayloadProgress_(job.job_id, payload, isLastPart ? 100 : progress, msg, true);
  } else {
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      progress,
      'Импорт: ' + globalDone + '/' + chainTotal
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
    var fileUpdates = [];
    ctx.aclTouchedIds.forEach(function (catalogId) {
      if (!catalogId || seen[catalogId]) {
        return;
      }
      seen[catalogId] = true;
      var entries = buildAclEntriesFromObject_(engine, 'file', catalogId);
      var file = engine.filesByCatalogId[catalogId];
      var approved = file && parseBoolean_(file.approved);
      var labels = aclRowsToCacheLabels_(
        engine,
        (entries || []).map(function (e) {
          return {
            principal_type: e.principalType,
            principal_id: e.principalId,
            permission_level: e.permissionLevel
          };
        }),
        approved
      );
      fileUpdates.push({
        catalogId: catalogId,
        aclEditors: formatAclCacheField_(labels.editors),
        aclCommenters: formatAclCacheField_(labels.commenters),
        aclReaders: formatAclCacheField_(labels.readers)
      });
    });
    writeFilesAclCacheBatch_(fileUpdates);
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
    permissionLevel,
    '+'
  ]);
}

/**
 * @param {Object} batch
 */
function commitAclAppendBatch_(batch) {
  if (!batch || !batch.rows.length) {
    return;
  }
  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    return;
  }
  var headers = readSheetHeaderRow_(sheet, 7);
  var width = Math.max(headers.length, 7);
  var start = sheet.getLastRow() + 1;
  var rows = batch.rows.map(function (r) {
    var line = r.slice();
    while (line.length < width) {
      line.push('');
    }
    return line.slice(0, width);
  });
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
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

/**
 * Отмена активной очереди Jobs (§9.6a).
 * import_drive + mode=copy: ready → корзина; pending без file_id → удалить строки.
 *
 * @returns {{
 *   ok: true,
 *   cancelled: boolean,
 *   jobType: string,
 *   mode: string,
 *   chainId: string,
 *   trashed: number,
 *   removedPending: number,
 *   done: number,
 *   chainTotal: number
 * }}
 */
function cancelActiveCatalogJobs() {
  assertCatalogReady_();
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw catalogError_(
      'JOBS_BUSY',
      'Сейчас идёт обработка чанка. Повторите отмену через несколько секунд.'
    );
  }

  try {
    var active = findActiveCatalogJob_();
    if (!active) {
      return {
        ok: true,
        cancelled: false,
        jobType: '',
        mode: '',
        chainId: '',
        trashed: 0,
        removedPending: 0,
        done: 0,
        chainTotal: 0
      };
    }

    var payload = parseJobPayload_(active);
    var chainId = String(payload.chainId || '').trim();
    var mode = String(payload.mode || 'copy');
    var jobType = String(active.job_type || '');
    var chainTotal =
      Number(payload.chainTotalFiles) ||
      ((payload.items && payload.items.length) || 0);
    var chainBefore = Number(payload.chainDoneBefore) || 0;
    var doneInJob = 0;
    (payload.items || []).forEach(function (it) {
      if (it && it.copyDone) {
        doneInJob++;
      }
    });
    var done = chainBefore + doneInJob;

    cancelJobsByChainOrId_(chainId, String(active.job_id));

    var trashed = 0;
    var removedPending = 0;
    if (jobType === 'import_drive' && chainId) {
      if (mode !== 'move') {
        trashed = trashReadyFilesByImportChain_(chainId);
      }
      removedPending = deleteIncompleteFilesByImportChain_(chainId);
    }

    bumpCatalogRev_();
    return {
      ok: true,
      cancelled: true,
      jobType: jobType,
      mode: mode,
      chainId: chainId,
      trashed: trashed,
      removedPending: removedPending,
      done: done,
      chainTotal: chainTotal
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * @param {string} chainId
 * @param {string} fallbackJobId
 */
function cancelJobsByChainOrId_(chainId, fallbackJobId) {
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
  var statusCol = headers.indexOf('status');
  var typeCol = headers.indexOf('job_type');
  var payloadCol = headers.indexOf('payload_json');
  var msgCol = headers.indexOf('progress_message');
  var completedCol = headers.indexOf('completed_at');
  if (idCol < 0 || statusCol < 0) {
    return;
  }

  var now = new Date();
  for (var r = 1; r < values.length; r++) {
    var st = String(values[r][statusCol] || '').toLowerCase();
    if (st !== 'pending' && st !== 'running') {
      continue;
    }
    var match = false;
    if (chainId && payloadCol >= 0) {
      try {
        var pl = JSON.parse(String(values[r][payloadCol] || '{}'));
        if (pl && String(pl.chainId || '') === chainId) {
          match = true;
        }
      } catch (eParse) {
        // ignore
      }
    }
    if (!match && fallbackJobId && String(values[r][idCol]) === String(fallbackJobId)) {
      match = true;
    }
    if (!match && !chainId && fallbackJobId && String(values[r][idCol]) === String(fallbackJobId)) {
      match = true;
    }
    if (!match) {
      continue;
    }
    values[r][statusCol] = 'cancelled';
    if (completedCol >= 0) {
      values[r][completedCol] = now;
    }
    if (msgCol >= 0) {
      values[r][msgCol] = 'Прервано';
    }
  }
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
}

/**
 * Ready-файлы цепочки → `__TRASH__` (без тяжёлого ACL-пересчёта).
 * @param {string} chainId
 * @returns {number}
 */
function trashReadyFilesByImportChain_(chainId) {
  chainId = String(chainId || '').trim();
  if (!chainId) {
    return 0;
  }
  ensureCatalogSchemaUpToDate_();
  var rows = readSheetRecords_('Files');
  var updates = [];
  rows.forEach(function (row) {
    if (String(row.import_chain_id || '') !== chainId) {
      return;
    }
    if (String(row.status || '').toLowerCase() !== 'ready') {
      return;
    }
    if (!String(row.file_id || '').trim()) {
      return;
    }
    if (String(row.folder_id || '') === '__TRASH__') {
      return;
    }
    updates.push({
      catalogId: String(row.catalog_id),
      folderId: '__TRASH__'
    });
  });
  if (!updates.length) {
    return 0;
  }
  applyFileFolderUpdates_(updates, []);
  return updates.length;
}

/**
 * Удаляет pending/failed строки цепочки без готового file_id.
 * @param {string} chainId
 * @returns {number}
 */
function deleteIncompleteFilesByImportChain_(chainId) {
  chainId = String(chainId || '').trim();
  if (!chainId) {
    return 0;
  }
  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    return 0;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return 0;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('catalog_id');
  var chainCol = headers.indexOf('import_chain_id');
  var statusCol = headers.indexOf('status');
  var fileIdCol = headers.indexOf('file_id');
  if (idCol < 0 || chainCol < 0) {
    return 0;
  }

  var keep = [values[0]];
  var removed = 0;
  var removedIds = [];
  for (var r = 1; r < values.length; r++) {
    var rowChain = String(values[r][chainCol] || '');
    if (rowChain !== chainId) {
      keep.push(values[r]);
      continue;
    }
    var st = String(statusCol >= 0 ? values[r][statusCol] : '').toLowerCase();
    var fileId = String(fileIdCol >= 0 ? values[r][fileIdCol] : '').trim();
    if (st === 'ready' && fileId) {
      keep.push(values[r]);
      continue;
    }
    removed++;
    removedIds.push(String(values[r][idCol] || ''));
  }
  if (!removed) {
    return 0;
  }
  sheet.clearContents();
  if (keep.length) {
    sheet.getRange(1, 1, keep.length, headers.length).setValues(keep);
  }
  if (removedIds.length) {
    clearAclRowsForObjects_(
      removedIds.map(function (id) {
        return { objectType: 'file', objectId: id };
      })
    );
  }
  return removed;
}
