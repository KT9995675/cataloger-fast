/**
 * §9.6 — очередь фоновых Jobs + планировщик (time-driven trigger).
 * Импорт Drive / локальный / очистка корзины / копирование F5: через Jobs;
 * одновременно только одна активная очередь.
 */

/** @const {string} */
var JOBS_PROCESS_HANDLER_ = 'processCatalogJobs';

/** @const {number} Файлов за один прогон воркера (= чанк внутри job). */
var JOBS_CHUNK_SIZE_ = 50;

/**
 * Максимум файлов в payload одной job import_drive / copy_catalog.
 * Лимит ячейки Sheets = 50 000 символов; имена в JSON (особенно кириллица `\uXXXX`)
 * быстро раздувают payload — держим запас.
 */
var IMPORT_DRIVE_JOB_MAX_FILES_ = 80;

/** @const {number} То же для F5 copy_catalog (цепочка Jobs). */
var COPY_CATALOG_JOB_MAX_FILES_ = 80;

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
  var jobType = String(active.job_type || '');
  var chainTotal =
    Number(payload.chainTotalFiles) ||
    Number(payload.initialFileCount) ||
    ((payload.items && payload.items.length) || 0);
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var items = payload.items || [];
  var doneInJob = 0;
  items.forEach(function (it) {
    if (jobType === 'clear_catalog') {
      if (it && it.done) {
        doneInJob++;
      }
    } else if (
      jobType === 'move_catalog' ||
      jobType === 'copy_catalog' ||
      jobType === 'empty_trash'
    ) {
      if (it && it.done) {
        doneInJob++;
      }
    } else if (jobType === 'import_upload') {
      if (it && (it.status === 'done' || it.status === 'failed')) {
        doneInJob++;
      }
    } else if (it && it.copyDone) {
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
    estimateRemainingLabel:
      jobType === 'clear_catalog' ? 'несколько секунд' : eta.estimateCopyLabel
  };
}

/**
 * Воркер планировщика / one-shot kick.
 * @returns {{ ok: true, processed: number, done: boolean, busy: boolean }}
 */
function processCatalogJobs() {
  assertCatalogReadyLight_();
  // Пауза абсолютна (старт UI). Не снимать из‑за активных jobs — иначе воркер
  // держит lock и getCatalogTreeSnapshot / Files transfer висят на «Загрузка».
  if (isCatalogJobsPaused_()) {
    return { ok: true, processed: 0, done: true, busy: false, paused: true };
  }
  ensureCatalogJobsTrigger_();

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
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
    } else if (type === 'clear_catalog') {
      processed = processClearCatalogJobChunk_(job);
    } else if (type === 'archive_catalog') {
      processed = processArchiveCatalogJobChunk_(job);
    } else if (type === 'copy_catalog') {
      processed = processCopyCatalogJobChunk_(job);
    } else if (type === 'move_catalog') {
      processed = processMoveCatalogJobChunk_(job);
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
 * @returns {boolean}
 */
function isCatalogJobsPaused_() {
  return (
    PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_JOBS_PAUSED_) === '1'
  );
}

/**
 * @param {boolean} paused
 */
function setCatalogJobsPaused_(paused) {
  var props = PropertiesService.getDocumentProperties();
  if (paused) {
    props.setProperty(PROP_CATALOG_JOBS_PAUSED_, '1');
  } else {
    props.deleteProperty(PROP_CATALOG_JOBS_PAUSED_);
  }
}

/**
 * Удаляет time-driven / one-shot триггеры processCatalogJobs (на время загрузки UI).
 * @returns {number}
 */
function removeCatalogJobsTriggers_() {
  var removed = 0;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === JOBS_PROCESS_HANDLER_) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return removed;
}

/**
 * Помечает pending/running как failed — только колонка status (без parse JSON payload).
 * @returns {number}
 */
function abortActiveJobStatusesLight_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jobs');
  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  var statusCol = headers.indexOf('status');
  var errorCol = headers.indexOf('last_error');
  var doneCol = headers.indexOf('completed_at');
  if (statusCol < 0) {
    return 0;
  }
  var lastRow = sheet.getLastRow();
  var n = lastRow - 1;
  var statuses = sheet.getRange(2, statusCol + 1, n, 1).getValues();
  var aborted = 0;
  var now = new Date();
  var errors = errorCol >= 0 ? sheet.getRange(2, errorCol + 1, n, 1).getValues() : null;
  var dones = doneCol >= 0 ? sheet.getRange(2, doneCol + 1, n, 1).getValues() : null;
  for (var i = 0; i < statuses.length; i++) {
    var st = String(statuses[i][0] || '').toLowerCase();
    if (st !== 'pending' && st !== 'running') {
      continue;
    }
    statuses[i][0] = 'failed';
    if (errors) {
      errors[i][0] = 'Прервано: разблокировка загрузки UI';
    }
    if (dones) {
      dones[i][0] = now;
    }
    aborted++;
  }
  if (aborted) {
    sheet.getRange(2, statusCol + 1, n, 1).setValues(statuses);
    if (errors) {
      sheet.getRange(2, errorCol + 1, n, 1).setValues(errors);
    }
    if (dones) {
      sheet.getRange(2, doneCol + 1, n, 1).setValues(dones);
    }
  }
  return aborted;
}

/**
 * One-shot kick ~через 1 с (дополнительно к минутному триггеру).
 */
function kickCatalogJobsProcessing_() {
  setCatalogJobsPaused_(false);
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
  var compact = compactCatalogJobPayload_(payload || {});
  var json = JSON.stringify(compact);
  if (json.length > 49000) {
    compact = compactCatalogJobPayloadAggressive_(compact);
    json = JSON.stringify(compact);
  }
  if (json.length > 50000) {
    throw catalogError_(
      'JOB_PAYLOAD_TOO_LARGE',
      'Слишком большой payload Jobs (' + json.length + ' симв.).'
    );
  }
  sheet.appendRow([
    jobId,
    jobType,
    'pending',
    catalogId || '',
    json,
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
  if (t === 'clear_catalog') {
    var clearTotal =
      Number(payload && payload.initialFileCount) || n || 0;
    return 'Очистка каталога: 0/' + clearTotal;
  }
  if (t === 'archive_catalog') {
    var folders = Number(payload && payload.foldersTotal) || 0;
    var files = Number(payload && payload.filesTotal) || 0;
    if (folders) {
      return 'Архивация (папки): 0/' + folders;
    }
    return 'Архивация (файлы): 0/' + files;
  }
  if (t === 'copy_catalog') {
    var copyTotal = Number(payload && payload.chainTotalFiles) || n;
    var copyBefore = Number(payload && payload.chainDoneBefore) || 0;
    return 'Копирование: ' + copyBefore + '/' + copyTotal;
  }
  if (t === 'move_catalog') {
    var moveTarget = String((payload && payload.targetFolderId) || '');
    var moveLabel = moveTarget === '__TRASH__' ? 'Удаление' : 'Перемещение';
    return moveLabel + ': 0/' + n;
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
  // Move: batch place + отложенный ACL (§9.12) — заметно быстрее copy.
  var moveSec = Math.round(files * 0.35 + bytesTerm * 0.15 + files * 0.4);
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
        phase: mode === 'move' ? 'place' : 'copy',
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
  var clearJob = null;
  var frontJob = null; // move/place или copy/meta|copy — раньше ACL
  var otherJob = null;
  for (var i = 0; i < rows.length; i++) {
    var st = String(rows[i].status || '').toLowerCase();
    if (st !== 'running' && st !== 'pending') {
      continue;
    }
    var type = String(rows[i].job_type || '');
    var payload = parseJobPayload_(rows[i]);
    var phase = String(payload.phase || '');
    var mode = String(payload.mode || '');
    if (type === 'clear_catalog') {
      if (!clearJob) {
        clearJob = rows[i];
      }
      continue;
    }
    // Move place / copy meta|copy — приоритетнее ACL по цепочке.
    if (type === 'import_drive' && mode === 'move' && (phase === 'place' || phase === '')) {
      if (!frontJob) {
        frontJob = rows[i];
      }
      continue;
    }
    if (
      type === 'import_drive' &&
      mode !== 'move' &&
      (phase === 'meta' || phase === 'copy' || phase === 'work' || phase === '')
    ) {
      if (!frontJob) {
        frontJob = rows[i];
      }
      continue;
    }
    if (!otherJob) {
      otherJob = rows[i];
    }
  }
  return clearJob || frontJob || otherJob;
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
  var compact = compactCatalogJobPayload_(payload);
  var json = JSON.stringify(compact);
  // Защита от лимита ячейки Sheets (50000).
  if (json.length > 49000) {
    compact = compactCatalogJobPayloadAggressive_(compact);
    json = JSON.stringify(compact);
  }
  if (json.length > 50000) {
    throw catalogError_(
      'JOB_PAYLOAD_TOO_LARGE',
      'Слишком большой payload Jobs (' +
        json.length +
        ' симв.). Уменьшите порцию импорта.'
    );
  }
  var fields = {
    payload_json: json,
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
 * Убирает из payload тяжёлые поля (имена, размеры) — они есть в Files.
 * @param {Object} payload
 * @returns {Object}
 */
function compactCatalogJobPayload_(payload) {
  payload = payload || {};
  var out = {};
  Object.keys(payload).forEach(function (k) {
    if (k === 'items') {
      return;
    }
    out[k] = payload[k];
  });
  out.items = (payload.items || []).map(function (it) {
    if (!it) {
      return it;
    }
    var sourceId = String(it.sourceFileId || it.sourceDriveFileId || '');
    var row = {
      catalogId: it.catalogId,
      sourceFileId: sourceId,
      done: !!it.done
    };
    // empty_trash / clear — нужен drive file id (не путать с sourceFileId импорта).
    if (it.driveFileId) {
      row.driveFileId = String(it.driveFileId);
    }
    if (it.sourceDriveFileId && !row.sourceFileId) {
      row.sourceFileId = String(it.sourceDriveFileId);
    }
    // import_upload — клиентский upload; нельзя срезать status/relativePath.
    if (
      Object.prototype.hasOwnProperty.call(it, 'relativePath') ||
      it.status === 'waiting' ||
      it.status === 'done' ||
      it.status === 'failed'
    ) {
      row.relativePath = String(it.relativePath || '');
      row.fileName = String(it.fileName || '').slice(0, 200);
      row.parentFolderId = String(it.parentFolderId || '');
      row.status = String(it.status || 'waiting');
      row.sizeBytes = parseNumber_(it.sizeBytes);
      row.mimeType = String(it.mimeType || '').slice(0, 120);
      if (it.index != null) {
        row.index = Number(it.index) || 0;
      }
      if (it.error) {
        row.error = String(it.error).slice(0, 160);
      }
      return row;
    }
    // import_drive (и др. с фазами) — сохраняем флаги; lean copy_catalog — только id+source+done
    if (
      Object.prototype.hasOwnProperty.call(it, 'copyDone') ||
      Object.prototype.hasOwnProperty.call(it, 'aclDone') ||
      Object.prototype.hasOwnProperty.call(it, 'metaDone')
    ) {
      row.copyDone = !!it.copyDone;
      row.aclDone = !!it.aclDone;
      row.metaDone = it.metaDone !== false;
    }
    if (it.error) {
      row.error = String(it.error).slice(0, 160);
    }
    if (it.placedFileId) {
      row.placedFileId = String(it.placedFileId);
    }
    if (it.appliedMode) {
      row.appliedMode = String(it.appliedMode);
    }
    // Имена не кладём при chain (лимит 50k) — есть в Files / Drive.
    if (it.displayName && !payload.chainId) {
      row.displayName = String(it.displayName).slice(0, 80);
    }
    // move_catalog folder/file shape
    if (it.kind) {
      row.kind = String(it.kind);
    }
    if (it.folderId) {
      row.folderId = String(it.folderId);
    }
    if (it.parentFolderId) {
      row.parentFolderId = String(it.parentFolderId);
    }
    return row;
  });
  return out;
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function compactCatalogJobPayloadAggressive_(payload) {
  var out = compactCatalogJobPayload_(payload);
  (out.items || []).forEach(function (it) {
    if (!it) {
      return;
    }
    delete it.displayName;
    delete it.error;
    delete it.appliedMode;
  });
  return out;
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
  var nameCol = batch && batch.headers ? batch.headers.indexOf('display_name') : -1;
  var sizeCol = batch && batch.headers ? batch.headers.indexOf('size_bytes') : -1;
  var processed = 0;
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var chainTotal = Number(payload.chainTotalFiles) || items.length;

  for (var i = 0; i < items.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var item = items[i];
    if (item.done) {
      continue;
    }
    try {
      var sourceId = String(item.sourceDriveFileId || item.sourceFileId || '');
      var sourceDriveFile = DriveApp.getFileById(sourceId);
      var rowMeta = findFilesBatchRowMeta_(batch, item.catalogId, nameCol, sizeCol);
      var driveCopy = sourceDriveFile.makeCopy(
        String(item.displayName || rowMeta.displayName || sourceDriveFile.getName()),
        catalogRootFolder
      );
      var mimeType = getDriveFileMimeType_(driveCopy) || '';
      // Docs/Sheets: getSize() stub 0/1 — resolve (+ quotaBytesUsed); не наследовать stub с источника.
      var sizeBytes = resolveDriveFileSizeBytes_(driveCopy, mimeType);
      if (!sizeBytes) {
        var fromSource =
          Number(item.sourceSizeBytes) || Number(rowMeta.sizeBytes) || 0;
        sizeBytes = normalizeCatalogSizeBytes_(fromSource, mimeType);
      }
      patchFilesBatchRow_(batch, item.catalogId, {
        fileId: driveCopy.getId(),
        sizeBytes: sizeBytes,
        driveModifiedAt: driveCopy.getLastUpdated(),
        mimeType: mimeType,
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

  var globalDone = chainBefore + doneCount;
  var message = 'Копирование: ' + globalDone + '/' + chainTotal;
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    chainTotal ? (globalDone / chainTotal) * 100 : 100,
    message,
    doneCount >= items.length
  );
  return processed;
}

/**
 * @param {Object|null} batch
 * @param {string} catalogId
 * @param {number} nameCol
 * @param {number} sizeCol
 * @returns {{ displayName: string, sizeBytes: number }}
 */
function findFilesBatchRowMeta_(batch, catalogId, nameCol, sizeCol) {
  var out = { displayName: '', sizeBytes: 0 };
  if (!batch || !catalogId) {
    return out;
  }
  for (var r = 1; r < batch.values.length; r++) {
    if (String(batch.values[r][batch.idCol]) !== String(catalogId)) {
      continue;
    }
    if (nameCol >= 0) {
      out.displayName = String(batch.values[r][nameCol] || '');
    }
    if (sizeCol >= 0) {
      out.sizeBytes = Number(batch.values[r][sizeCol]) || 0;
    }
    return out;
  }
  return out;
}

/**
 * §0.4a — F6/F8: чанки folder_id / parent + cascade trash + ACL в конце.
 *
 * @param {Object} job
 * @returns {number}
 */
function processMoveCatalogJobChunk_(job) {
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  var targetFolderId = String(payload.targetFolderId || '');
  var label = targetFolderId === '__TRASH__' ? 'Удаление' : 'Перемещение';
  var phase = String(payload.phase || 'apply');

  if (!items.length) {
    saveJobPayloadProgress_(job.job_id, payload, 100, label + ': 0/0', true);
    return 0;
  }

  if (phase === 'acl') {
    // keep-deltas: сброс/кэш не делаем (иначе долгий rewrite + lock на загрузке UI).
    var doneAll = items.length;
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      100,
      label + ': ' + doneAll + '/' + doneAll,
      true
    );
    bumpCatalogRev_();
    return 1;
  }

  var fileUpdates = [];
  var folderUpdates = [];
  var processed = 0;
  for (var i = 0; i < items.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var it = items[i];
    if (it.done) {
      continue;
    }
    if (it.kind === 'file') {
      fileUpdates.push({
        catalogId: String(it.catalogId),
        folderId: String(it.folderId)
      });
    } else {
      folderUpdates.push({
        folderId: String(it.folderId),
        parentFolderId: String(it.parentFolderId)
      });
    }
    it.done = true;
    processed += 1;
  }

  if (fileUpdates.length || folderUpdates.length) {
    applyFileFolderUpdates_(fileUpdates, folderUpdates);
  }

  var doneCount = 0;
  items.forEach(function (x) {
    if (x.done) {
      doneCount += 1;
    }
  });

  if (doneCount >= items.length) {
    var trashCascadeTargets = payload.trashCascadeTargets || [];
    var filesMovedToTrash = payload.filesMovedToTrash || [];
    if (trashCascadeTargets.length || filesMovedToTrash.length) {
      if (trashCascadeTargets.length) {
        deleteMirrorsPointingToFolders_(trashCascadeTargets);
      }
      if (filesMovedToTrash.length) {
        deleteFileShortcutsPointingToCatalogIds_(filesMovedToTrash);
      }
    }
    payload.phase = 'acl';
    payload.items = items;
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      Math.max(1, (doneCount / items.length) * 95),
      label + ': ' + doneCount + '/' + items.length,
      false
    );
    return processed || 1;
  }

  payload.items = items;
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    (doneCount / items.length) * 100,
    label + ': ' + doneCount + '/' + items.length,
    false
  );
  bumpCatalogRev_();
  return processed;
}

/**
 * Чанк очистки корзины: Drive trashed + удаление строк Files/ACL; папки Tree — в конце.
 * Перед завершением — повторный collect с листа (файлы «просто в корзине» без подпапок
 * не должны оставаться после «Готово»).
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
    // Пустой payload при постановке — всё равно проверим лист.
    var seed = collectEmptyTrashFileItemsFromSheet_();
    var seedFolders = collectEmptyTrashFolderIdsFromSheet_();
    if (seed.length || seedFolders.length) {
      payload.items = seed;
      payload.folderIds = seedFolders;
      payload.foldersDone = false;
      items = seed;
      folderIds = seedFolders;
      progressTotal = items.length || (folderIds.length ? 1 : 0);
    } else {
      saveJobPayloadProgress_(job.job_id, payload, 100, 'Очистка: 0/0', true);
      return 0;
    }
  }

  var processed = 0;
  var chunkCatalogIds = [];
  var driveErrors = Number(payload.driveErrors) || 0;

  for (var i = 0; i < items.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var item = items[i];
    if (item.done) {
      continue;
    }
    var driveId = String(item.driveFileId || item.sourceFileId || '').trim();
    if (!driveId && item.catalogId) {
      try {
        var fr = readSheetRecords_('Files');
        for (var fi = 0; fi < fr.length; fi++) {
          if (String(fr[fi].catalog_id || '') === String(item.catalogId)) {
            driveId = String(fr[fi].file_id || '').trim();
            break;
          }
        }
        if (driveId) {
          item.driveFileId = driveId;
        }
      } catch (eLookup) {
        // ignore
      }
    }
    if (driveId) {
      try {
        // Не throw на ошибке одного файла — строку каталога всё равно убираем.
        moveDriveFilesToTrashBatch_([driveId]);
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
    var midProgress = (doneCount / Math.max(items.length, 1)) * (folderIds.length ? 95 : 100);
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      midProgress,
      'Очистка: ' + doneCount + '/' + items.length,
      false
    );
    return processed;
  }

  // Хвост: файлы, не попавшие в исходный payload (часто «просто в __TRASH__»).
  var leftovers = collectEmptyTrashFileItemsFromSheet_();
  if (leftovers.length) {
    payload.items = items.concat(leftovers);
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      90,
      'Очистка: добор ' + leftovers.length + '…',
      false
    );
    return leftovers.length;
  }

  if (!payload.foldersDone) {
    var foldersLeft = collectEmptyTrashFolderIdsFromSheet_();
    if (foldersLeft.length) {
      removeCatalogTreeRows_(foldersLeft);
      removeAclForTrashObjects_([], foldersLeft);
    } else if (folderIds.length) {
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
  bumpCatalogRev_();
  return processed || 1;
}

/**
 * Актуальные файлы в поддереве корзины с листа Files.
 * @returns {Array<{ catalogId: string, driveFileId: string, done: boolean }>}
 */
function collectEmptyTrashFileItemsFromSheet_() {
  var engine = createAclEngine_();
  if (!engine.foldersById[TRASH_FOLDER_ID_EMPTY_]) {
    return [];
  }
  var trashFolderIds = collectTrashSubtreeFolderIds_(engine);
  var items = [];
  Object.keys(engine.filesByCatalogId).forEach(function (catalogId) {
    var file = engine.filesByCatalogId[catalogId];
    var folderId = String(file.folder_id || '').trim();
    if (!trashFolderIds[folderId]) {
      return;
    }
    items.push({
      catalogId: String(catalogId),
      driveFileId: String(file.file_id || '').trim(),
      done: false
    });
  });
  return items;
}

/**
 * Папки внутри корзины (без самой __TRASH__).
 * @returns {string[]}
 */
function collectEmptyTrashFolderIdsFromSheet_() {
  var engine = createAclEngine_();
  if (!engine.foldersById[TRASH_FOLDER_ID_EMPTY_]) {
    return [];
  }
  var trashFolderIds = collectTrashSubtreeFolderIds_(engine);
  return Object.keys(trashFolderIds).filter(function (id) {
    return id !== TRASH_FOLDER_ID_EMPTY_;
  });
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
  var mode = String(payload.mode || 'copy');
  if (mode === 'move') {
    return processImportDriveMoveJobChunkWithBudget_(job, payload);
  }
  return processImportDriveCopyJobChunk_(job, payload);
}

/**
 * Move: фаза place (batch) → фаза acl (batch permissions → Users/ACL).
 * Place-джобы цепочки приоритетнее ACL (§ findActiveCatalogJob_).
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveMoveJobChunkWithBudget_(job, payload) {
  var startMs = Date.now();
  var budgetMs = 8000;
  var totalProcessed = 0;
  while (Date.now() - startMs < budgetMs) {
    if (isCatalogJobsPaused_()) {
      break;
    }
    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    if (String(job.status || '').toLowerCase() === 'cancelled') {
      break;
    }
    payload = parseJobPayload_(job);
    var phase = String(payload.phase || 'place');
    if (phase === 'acl' && hasOtherImportDriveMovePlaceJobs_(job.job_id)) {
      // Сначала добить все place по цепочке.
      break;
    }
    var n =
      phase === 'acl'
        ? processImportDriveMoveAclPhase_(job, payload)
        : processImportDriveMovePlacePhase_(job, payload);
    totalProcessed += n || 0;
    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    if (!n) {
      break;
    }
  }
  return totalProcessed;
}

/**
 * Есть ли ещё move/place jobs (кроме текущей).
 *
 * @param {string} exceptJobId
 * @returns {boolean}
 */
function hasOtherImportDriveMovePlaceJobs_(exceptJobId) {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].job_id) === String(exceptJobId)) {
      continue;
    }
    var st = String(rows[i].status || '').toLowerCase();
    if (st !== 'running' && st !== 'pending') {
      continue;
    }
    if (String(rows[i].job_type || '') !== 'import_drive') {
      continue;
    }
    var payload = parseJobPayload_(rows[i]);
    if (String(payload.mode || '') !== 'move') {
      continue;
    }
    var phase = String(payload.phase || 'place');
    if (phase === 'place' || phase === '') {
      return true;
    }
  }
  return false;
}

/**
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveMovePlacePhase_(job, payload) {
  var items = payload.items || [];
  if (!items.length) {
    payload.phase = 'acl';
    saveJobPayloadProgress_(job.job_id, payload, 50, 'Импорт: права…', false);
    return 1;
  }

  var pending = [];
  items.forEach(function (item) {
    if (item.copyDone) {
      return;
    }
    pending.push(item);
  });

  if (!pending.length) {
    payload.phase = 'acl';
    saveImportDriveProgress_(job, payload, items, 0);
    return 1;
  }

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var chunk = pending.slice(0, JOBS_CHUNK_SIZE_);
  var filesBatch = beginFilesUpdateBatch_();
  var placed = 0;

  for (var i = 0; i < chunk.length; i++) {
    var item = chunk[i];
    try {
      var catalogFile = longOpsTransferIntoCatalogRoot_(
        String(item.sourceFileId),
        'move',
        catalogRootFolder
      );
      var sizeBytes = 0;
      var mimeType = '';
      var modifiedAt = new Date();
      try {
        mimeType = getDriveFileMimeType_(catalogFile) || '';
        sizeBytes = resolveDriveFileSizeBytes_(catalogFile, mimeType);
        modifiedAt = catalogFile.getLastUpdated();
      } catch (eMeta) {
        // keep defaults
      }
      item.copyDone = true;
      item.appliedMode = 'move';
      item.error = '';
      item.placedFileId = catalogFile.getId();
      patchFilesBatchRow_(filesBatch, item.catalogId, {
        fileId: catalogFile.getId(),
        sizeBytes: sizeBytes,
        driveModifiedAt: modifiedAt,
        mimeType: mimeType,
        sourceFileId: '',
        status: 'ready',
        lastError: ''
      });
      placed += 1;
    } catch (eMove) {
      item.copyDone = true;
      item.aclDone = true;
      item.error = (eMove && eMove.message) || String(eMove);
      longOpsMarkFileFailed_(item.catalogId, item.error, filesBatch);
    }
  }

  commitFilesUpdateBatch_(filesBatch);
  payload.items = items;
  var leftPlace = 0;
  items.forEach(function (it) {
    if (!it.copyDone) {
      leftPlace += 1;
    }
  });
  if (leftPlace === 0) {
    payload.phase = 'acl';
  }
  saveImportDriveProgress_(job, payload, items, placed || chunk.length);
  return chunk.length;
}

/**
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveMoveAclPhase_(job, payload) {
  var items = payload.items || [];
  var createdBy = String(job.created_by || '');
  var pendingAcl = [];
  items.forEach(function (item) {
    if (!item.copyDone || item.aclDone || item.error) {
      return;
    }
    var fileId = String(item.placedFileId || item.sourceFileId || '').trim();
    if (!fileId) {
      item.aclDone = true;
      return;
    }
    pendingAcl.push(item);
  });

  if (!pendingAcl.length) {
    saveImportDriveProgress_(job, payload, items, 0);
    return 1;
  }

  var chunk = pendingAcl.slice(0, JOBS_CHUNK_SIZE_);
  var workCtx = beginImportDriveWorkContext_(createdBy);
  for (var i = 0; i < chunk.length; i++) {
    var item = chunk[i];
    var catalogId = String(item.catalogId || '');
    var fileId = String(item.placedFileId || item.sourceFileId || '');
    if (workCtx.aclTouchedIds && catalogId) {
      workCtx.aclTouchedIds.push(catalogId);
    }
    try {
      var driveFile = DriveApp.getFileById(fileId);
      applyDriveFileAclToCatalogFile_(driveFile, catalogId, createdBy, workCtx);
      item.aclDone = true;
    } catch (eAcl) {
      item.aclDone = true;
      item.error = (eAcl && eAcl.message) || String(eAcl);
      longOpsMarkFileFailed_(catalogId, item.error, workCtx.filesBatch);
    }
  }
  commitImportDriveWorkContext_(workCtx);
  payload.items = items;
  saveImportDriveProgress_(job, payload, items, chunk.length);
  return chunk.length;
}

/**
 * §4.4a / ремонт ACL_REDUNDANT_PLUS: не писать delta=+ на файл,
 * если у родительской папки уже есть тот же principal с уровнем ≥.
 *
 * @param {Object} engine
 * @param {string} catalogId
 * @param {string} email
 * @param {string} level
 * @returns {boolean} true → пропускаем запись ACL
 */
function shouldSkipRedundantFileAclPlus_(engine, catalogId, email, level) {
  if (!engine || !catalogId || !email || !level || level === 'none') {
    return false;
  }
  var file = engine.filesByCatalogId && engine.filesByCatalogId[String(catalogId)];
  if (!file || !file.folder_id) {
    return false;
  }
  var parentMap = getEffectiveAclMapFromEngine_(
    engine,
    'folder',
    String(file.folder_id)
  );
  var pKey = aclPrincipalMapKey_('user', email);
  if (!pKey) {
    return false;
  }
  var parentLevel = parentMap[pKey] ? parentMap[pKey].level : 'none';
  var rank = { none: 0, reader: 1, commenter: 2, editor: 3 };
  return (rank[parentLevel] || 0) >= (rank[level] || 0);
}

/**
 * Размер файла из Drive API (size или quotaBytesUsed).
 * @param {*} raw
 * @returns {number}
 */
function parseDriveSizeBytes_(raw) {
  if (raw == null || raw === '' || typeof raw === 'boolean') {
    return 0;
  }
  var n = Number(raw);
  if (!isFinite(n) || n < 0) {
    return 0;
  }
  return Math.floor(n);
}

/**
 * Batch: permissions.list → Users + ACL листа (без DriveApp getEditors).
 * Не пишет ACL_REDUNDANT_PLUS (noop + относительно матери).
 *
 * @param {Array<{ item: Object, fileId: string }>} pairs
 * @param {string} createdBy
 * @param {Object} workCtx
 */
function applyDriveFilePermissionsBatchToCatalog_(pairs, createdBy, workCtx) {
  if (!pairs || !pairs.length) {
    return;
  }
  var token = ScriptApp.getOAuthToken();
  var listReqs = pairs.map(function (p) {
    return {
      url:
        'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(p.fileId) +
        '/permissions?supportsAllDrives=true&fields=' +
        encodeURIComponent('permissions(id,type,role,emailAddress,displayName)'),
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };
  });
  var listResps = UrlFetchApp.fetchAll(listReqs);
  var engine = createAclEngine_();

  for (var i = 0; i < pairs.length; i++) {
    var catalogId = String(pairs[i].item.catalogId || '');
    if (!catalogId) {
      continue;
    }
    // Кэш пересчитаем даже если все ACL отфильтрованы как redundant.
    if (workCtx && workCtx.aclTouchedIds) {
      workCtx.aclTouchedIds.push(catalogId);
    }
    var resp = listResps[i];
    if (!resp || resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
      continue;
    }
    var body = {};
    try {
      body = JSON.parse(resp.getContentText() || '{}');
    } catch (eParse) {
      continue;
    }
    var perms = body.permissions || [];
    for (var p = 0; p < perms.length; p++) {
      var perm = perms[p];
      if (!perm || String(perm.type || '') !== 'user') {
        continue;
      }
      var email = String(perm.emailAddress || '').trim();
      if (!email) {
        continue;
      }
      var role = String(perm.role || '');
      var level = drivePermissionRoleToCatalogLevel_(role);
      if (!level) {
        continue;
      }
      var controllerLc = String(
        PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || ''
      )
        .trim()
        .toLowerCase();
      if (controllerLc && email.toLowerCase() === controllerLc) {
        continue;
      }
      if (shouldSkipRedundantFileAclPlus_(engine, catalogId, email, level)) {
        continue;
      }
      var displayName = String(perm.displayName || '').trim() || email;
      if (workCtx && workCtx.usersBatch) {
        ensureUserInBatch_(workCtx.usersBatch, {
          email: email,
          displayName: displayName,
          addedBy: createdBy || ''
        });
        appendAclInBatch_(workCtx.aclBatch, 'file', catalogId, email, level);
      } else {
        appendOrEnsureUserRow_({
          email: email,
          loginRole: 'user',
          addedBy: createdBy || '',
          displayName: displayName
        });
        appendExplicitUserAclRow_('file', catalogId, email, level);
      }
    }
  }
}

/**
 * Copy: meta (объём+права источника) → copy (makeCopy) → acl (фон из кэша).
 * Meta/copy jobs цепочки приоритетнее ACL (§ findActiveCatalogJob_).
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveCopyJobChunk_(job, payload) {
  var startMs = Date.now();
  var budgetMs = 25000;
  var totalProcessed = 0;
  while (Date.now() - startMs < budgetMs) {
    if (isCatalogJobsPaused_()) {
      break;
    }
    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    if (String(job.status || '').toLowerCase() === 'cancelled') {
      break;
    }
    payload = parseJobPayload_(job);
    var phase = String(payload.phase || 'meta');
    if (phase === 'work') {
      phase = 'copy';
      payload.phase = 'copy';
    }
    if (phase === 'acl' && hasOtherImportDriveCopyFrontJobs_(job.job_id)) {
      break;
    }
    var n = 0;
    if (phase === 'meta') {
      n = processImportDriveCopyMetaPhase_(job, payload);
    } else if (phase === 'acl') {
      n = processImportDriveCopyAclPhase_(job, payload);
    } else {
      n = processImportDriveCopyCopyPhase_(job, payload);
    }
    totalProcessed += n || 0;
    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    if (!n) {
      break;
    }
  }
  return totalProcessed;
}

/**
 * Есть ли ещё copy jobs в фазе meta/copy (кроме текущей).
 *
 * @param {string} exceptJobId
 * @returns {boolean}
 */
function hasOtherImportDriveCopyFrontJobs_(exceptJobId) {
  var rows = readSheetRecords_('Jobs');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].job_id) === String(exceptJobId)) {
      continue;
    }
    var st = String(rows[i].status || '').toLowerCase();
    if (st !== 'running' && st !== 'pending') {
      continue;
    }
    if (String(rows[i].job_type || '') !== 'import_drive') {
      continue;
    }
    var payload = parseJobPayload_(rows[i]);
    if (String(payload.mode || 'copy') === 'move') {
      continue;
    }
    var phase = String(payload.phase || 'meta');
    if (phase === 'meta' || phase === 'copy' || phase === 'work' || phase === '') {
      return true;
    }
  }
  return false;
}

/**
 * Фаза meta (legacy): больше не читаем Drive заранее — сразу в copy.
 * Старые jobs в phase=meta тоже быстро переводим дальше (иначе 0/N при таймауте).
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveCopyMetaPhase_(job, payload) {
  var items = payload.items || [];
  items.forEach(function (item) {
    if (!item.metaDone && !item.error) {
      item.metaDone = true;
    }
  });
  payload.items = items;
  payload.phase = 'copy';
  saveImportDriveCopyPhaseProgress_(job, payload, items, 'copy');
  return 1;
}

/**
 * Фаза copy: только makeCopy + Files (без ACL).
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveCopyCopyPhase_(job, payload) {
  var items = payload.items || [];
  if (!items.length) {
    payload.phase = 'acl';
    saveJobPayloadProgress_(job.job_id, payload, 50, 'Импорт (права)…', false);
    return 1;
  }

  // Legacy: без отдельной meta — размеры берём после makeCopy.
  items.forEach(function (it) {
    if (!it.metaDone && !it.error) {
      it.metaDone = true;
    }
  });

  var pending = [];
  items.forEach(function (item) {
    if (item.copyDone || item.error) {
      return;
    }
    pending.push(item);
  });

  if (!pending.length) {
    payload.phase = 'acl';
    saveImportDriveCopyPhaseProgress_(job, payload, items, 'acl');
    return 1;
  }

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var workCtx = beginImportDriveWorkContext_(String(job.created_by || ''));
  // makeCopy тяжёлый — меньше за тик + промежуточный прогресс (не висим на 0/N).
  var copyChunkSize = Math.min(10, JOBS_CHUNK_SIZE_);
  var chunk = pending.slice(0, copyChunkSize);
  var processed = 0;

  for (var i = 0; i < chunk.length; i++) {
    if (isCatalogJobsPaused_()) {
      break;
    }
    var item = chunk[i];
    try {
      var catalogFile = longOpsTransferIntoCatalogRoot_(
        String(item.sourceFileId),
        'copy',
        catalogRootFolder
      );
      var mimeType = getDriveFileMimeType_(catalogFile) || '';
      patchFilesBatchRow_(workCtx.filesBatch, item.catalogId, {
        fileId: catalogFile.getId(),
        sizeBytes: resolveDriveFileSizeBytes_(catalogFile, mimeType),
        driveModifiedAt: catalogFile.getLastUpdated(),
        mimeType: mimeType,
        sourceFileId: String(item.sourceFileId),
        status: 'ready',
        lastError: ''
      });
      item.appliedMode = 'copy';
      item.copyDone = true;
      item.metaDone = true;
      item.error = '';
    } catch (eWork) {
      item.copyDone = true;
      item.metaDone = true;
      item.aclDone = true;
      item.error = (eWork && eWork.message) || String(eWork);
      longOpsMarkFileFailed_(item.catalogId, item.error, workCtx.filesBatch);
    }
    processed += 1;
    if (processed % 5 === 0) {
      payload.items = items;
      saveImportDriveCopyPhaseProgress_(job, payload, items, 'copy');
    }
  }

  commitFilesUpdateBatch_(workCtx.filesBatch);
  payload.items = items;
  saveImportDriveCopyPhaseProgress_(job, payload, items, 'copy');
  return processed;
}

/**
 * Фаза acl: права из CacheService (сняты в meta), без повторного list при наличии кэша.
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processImportDriveCopyAclPhase_(job, payload) {
  var items = payload.items || [];
  var createdBy = String(job.created_by || '');
  var pendingAcl = [];
  items.forEach(function (item) {
    if (!item.copyDone || item.aclDone || item.error) {
      return;
    }
    pendingAcl.push(item);
  });

  if (!pendingAcl.length) {
    saveImportDriveProgress_(job, payload, items, 0);
    return 1;
  }

  // DriveApp editors/viewers — без UrlFetch (дневная квота).
  var chunk = pendingAcl.slice(0, JOBS_CHUNK_SIZE_);
  var workCtx = beginImportDriveWorkContext_(createdBy);

  for (var i = 0; i < chunk.length; i++) {
    var item = chunk[i];
    var catalogId = String(item.catalogId || '');
    if (!catalogId) {
      item.aclDone = true;
      continue;
    }
    if (workCtx.aclTouchedIds) {
      workCtx.aclTouchedIds.push(catalogId);
    }
    try {
      var sourceFile = DriveApp.getFileById(String(item.sourceFileId || ''));
      applyDriveFileAclToCatalogFile_(sourceFile, catalogId, createdBy, workCtx);
      item.aclDone = true;
    } catch (eAcl) {
      item.aclDone = true;
      item.error = (eAcl && eAcl.message) || String(eAcl);
      longOpsMarkFileFailed_(catalogId, item.error, workCtx.filesBatch);
    }
  }

  commitImportDriveWorkContext_(workCtx);
  payload.items = items;
  saveImportDriveProgress_(job, payload, items, chunk.length);
  return chunk.length;
}

/**
 * @param {string} role
 * @returns {(string|null)} catalog level or null to skip (unknown)
 */
function drivePermissionRoleToCatalogLevel_(role) {
  var r = String(role || '');
  // Исходный owner на Drive → editor в каталоге (не технический CONTROLLER).
  if (r === 'owner') {
    return 'editor';
  }
  if (r === 'writer' || r === 'fileOrganizer' || r === 'organizer') {
    return 'editor';
  }
  if (r === 'commenter') {
    return 'commenter';
  }
  if (r === 'reader') {
    return 'reader';
  }
  return null;
}

/**
 * @param {string} jobId
 * @param {string} catalogId
 * @returns {string}
 */
function importSourcePermsCacheKey_(jobId, catalogId) {
  return (
    'impP:' +
    String(jobId || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20) +
    ':' +
    String(catalogId || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 40)
  );
}

/**
 * @param {Object.<string, string>} map
 */
function putImportSourcePermsCacheAll_(map) {
  var keys = Object.keys(map || {});
  if (!keys.length) {
    return;
  }
  var cache = CacheService.getDocumentCache();
  for (var off = 0; off < keys.length; off += 100) {
    var slice = {};
    keys.slice(off, off + 100).forEach(function (k) {
      slice[k] = map[k];
    });
    try {
      cache.putAll(slice, 21600);
    } catch (ePut) {
      keys.slice(off, off + 100).forEach(function (k) {
        try {
          cache.put(k, map[k], 21600);
        } catch (eOne) {
          // ignore
        }
      });
    }
  }
}

/**
 * @param {string} catalogId
 * @param {Array<{ email: string, displayName?: string, level: string }>} permsSlim
 * @param {string} createdBy
 * @param {Object} workCtx
 */
function applyCachedSourcePermsToCatalogFile_(catalogId, permsSlim, createdBy, workCtx) {
  if (!catalogId || !permsSlim || !permsSlim.length) {
    return;
  }
  var engine =
    (workCtx && workCtx.aclEngine) || createAclEngine_();
  if (workCtx) {
    workCtx.aclEngine = engine;
  }
  permsSlim.forEach(function (p) {
    var email = String((p && p.email) || '').trim();
    var level = String((p && p.level) || '').trim();
    if (!email || !level || level === 'none') {
      return;
    }
    if (shouldSkipRedundantFileAclPlus_(engine, catalogId, email, level)) {
      if (workCtx && workCtx.usersBatch) {
        ensureUserInBatch_(workCtx.usersBatch, {
          email: email,
          displayName: String((p && p.displayName) || email),
          addedBy: createdBy || ''
        });
      }
      return;
    }
    if (workCtx && workCtx.usersBatch) {
      ensureUserInBatch_(workCtx.usersBatch, {
        email: email,
        displayName: String((p && p.displayName) || email),
        addedBy: createdBy || ''
      });
      appendAclInBatch_(workCtx.aclBatch, 'file', catalogId, email, level);
    } else {
      appendOrEnsureUserRow_({
        email: email,
        loginRole: 'user',
        addedBy: createdBy || '',
        displayName: String((p && p.displayName) || email)
      });
      appendExplicitUserAclRow_('file', catalogId, email, level);
    }
  });
}

/**
 * Прогресс для meta/copy (счётчик по своей фазе).
 *
 * @param {Object} job
 * @param {Object} payload
 * @param {Array} items
 * @param {string} phase
 */
function saveImportDriveCopyPhaseProgress_(job, payload, items, phase) {
  var total = items.length;
  var chainTotal = Number(payload.chainTotalFiles) || total;
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var metaDone = 0;
  var copyDone = 0;
  var failed = 0;
  items.forEach(function (it) {
    if (it.error) {
      failed += 1;
    }
    if (it.metaDone || it.error) {
      metaDone += 1;
    }
    if (it.copyDone || it.error) {
      copyDone += 1;
    }
  });

  if (phase === 'meta') {
    var gMeta = chainBefore + metaDone;
    var msgMeta = 'Импорт (мета): ' + gMeta + '/' + chainTotal;
    var progMeta = chainTotal ? Math.round((gMeta / chainTotal) * 30) : 30;
    if (metaDone >= total) {
      payload.phase = 'copy';
    }
    saveJobPayloadProgress_(job.job_id, payload, progMeta, msgMeta, false);
    return;
  }

  if (phase === 'copy') {
    var gCopy = chainBefore + copyDone;
    var msgCopy = 'Импорт: ' + gCopy + '/' + chainTotal;
    var progCopy = chainTotal
      ? Math.min(90, 30 + Math.round((gCopy / chainTotal) * 50))
      : 80;
    if (copyDone >= total) {
      payload.phase = 'acl';
      msgCopy = 'Импорт (права)…';
    }
    saveJobPayloadProgress_(job.job_id, payload, progCopy, msgCopy, false);
    return;
  }

  saveImportDriveProgress_(job, payload, items, 0);
}

/**
 * @param {Object} job
 * @param {Object} payload
 * @param {Array} items
 * @param {number} processed
 */
function saveImportDriveProgress_(job, payload, items, processed) {
  var left = 0;
  var failed = 0;
  var done = 0;
  var aclLeft = 0;
  items.forEach(function (it) {
    if (it.copyDone) {
      done++;
    } else {
      left++;
    }
    if (it.copyDone && !it.aclDone && !it.error) {
      aclLeft++;
    }
    if (it.error) {
      failed++;
    }
  });

  var total = items.length;
  var chainTotal = Number(payload.chainTotalFiles) || total;
  var chainBefore = Number(payload.chainDoneBefore) || 0;
  var globalDone = chainBefore + done;
  var progress = chainTotal ? Math.round((globalDone / chainTotal) * 100) : 100;
  if (String(payload.phase) === 'acl' && done >= total) {
    progress = Math.min(
      99,
      50 + Math.round(((total - aclLeft) / Math.max(total, 1)) * 50)
    );
  }

  var fresh = getCatalogJobById_(job.job_id);
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
    return;
  }

  if (left === 0 && aclLeft === 0) {
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
    payload.phase = 'done';
    saveJobPayloadProgress_(job.job_id, payload, isLastPart ? 100 : progress, msg, true);
    return;
  }

  var statusMsg =
    String(payload.phase) === 'acl'
      ? 'Импорт (права): ' + (total - aclLeft) + '/' + total
      : 'Импорт: ' + globalDone + '/' + chainTotal;
  saveJobPayloadProgress_(job.job_id, payload, progress, statusMsg, false);
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
  var mimeType = getDriveFileMimeType_(catalogFile);
  patchFilesBatchRow_(workCtx.filesBatch, item.catalogId, {
    fileId: catalogFile.getId(),
    sizeBytes: resolveDriveFileSizeBytes_(catalogFile, mimeType),
    driveModifiedAt: catalogFile.getLastUpdated(),
    mimeType: mimeType,
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
  var mimeType = getDriveFileMimeType_(catalogFile);
  updateCatalogFileAfterDriveImport_(item.catalogId, {
    fileId: catalogFile.getId(),
    sizeBytes: resolveDriveFileSizeBytes_(catalogFile, mimeType),
    driveModifiedAt: catalogFile.getLastUpdated(),
    mimeType: mimeType,
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

  // Владелец на Drive → редактор в каталоге (§9.4); CONTROLLER не пишем.
  rememberDriveSourceOwnerAsEditor_(driveFile, levelsByEmail, namesByEmail, emailByKey);
  remember(driveFile.getEditors(), 'editor');
  try {
    remember(driveFile.getCommenters(), 'commenter');
  } catch (eC) {
    // optional
  }
  remember(driveFile.getViewers(), 'reader');

  var engineForSkip =
    (workCtx && workCtx.aclEngine) || createAclEngine_();
  if (workCtx) {
    workCtx.aclEngine = engineForSkip;
  }

  Object.keys(levelsByEmail).forEach(function (key) {
    var level = levelsByEmail[key];
    if (!level || level === 'none') {
      return;
    }
    var email = emailByKey[key] || key;
    var displayName = namesByEmail[key] || email;
    if (shouldSkipRedundantFileAclPlus_(engineForSkip, catalogId, email, level)) {
      if (workCtx && workCtx.usersBatch) {
        ensureUserInBatch_(workCtx.usersBatch, {
          email: email,
          displayName: displayName,
          addedBy: addedBy || ''
        });
        if (workCtx.aclTouchedIds) {
          workCtx.aclTouchedIds.push(String(catalogId));
        }
      }
      return;
    }
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
  } else if (workCtx.aclTouchedIds) {
    workCtx.aclTouchedIds.push(String(catalogId));
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
    setCatalogOpStatus_('Прерывание: останавливаем очередь Jobs…');
    var active = findActiveCatalogJob_();
    if (!active) {
      clearCatalogOpStatus_();
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
    var needsFirstLaunch = false;
    if (jobType === 'clear_catalog') {
      setCatalogOpStatus_('Прерывание очистки: добиваем корзину Drive…');
      var finished = finishClearCatalogNow_(payload);
      trashed = Number(finished.trashed) || 0;
      needsFirstLaunch = !!finished.needsFirstLaunch;
    } else if (jobType === 'import_drive' && chainId) {
      if (mode !== 'move') {
        setCatalogOpStatus_(
          'Прерывание импорта: перенос скопированного в корзину каталога…'
        );
        trashed = trashReadyFilesByImportChain_(chainId);
      } else {
        setCatalogOpStatus_('Прерывание импорта (перемещение): очистка очереди…');
      }
      setCatalogOpStatus_('Прерывание импорта: удаление незавершённых (pending)…');
      removedPending = deleteIncompleteFilesByImportChain_(chainId);
    } else {
      setCatalogOpStatus_('Прерывание операции…');
    }

    bumpCatalogRev_();
    clearCatalogOpStatus_();
    return {
      ok: true,
      cancelled: true,
      jobType: jobType,
      mode: mode,
      chainId: chainId,
      trashed: trashed,
      removedPending: removedPending,
      done: done,
      chainTotal: chainTotal,
      needsFirstLaunch: needsFirstLaunch
    };
  } catch (eCancel) {
    clearCatalogOpStatus_();
    throw eCancel;
  } finally {
    lock.releaseLock();
  }
}

/**
 * @param {string} message
 */
function setCatalogOpStatus_(message) {
  PropertiesService.getDocumentProperties().setProperty(
    PROP_CATALOG_OP_STATUS_,
    String(message || '')
  );
}

function clearCatalogOpStatus_() {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_CATALOG_OP_STATUS_);
}

/**
 * Статус долгой операции (прерывание импорта) для poll UI.
 * @returns {{ ok: true, message: string }}
 */
function getCatalogOpStatus() {
  var msg =
    PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_OP_STATUS_) ||
    '';
  return { ok: true, message: String(msg) };
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
        var rawPl = String(values[r][payloadCol] || '{}');
        // Не JSON.parse огромных битых payload целиком, если нет chainId в префиксе —
        // но chainId обычно в начале объекта. Парсим осторожно.
        if (rawPl.indexOf(chainId) >= 0) {
          var pl = JSON.parse(rawPl);
          if (pl && String(pl.chainId || '') === chainId) {
            match = true;
          }
        }
      } catch (eParse) {
        // ignore
      }
    }
    if (!match && fallbackJobId && String(values[r][idCol]) === String(fallbackJobId)) {
      match = true;
    }
    if (!match) {
      continue;
    }
    // Пишем только нужные ячейки — не setValues всей строки/листа
    // (иначе падаем на payload > 50000 при перезаписи).
    sheet.getRange(r + 1, statusCol + 1).setValue('cancelled');
    if (completedCol >= 0) {
      sheet.getRange(r + 1, completedCol + 1).setValue(now);
    }
    if (msgCol >= 0) {
      sheet.getRange(r + 1, msgCol + 1).setValue('Прервано');
    }
  }
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
  setCatalogOpStatus_(
    'Прерывание импорта: в корзину каталога — ' + updates.length + ' файл(ов)…'
  );
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
