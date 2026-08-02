/**
 * §9.13 — импорт с локального диска (браузер → Drive + Tree/Files).
 * Только загрузка; структура папок — виртуальная в Tree.
 */

/** @const {number} Порог предупреждения на клиенте (байт); сервер не режет. */
var LOCAL_IMPORT_WARN_BYTES_ = 10 * 1024 * 1024;

/** @const {number} Макс. файлов в одном RPC upload-пакете (§0.4a / §9.13). */
var LOCAL_IMPORT_UPLOAD_BATCH_FILES_ = 5;

/** @const {number} Бюджет сырого размера файлов в одном пакете (байт). */
var LOCAL_IMPORT_UPLOAD_BATCH_BYTES_ = 1.5 * 1024 * 1024;

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

  var pathToFolderId = {};
  var now = new Date();
  var created = 0;

  var createdFolders = [];
  var treeBatch = [];
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
      treeBatch.push({
        folderId: folderId,
        parentFolderId: parentId,
        name: parts[i],
        folderCreatedAt: now,
        isSystem: false
      });
      pathToFolderId[accum] = folderId;
      createdFolders.push({ folderId: folderId, parentId: parentId, name: parts[i] });
      parentId = folderId;
      created++;
    }
  });
  if (treeBatch.length) {
    appendTreeFolderRowsBatch_(treeBatch);
  }
  if (createdFolders.length) {
    var engine = createAclEngine_();
    createdFolders.forEach(function (f) {
      engine.foldersById[f.folderId] = {
        folder_id: f.folderId,
        parent_folder_id: f.parentId,
        name: f.name,
        is_system: false
      };
      copyExplicitAclFromParentFolder_(engine, 'folder', f.folderId, f.parentId);
    });
  }

  return {
    ok: true,
    pathToFolderId: pathToFolderId,
    folderCount: created,
    createdFolders: createdFolders.map(function (f) {
      return {
        id: f.folderId,
        parentFolderId: f.parentId,
        name: f.name,
        pendingImport: true
      };
    })
  };
}

/**
 * §9.13 — старт локального импорта как Jobs-очереди (байтты шлёт клиент).
 *
 * @param {{
 *   targetFolderId: string,
 *   folderPaths: string[],
 *   files: Array<{ relativePath: string, fileName: string, sizeBytes?: number, mimeType?: string }>
 * }} input
 * @returns {{
 *   ok: true,
 *   queued: true,
 *   jobId: string,
 *   pathToFolderId: Object.<string, string>,
 *   folderCount: number,
 *   fileCount: number
 * }}
 */
function startLocalImportJob(input) {
  assertCatalogReady_();
  assertNoActiveCatalogJobs_();

  input = input || {};
  var files = Array.isArray(input.files) ? input.files : [];
  if (!files.length) {
    throw catalogError_('INVALID_INPUT', 'Нет файлов для загрузки.');
  }

  var tree = prepareLocalImportTree({
    targetFolderId: input.targetFolderId,
    folderPaths: input.folderPaths || []
  });

  var userEmail = Session.getActiveUser().getEmail() || '';
  var targetFolderId = String(input.targetFolderId || '').trim();
  var pathToFolderId = tree.pathToFolderId || {};

  var engine = createAclEngine_();
  var pendingRows = [];
  var createdFiles = [];
  var items = files.map(function (f, index) {
    var relativePath = String(f.relativePath || f.fileName || '').trim();
    var fileName = String(f.fileName || '').trim();
    if (!fileName && relativePath) {
      var slash = relativePath.lastIndexOf('/');
      fileName = slash >= 0 ? relativePath.substring(slash + 1) : relativePath;
    }
    var dir = '';
    if (relativePath.indexOf('/') >= 0) {
      dir = relativePath.substring(0, relativePath.lastIndexOf('/'));
    }
    var parentFolderId = dir ? pathToFolderId[dir] || targetFolderId : targetFolderId;
    var catalogId = Utilities.getUuid();
    var sizeBytes = parseNumber_(f.sizeBytes);
    var mimeType = String(f.mimeType || 'application/octet-stream');
    pendingRows.push({
      catalogId: catalogId,
      folderId: parentFolderId,
      fileId: '',
      displayName: fileName,
      sizeBytes: sizeBytes,
      driveModifiedAt: '',
      sourceFileId: '',
      mimeType: mimeType,
      status: 'pending'
    });
    createdFiles.push({
      id: catalogId,
      folderId: parentFolderId,
      name: fileName,
      mimeType: mimeType,
      sizeBytes: sizeBytes,
      status: 'pending'
    });
    return {
      index: index,
      relativePath: relativePath,
      fileName: fileName,
      parentFolderId: parentFolderId,
      sizeBytes: sizeBytes,
      mimeType: mimeType,
      status: 'waiting',
      catalogId: catalogId,
      error: ''
    };
  });

  if (pendingRows.length) {
    appendCatalogFileRowsBatch_(pendingRows);
    pendingRows.forEach(function (row) {
      engine.filesByCatalogId[row.catalogId] = {
        catalog_id: row.catalogId,
        folder_id: row.folderId,
        approved: false
      };
      copyExplicitAclFromParentFolder_(engine, 'file', row.catalogId, row.folderId);
    });
  }

  var jobId = enqueueCatalogJob_(
    'import_upload',
    {
      targetFolderId: targetFolderId,
      items: items
    },
    userEmail
  );
  setCatalogJobsPaused_(false);
  ensureCatalogJobsTrigger_();
  markCatalogJobRunning_(jobId);
  patchCatalogJobRow_(jobId, {
    progress: 0,
    progress_message: 'Загрузка: 0/' + items.length
  });
  bumpCatalogRev_();
  kickCatalogJobsProcessing_();

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    pathToFolderId: pathToFolderId,
    folderCount: tree.folderCount || 0,
    fileCount: items.length,
    created: {
      folders: tree.createdFolders || [],
      files: createdFiles
    }
  };
}

/**
 * Загружает один локальный файл — тонкая обёртка над пакетом.
 *
 * @param {Object} input
 * @returns {Object}
 */
function importLocalFile(input) {
  var batch = importLocalFilesBatch({
    jobId: input && input.jobId,
    files: [input || {}]
  });
  var first = (batch.results && batch.results[0]) || null;
  if (!first || !first.ok) {
    throw catalogError_(
      (first && first.code) || 'IMPORT_FAILED',
      (first && first.error) || 'Не удалось загрузить файл.'
    );
  }
  return {
    ok: true,
    catalogId: first.catalogId,
    displayName: first.displayName,
    fileId: first.fileId,
    sizeBytes: first.sizeBytes,
    jobId: batch.jobId
  };
}

/**
 * §0.4a / §9.13 — пакетный upload с клиента (фаза place).
 * Один вызов: один ACL-engine, один batch Files, один апдейт Jobs.
 *
 * @param {{
 *   jobId?: string,
 *   files: Array<{
 *     parentFolderId: string,
 *     fileName: string,
 *     mimeType?: string,
 *     base64Data: string,
 *     sizeBytes?: number,
 *     relativePath?: string,
 *     catalogId?: string
 *   }>
 * }} input
 * @returns {{
 *   ok: true,
 *   jobId?: string,
 *   doneCount: number,
 *   failedCount: number,
 *   results: Array<Object>
 * }}
 */
function importLocalFilesBatch(input) {
  assertCatalogReady_();
  input = input || {};
  var files = Array.isArray(input.files) ? input.files : [];
  if (!files.length) {
    throw catalogError_('INVALID_INPUT', 'Нет файлов для загрузки.');
  }
  if (files.length > LOCAL_IMPORT_UPLOAD_BATCH_FILES_) {
    throw catalogError_(
      'INVALID_INPUT',
      'Слишком много файлов в пакете (макс. ' + LOCAL_IMPORT_UPLOAD_BATCH_FILES_ + ').'
    );
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var jobId = String(input.jobId || '').trim();
  if (jobId) {
    var job = getCatalogJobById_(jobId);
    if (!job || String(job.job_type) !== 'import_upload') {
      throw catalogError_('JOB_NOT_FOUND', 'Задача локального импорта не найдена.');
    }
    var st = String(job.status || '').toLowerCase();
    if (st !== 'running' && st !== 'pending') {
      throw catalogError_('JOB_NOT_ACTIVE', 'Очередь импорта уже завершена.');
    }
  }

  var engine = createAclEngine_();
  var checkedParents = {};
  files.forEach(function (f) {
    var parentFolderId = String((f && f.parentFolderId) || '').trim();
    if (!parentFolderId) {
      throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
    }
    if (checkedParents[parentFolderId]) {
      return;
    }
    if (!engine.foldersById[parentFolderId]) {
      throw catalogError_('FOLDER_NOT_FOUND', 'Parent folder not found: ' + parentFolderId);
    }
    assertEditorOnFolderForMove_(engine, userEmail, loginRole, parentFolderId);
    checkedParents[parentFolderId] = true;
  });

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var filesBatch = beginFilesUpdateBatch_();
  var results = [];
  var marks = [];
  var failedCount = 0;

  for (var i = 0; i < files.length; i++) {
    var f = files[i] || {};
    var parentFolderId = String(f.parentFolderId || '').trim();
    var fileName = String(f.fileName || '').trim();
    var mimeType = String(f.mimeType || 'application/octet-stream').trim();
    var relativePath = String(f.relativePath || '').trim();
    var catalogId = String(f.catalogId || '').trim();
    var sizeBytes = parseNumber_(f.sizeBytes);
    var base64Data = stripLocalImportBase64Prefix_(String(f.base64Data || ''));

    if (!fileName || !base64Data) {
      failedCount++;
      results.push({
        ok: false,
        code: 'INVALID_INPUT',
        error: 'fileName/base64Data required',
        relativePath: relativePath,
        catalogId: catalogId
      });
      if (jobId) {
        marks.push({
          relativePath: relativePath,
          fileName: fileName,
          status: 'failed',
          catalogId: catalogId,
          error: 'fileName/base64Data required'
        });
      }
      continue;
    }

    try {
      var bytes = Utilities.base64Decode(base64Data);
      if (!sizeBytes || sizeBytes < 0) {
        sizeBytes = bytes.length;
      }
      var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
      var driveFile = catalogRootFolder.createFile(blob);
      if (controllerEmail) {
        transferDriveFileToController_(driveFile, controllerEmail, userEmail);
      }
      if (!catalogId && jobId) {
        catalogId = findLocalImportCatalogIdFromPayloadItems_(
          jobId,
          relativePath,
          fileName
        );
      }
      var resolvedMime = '';
      try {
        resolvedMime = String(driveFile.getMimeType() || mimeType || '');
      } catch (eMime) {
        resolvedMime = mimeType || '';
      }

      if (catalogId && engine.filesByCatalogId[catalogId]) {
        patchFilesBatchRow_(filesBatch, catalogId, {
          fileId: driveFile.getId(),
          sizeBytes: sizeBytes,
          mimeType: resolvedMime,
          status: 'ready',
          driveModifiedAt: driveFile.getLastUpdated()
        });
      } else {
        catalogId = catalogId || Utilities.getUuid();
        appendCatalogFileRow_({
          catalogId: catalogId,
          folderId: parentFolderId,
          fileId: driveFile.getId(),
          displayName: fileName,
          sizeBytes: sizeBytes,
          driveModifiedAt: driveFile.getLastUpdated(),
          sourceFileId: '',
          mimeType: resolvedMime,
          status: 'ready'
        });
        engine.filesByCatalogId[catalogId] = {
          catalog_id: catalogId,
          folder_id: parentFolderId,
          approved: false
        };
        copyExplicitAclFromParentFolder_(engine, 'file', catalogId, parentFolderId);
      }

      results.push({
        ok: true,
        catalogId: catalogId,
        displayName: fileName,
        fileId: driveFile.getId(),
        sizeBytes: sizeBytes,
        relativePath: relativePath
      });
      if (jobId) {
        marks.push({
          relativePath: relativePath,
          fileName: fileName,
          status: 'done',
          catalogId: catalogId,
          error: ''
        });
      }
    } catch (eUp) {
      failedCount++;
      var errMsg = (eUp && eUp.message) || String(eUp);
      results.push({
        ok: false,
        code: (eUp && eUp.name) || 'IMPORT_FAILED',
        error: errMsg,
        relativePath: relativePath,
        catalogId: catalogId
      });
      if (jobId) {
        marks.push({
          relativePath: relativePath,
          fileName: fileName,
          status: 'failed',
          catalogId: catalogId,
          error: errMsg
        });
      }
    }
  }

  commitFilesUpdateBatch_(filesBatch);
  if (jobId && marks.length) {
    markLocalImportItemsBatch_(jobId, marks);
  }
  bumpCatalogRev_();

  return {
    ok: true,
    jobId: jobId || undefined,
    doneCount: results.length - failedCount,
    failedCount: failedCount,
    results: results
  };
}

/**
 * @param {string} raw
 * @returns {string}
 */
function stripLocalImportBase64Prefix_(raw) {
  var base64Data = String(raw || '').trim();
  var comma = base64Data.indexOf(',');
  if (base64Data.indexOf('base64') >= 0 && comma >= 0) {
    return base64Data.substring(comma + 1);
  }
  return base64Data;
}

/**
 * @param {string} jobId
 * @param {string} relativePath
 * @param {string} fileName
 * @returns {string}
 */
function findLocalImportCatalogId_(jobId, relativePath, fileName) {
  return findLocalImportCatalogIdFromPayloadItems_(jobId, relativePath, fileName);
}

/**
 * @param {string} jobId
 * @param {string} relativePath
 * @param {string} fileName
 * @returns {string}
 */
function findLocalImportCatalogIdFromPayloadItems_(jobId, relativePath, fileName) {
  var job = getCatalogJobById_(jobId);
  if (!job) {
    return '';
  }
  var items = parseJobPayload_(job).items || [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (
      (relativePath && it.relativePath === relativePath) ||
      (!relativePath && it.fileName === fileName)
    ) {
      return String(it.catalogId || '');
    }
  }
  return '';
}

/**
 * @param {string} jobId
 * @param {string} relativePath
 * @param {string} fileName
 * @param {'done'|'failed'} status
 * @param {string} catalogId
 * @param {string} error
 */
function markLocalImportItem_(jobId, relativePath, fileName, status, catalogId, error) {
  markLocalImportItemsBatch_(jobId, [
    {
      relativePath: relativePath,
      fileName: fileName,
      status: status,
      catalogId: catalogId,
      error: error
    }
  ]);
}

/**
 * @param {string} jobId
 * @param {Array<{ relativePath?: string, fileName?: string, status: string, catalogId?: string, error?: string }>} marks
 */
function markLocalImportItemsBatch_(jobId, marks) {
  if (!jobId || !marks || !marks.length) {
    return;
  }
  var job = getCatalogJobById_(jobId);
  if (!job) {
    return;
  }
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  marks.forEach(function (m) {
    var relativePath = String(m.relativePath || '');
    var fileName = String(m.fileName || '');
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (it.status === 'done' || it.status === 'failed') {
        continue;
      }
      if (
        (relativePath && it.relativePath === relativePath) ||
        (!relativePath && it.fileName === fileName) ||
        (m.catalogId && String(it.catalogId || '') === String(m.catalogId))
      ) {
        it.status = m.status;
        it.catalogId = m.catalogId || it.catalogId || '';
        it.error = m.error || '';
        break;
      }
    }
  });
  payload.items = items;
  var done = 0;
  var failed = 0;
  items.forEach(function (x) {
    if (x.status === 'done') {
      done++;
    } else if (x.status === 'failed') {
      failed++;
      done++;
    }
  });
  var progress = Math.round((done / Math.max(items.length, 1)) * 100);
  var finished = done >= items.length;
  var msg = finished
    ? failed
      ? 'Загрузка завершена с ошибками: ' + failed
      : 'Загрузка с компьютера завершена'
    : 'Загрузка: ' + done + '/' + items.length;
  saveJobPayloadProgress_(jobId, payload, progress, msg, finished);
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
