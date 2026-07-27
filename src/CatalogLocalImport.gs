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

  var createdFolders = [];
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
      treeSheet.appendRow([folderId, parentId, parts[i], now, false, '', '', '']);
      pathToFolderId[accum] = folderId;
      createdFolders.push({ folderId: folderId, parentId: parentId, name: parts[i] });
      parentId = folderId;
      created++;
    }
  });

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
    folderCount: created
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
    return {
      index: index,
      relativePath: relativePath,
      fileName: fileName,
      parentFolderId: parentFolderId,
      sizeBytes: parseNumber_(f.sizeBytes),
      mimeType: String(f.mimeType || 'application/octet-stream'),
      status: 'waiting',
      catalogId: '',
      error: ''
    };
  });

  var jobId = enqueueCatalogJob_(
    'import_upload',
    {
      targetFolderId: targetFolderId,
      items: items
    },
    userEmail
  );
  ensureCatalogJobsTrigger_();
  markCatalogJobRunning_(jobId);
  patchCatalogJobRow_(jobId, {
    progress: 0,
    progress_message: 'Загрузка: 0/' + items.length
  });

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    pathToFolderId: pathToFolderId,
    folderCount: tree.folderCount || 0,
    fileCount: items.length
  };
}

/**
 * Загружает один локальный файл в плоскую папку каталога на Drive + Files.
 * При jobId — часть очереди import_upload.
 *
 * @param {{
 *   parentFolderId: string,
 *   fileName: string,
 *   mimeType?: string,
 *   base64Data: string,
 *   sizeBytes?: number,
 *   jobId?: string,
 *   relativePath?: string
 * }} input
 * @returns {{
 *   ok: true,
 *   catalogId: string,
 *   displayName: string,
 *   fileId: string,
 *   sizeBytes: number,
 *   jobId?: string
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
  var jobId = String(input.jobId || '').trim();
  var relativePath = String(input.relativePath || '').trim();

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

  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Data);
  } catch (e) {
    if (jobId) {
      markLocalImportItem_(jobId, relativePath, fileName, 'failed', '', (e && e.message) || 'decode');
    }
    throw catalogError_('INVALID_INPUT', 'Не удалось разобрать содержимое файла.');
  }

  if (!sizeBytes || sizeBytes < 0) {
    sizeBytes = bytes.length;
  }

  try {
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
      mimeType: resolvedMime,
      status: 'ready'
    });

    engine.filesByCatalogId[catalogId] = {
      catalog_id: catalogId,
      folder_id: parentFolderId,
      approved: false
    };
    copyExplicitAclFromParentFolder_(engine, 'file', catalogId, parentFolderId);

    if (jobId) {
      markLocalImportItem_(jobId, relativePath, fileName, 'done', catalogId, '');
    }

    return {
      ok: true,
      catalogId: catalogId,
      displayName: fileName,
      fileId: driveFile.getId(),
      sizeBytes: sizeBytes,
      jobId: jobId || undefined
    };
  } catch (eUp) {
    if (jobId) {
      markLocalImportItem_(
        jobId,
        relativePath,
        fileName,
        'failed',
        '',
        (eUp && eUp.message) || String(eUp)
      );
    }
    throw eUp;
  }
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
  var job = getCatalogJobById_(jobId);
  if (!job) {
    return;
  }
  var payload = parseJobPayload_(job);
  var items = payload.items || [];
  var matched = false;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (it.status === 'done' || it.status === 'failed') {
      continue;
    }
    if (
      (relativePath && it.relativePath === relativePath) ||
      (!relativePath && it.fileName === fileName)
    ) {
      it.status = status;
      it.catalogId = catalogId || '';
      it.error = error || '';
      matched = true;
      break;
    }
  }
  if (!matched) {
    return;
  }
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
  var progress = Math.round((done / items.length) * 100);
  var finished = done >= items.length;
  var msg = finished
    ? failed
      ? 'Загрузка завершена с ошибками: ' + failed
      : 'Загрузка с компьютера завершена'
    : 'Загрузка с компьютера: ' + done + '/' + items.length;
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
