/**
 * §24.4 — архивация каталога (Управляющий).
 * Создаёт иерархию папок на Drive под выбранным пустым корнем,
 * переносит туда файлы (не ярлыки), в каталоге оставляет пустую структуру.
 */

/**
 * @param {{ targetFolderId: string }} input
 * @returns {{
 *   ok: true,
 *   queued: boolean,
 *   jobId: (string|undefined),
 *   folderCount: number,
 *   fileCount: number,
 *   shortcutCount: number,
 *   archiveRootId: string,
 *   archiveUrl: string
 * }}
 */
function archiveCatalog(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  input = input || {};

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var targetFolderId = String(input.targetFolderId || '').trim();
  if (!targetFolderId || targetFolderId === 'root') {
    throw catalogError_(
      'INVALID_INPUT',
      'Выберите пустую папку на Моём диске (не корень).'
    );
  }

  assertDriveFolderAccessible_(targetFolderId);
  assertArchiveTargetEmpty_(targetFolderId);
  assertArchiveTargetNotCatalogRoot_(targetFolderId);

  var plan = buildArchiveFolderPlan_(targetFolderId);
  var fileStats = countArchiveFileStats_();

  if (!plan.folderTasks.length && !fileStats.moveCount && !fileStats.shortcutCount) {
    clearCatalogSheetDataRows_('Files');
    removeAclForTrashObjects_(collectAllFileCatalogIds_(), []);
    bumpCatalogRev_();
    return {
      ok: true,
      queued: false,
      folderCount: 0,
      fileCount: 0,
      shortcutCount: 0,
      archiveRootId: targetFolderId,
      archiveUrl: driveFolderUrl_(targetFolderId)
    };
  }

  var jobId = enqueueCatalogJob_(
    'archive_catalog',
    {
      phase: plan.folderTasks.length ? 'folders' : 'files',
      archiveRootId: targetFolderId,
      folderMap: plan.folderMap,
      folderTasks: plan.folderTasks,
      foldersDone: 0,
      foldersTotal: plan.folderTasks.length,
      filesDone: 0,
      filesTotal: fileStats.moveCount,
      shortcutsTotal: fileStats.shortcutCount,
      moveErrors: 0
    },
    userEmail,
    ''
  );
  kickCatalogJobsProcessing_();

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    folderCount: plan.folderTasks.length,
    fileCount: fileStats.moveCount,
    shortcutCount: fileStats.shortcutCount,
    archiveRootId: targetFolderId,
    archiveUrl: driveFolderUrl_(targetFolderId)
  };
}

/**
 * @param {string} folderId
 */
function assertArchiveTargetEmpty_(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  if (folder.getFiles().hasNext() || folder.getFolders().hasNext()) {
    throw catalogError_(
      'INVALID_INPUT',
      'Папка архива должна быть пустой.'
    );
  }
}

/**
 * @param {string} folderId
 */
function assertArchiveTargetNotCatalogRoot_(folderId) {
  var catalogRoot = '';
  try {
    catalogRoot = getCatalogRootFolderId_();
  } catch (e) {
    return;
  }
  if (catalogRoot && folderId === catalogRoot) {
    throw catalogError_(
      'INVALID_INPUT',
      'Нельзя архивировать в плоскую папку каталога.'
    );
  }
}

/**
 * @param {string} folderId
 * @returns {string}
 */
function driveFolderUrl_(folderId) {
  return 'https://drive.google.com/drive/folders/' + folderId;
}

/**
 * Реальные папки Tree (без синих/красных ярлыков) → задачи создания на Drive.
 * Корень архива = targetFolderId (виртуальный корень каталога туда мапится).
 *
 * @param {string} archiveRootId
 * @returns {{
 *   folderMap: Object.<string, string>,
 *   folderTasks: Array<{ folderId: string, parentId: string, name: string, done: boolean, driveId: string }>
 * }}
 */
function buildArchiveFolderPlan_(archiveRootId) {
  var props = PropertiesService.getDocumentProperties();
  var virtualRoot =
    props.getProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_) || '';
  var rows = readSheetRecords_('Tree');
  var byId = {};
  rows.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (id) {
      byId[id] = row;
    }
  });
  if (!virtualRoot) {
    rows.forEach(function (row) {
      var id = String(row.folder_id || '').trim();
      var parent = String(row.parent_folder_id || '').trim();
      if (id && !parent && id !== '__TRASH__') {
        virtualRoot = id;
      }
    });
  }

  var folderMap = {};
  if (virtualRoot) {
    folderMap[virtualRoot] = archiveRootId;
  }

  var tasks = [];
  var queued = {};
  var queue = [];
  if (virtualRoot) {
    queue.push(virtualRoot);
    queued[virtualRoot] = true;
  }

  while (queue.length) {
    var parentId = queue.shift();
    rows.forEach(function (row) {
      var id = String(row.folder_id || '').trim();
      if (!id || queued[id]) {
        return;
      }
      if (String(row.parent_folder_id || '').trim() !== parentId) {
        return;
      }
      if (isMirrorFolderRow_(row)) {
        return;
      }
      tasks.push({
        folderId: id,
        parentId: parentId,
        name: String(row.name || id),
        done: false,
        driveId: ''
      });
      queue.push(id);
      queued[id] = true;
    });
  }

  // Реальные папки, недостижимые из корня (родитель — ярлык и т.п.) — под корень архива.
  rows.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (!id || queued[id] || id === virtualRoot) {
      return;
    }
    if (isMirrorFolderRow_(row)) {
      return;
    }
    tasks.push({
      folderId: id,
      parentId: virtualRoot || '',
      name: String(row.name || id),
      done: false,
      driveId: ''
    });
    queued[id] = true;
  });

  return { folderMap: folderMap, folderTasks: tasks };
}

/**
 * @returns {{ moveCount: number, shortcutCount: number }}
 */
function countArchiveFileStats_() {
  var moveCount = 0;
  var shortcutCount = 0;
  readSheetRecords_('Files').forEach(function (row) {
    if (isFileShortcutRow_(row)) {
      shortcutCount += 1;
    } else if (String(row.file_id || '').trim()) {
      moveCount += 1;
    }
  });
  return { moveCount: moveCount, shortcutCount: shortcutCount };
}

/**
 * @returns {string[]}
 */
function collectAllFileCatalogIds_() {
  return readSheetRecords_('Files')
    .map(function (row) {
      return String(row.catalog_id || '').trim();
    })
    .filter(Boolean);
}

/**
 * Воркер Jobs `archive_catalog`.
 * В одном вызове крутит фазы, пока есть бюджет времени (~4.5 мин).
 *
 * @param {Object} job
 * @returns {number}
 */
function processArchiveCatalogJobChunk_(job) {
  var startMs = Date.now();
  var budgetMs = 270000;
  var totalProcessed = 0;

  while (Date.now() - startMs < budgetMs) {
    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    var payload = parseJobPayload_(job);
    var phase = String(payload.phase || 'folders');
    var n = 0;
    if (phase === 'folders') {
      n = processArchiveFoldersChunk_(job, payload);
    } else if (phase === 'files') {
      n = processArchiveFilesChunk_(job, payload);
    } else if (phase === 'finalize') {
      n = processArchiveFinalizeChunk_(job, payload);
    } else {
      break;
    }
    totalProcessed += n || 0;

    job = getCatalogJobById_(job.job_id) || job;
    if (!job || String(job.status || '').toLowerCase() === 'done') {
      break;
    }
    payload = parseJobPayload_(job);
    if (String(payload.phase || '') === 'done') {
      break;
    }
    // Если чанк ничего не сделал — не крутиться вхолостую.
    if (!n) {
      break;
    }
  }
  return totalProcessed;
}

/**
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processArchiveFoldersChunk_(job, payload) {
  var tasks = payload.folderTasks || [];
  var folderMap = payload.folderMap || {};
  var processed = 0;
  var archiveRootId = String(payload.archiveRootId || '');
  var parentCache = {};

  for (var i = 0; i < tasks.length && processed < JOBS_CHUNK_SIZE_; i++) {
    var task = tasks[i];
    if (task.done) {
      continue;
    }
    var parentDriveId =
      folderMap[task.parentId] || archiveRootId || payload.archiveRootId;
    if (!parentDriveId) {
      task.done = true;
      task.driveId = '';
      processed += 1;
      continue;
    }
    try {
      var parentFolder = parentCache[parentDriveId];
      if (!parentFolder) {
        parentFolder = DriveApp.getFolderById(parentDriveId);
        parentCache[parentDriveId] = parentFolder;
      }
      var created = parentFolder.createFolder(String(task.name || task.folderId));
      task.driveId = created.getId();
      folderMap[task.folderId] = task.driveId;
      task.grantReadersPending = true;
    } catch (eCreate) {
      task.driveId = '';
      payload.moveErrors = (Number(payload.moveErrors) || 0) + 1;
    }
    task.done = true;
    processed += 1;
  }

  // Права каталога на папку → reader на Drive-папке архива (§24.4 п.5).
  var folderGrants = [];
  tasks.forEach(function (t) {
    if (t.grantReadersPending && t.driveId && t.folderId) {
      folderGrants.push({
        driveId: String(t.driveId),
        objectType: 'folder',
        objectId: String(t.folderId)
      });
      t.grantReadersPending = false;
    }
  });
  if (folderGrants.length) {
    grantCatalogAclAsDriveReadersBatch_(folderGrants, ScriptApp.getOAuthToken());
  }

  payload.folderMap = folderMap;
  var foldersDone = 0;
  tasks.forEach(function (t) {
    if (t.done) {
      foldersDone += 1;
    }
  });
  payload.foldersDone = foldersDone;
  payload.folderTasks = tasks;

  if (foldersDone < tasks.length) {
    var pct =
      tasks.length > 0 ? (foldersDone / tasks.length) * 40 : 40;
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      pct,
      'Архивация (папки): ' + foldersDone + '/' + tasks.length,
      false
    );
    return processed;
  }

  payload.phase = 'files';
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    40,
    'Архивация (файлы): 0/' + (payload.filesTotal || 0),
    false
  );
  return processed || 1;
}

/**
 * Перенос файлов пакетом (Drive API fetchAll) + demote прав пакетом.
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processArchiveFilesChunk_(job, payload) {
  var folderMap = payload.folderMap || {};
  var archiveRootId = String(payload.archiveRootId || '');
  var skip = payload.skipCatalogIds || {};
  var rows = readSheetRecords_('Files');
  var toMove = [];
  rows.forEach(function (row) {
    if (isFileShortcutRow_(row)) {
      return;
    }
    var fileId = String(row.file_id || '').trim();
    var catalogId = String(row.catalog_id || '').trim();
    if (!fileId || !catalogId || skip[catalogId]) {
      return;
    }
    toMove.push(row);
  });

  if (!toMove.length) {
    payload.phase = 'finalize';
    saveJobPayloadProgress_(
      job.job_id,
      payload,
      90,
      'Архивация: завершение…',
      false
    );
    return 1;
  }

  var catalogRootId = '';
  try {
    catalogRootId = getCatalogRootFolderId_();
  } catch (eRoot) {
    catalogRootId = '';
  }

  var chunk = toMove.slice(0, JOBS_CHUNK_SIZE_);
  var token = ScriptApp.getOAuthToken();
  var moveReqs = [];
  var meta = [];

  chunk.forEach(function (row) {
    var catalogId = String(row.catalog_id || '').trim();
    var fileId = String(row.file_id || '').trim();
    var folderId = String(row.folder_id || '').trim();
    var destId = folderMap[folderId] || archiveRootId;
    meta.push({
      catalogId: catalogId,
      fileId: fileId,
      destId: destId
    });
    var url =
      'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '?supportsAllDrives=true&fields=id';
    if (destId) {
      url += '&addParents=' + encodeURIComponent(destId);
    }
    if (catalogRootId && catalogRootId !== destId) {
      url += '&removeParents=' + encodeURIComponent(catalogRootId);
    }
    moveReqs.push({
      url: url,
      method: 'patch',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: '{}',
      muteHttpExceptions: true
    });
  });

  var moveResps = moveReqs.length ? UrlFetchApp.fetchAll(moveReqs) : [];
  var removedCatalogIds = [];
  var demoteFileIds = [];

  for (var i = 0; i < meta.length; i++) {
    var code = 0;
    try {
      code = moveResps[i] ? moveResps[i].getResponseCode() : 0;
    } catch (eCode) {
      code = 0;
    }
    if (code >= 200 && code < 300) {
      removedCatalogIds.push(meta[i].catalogId);
      demoteFileIds.push(meta[i].fileId);
    } else {
      payload.moveErrors = (Number(payload.moveErrors) || 0) + 1;
      skip[meta[i].catalogId] = true;
    }
  }

  payload.skipCatalogIds = skip;

  // Сначала выдать reader по ACL каталога (истина прав), затем demote writer→reader,
  // потом снести Files/ACL файла из каталога. Иначе на Drive остаётся только владелец.
  var grantPairs = [];
  for (var g = 0; g < meta.length; g++) {
    if (removedCatalogIds.indexOf(meta[g].catalogId) < 0) {
      continue;
    }
    grantPairs.push({
      driveId: String(meta[g].fileId || ''),
      objectType: 'file',
      objectId: String(meta[g].catalogId || '')
    });
  }
  if (grantPairs.length) {
    grantCatalogAclAsDriveReadersBatch_(grantPairs, token);
  }

  if (demoteFileIds.length) {
    demoteDriveFilesCollaboratorsBatch_(demoteFileIds, token);
  }

  if (removedCatalogIds.length) {
    deleteArchiveFileRowsByCatalogIds_(removedCatalogIds);
  }

  payload.filesDone =
    (Number(payload.filesDone) || 0) + removedCatalogIds.length;
  var filesTotal = Number(payload.filesTotal) || payload.filesDone;
  var folderPct = 40;
  var filePct =
    filesTotal > 0 ? Math.min(50, (payload.filesDone / filesTotal) * 50) : 50;
  saveJobPayloadProgress_(
    job.job_id,
    payload,
    folderPct + filePct,
    'Архивация (файлы): ' + payload.filesDone + '/' + filesTotal,
    false
  );
  return meta.length;
}

/**
 * Удалить оставшиеся Files (ярлыки) + ACL файлов; job done.
 *
 * @param {Object} job
 * @param {Object} payload
 * @returns {number}
 */
function processArchiveFinalizeChunk_(job, payload) {
  var leftoverIds = collectAllFileCatalogIds_();
  clearCatalogSheetDataRows_('Files');
  if (leftoverIds.length) {
    removeAclForTrashObjects_(leftoverIds, []);
  } else {
    // на случай ACL-сирот по файлам — не трогаем folder ACL
  }

  var msg = 'Архивация завершена';
  if (Number(payload.moveErrors) > 0) {
    msg += ' (ошибок: ' + payload.moveErrors + ')';
  }
  payload.phase = 'done';
  saveJobPayloadProgress_(job.job_id, payload, 100, msg, true);
  return 1;
}

/**
 * @param {string[]} catalogIds
 */
function deleteArchiveFileRowsByCatalogIds_(catalogIds) {
  var ids = {};
  (catalogIds || []).forEach(function (id) {
    if (id) {
      ids[String(id)] = true;
    }
  });
  var idList = Object.keys(ids);
  if (!idList.length) {
    return;
  }
  rewriteSheetRemovingRows_('Files', 'catalog_id', ids);
  removeAclForTrashObjects_(idList, []);
}

/**
 * §24.4 п.5 — выдать на Drive role=reader всем, у кого в каталоге был любой доступ
 * (user + участники групп). Владельца каталога / CONTROLLER_EMAIL пропускаем.
 * sendNotificationEmail=false.
 *
 * @param {Array<{ driveId: string, objectType: 'file'|'folder', objectId: string }>} pairs
 * @param {string} token
 */
function grantCatalogAclAsDriveReadersBatch_(pairs, token) {
  if (!pairs || !pairs.length || !token) {
    return;
  }
  var controllerLc = String(
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || ''
  )
    .trim()
    .toLowerCase();
  var engine = createAclEngine_();
  var postReqs = [];
  var seenGrant = {};

  pairs.forEach(function (pair) {
    var driveId = String((pair && pair.driveId) || '').trim();
    var objectType = String((pair && pair.objectType) || '').trim();
    var objectId = String((pair && pair.objectId) || '').trim();
    if (!driveId || !objectId || (objectType !== 'file' && objectType !== 'folder')) {
      return;
    }
    var emails = collectArchiveReaderEmailsFromCatalog_(engine, objectType, objectId, controllerLc);
    emails.forEach(function (email) {
      var key = driveId + '\t' + email;
      if (seenGrant[key]) {
        return;
      }
      seenGrant[key] = true;
      postReqs.push({
        url:
          'https://www.googleapis.com/drive/v3/files/' +
          encodeURIComponent(driveId) +
          '/permissions?supportsAllDrives=true&sendNotificationEmail=false',
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({
          role: 'reader',
          type: 'user',
          emailAddress: email
        }),
        muteHttpExceptions: true
      });
    });
  });

  for (var off = 0; off < postReqs.length; off += 100) {
    UrlFetchApp.fetchAll(postReqs.slice(off, off + 100));
  }
}

/**
 * @param {Object} engine
 * @param {'file'|'folder'} objectType
 * @param {string} objectId
 * @param {string} skipEmailLc
 * @returns {string[]} emails (original casing best-effort)
 */
function collectArchiveReaderEmailsFromCatalog_(engine, objectType, objectId, skipEmailLc) {
  var map = getEffectiveAclMapFromEngine_(engine, objectType, objectId);
  var byLc = {};
  Object.keys(map || {}).forEach(function (pKey) {
    var entry = map[pKey];
    if (!entry || !entry.level || entry.level === 'none') {
      return;
    }
    if (String(entry.principalType || '') === 'group') {
      var members = (engine.groupMembers && engine.groupMembers[entry.principalId]) || [];
      members.forEach(function (em) {
        var raw = String(em || '').trim();
        var lc = raw.toLowerCase();
        if (!raw || lc === skipEmailLc) {
          return;
        }
        byLc[lc] = raw;
      });
      return;
    }
    var userRaw = String(entry.principalId || '').trim();
    var userLc = userRaw.toLowerCase();
    if (!userRaw || userLc === skipEmailLc) {
      return;
    }
    byLc[userLc] = userRaw;
  });
  return Object.keys(byLc).map(function (k) {
    return byLc[k];
  });
}

/**
 * Редакторов/комментаторов → читатели (пакетно через Drive API).
 * Владелец не трогаем. Если writer/commenter нет — почти бесплатно (один list).
 *
 * @param {string[]} fileIds
 * @param {string} token
 */
function demoteDriveFilesCollaboratorsBatch_(fileIds, token) {
  var ids = [];
  var seen = {};
  (fileIds || []).forEach(function (id) {
    id = String(id || '').trim();
    if (!id || seen[id]) {
      return;
    }
    seen[id] = true;
    ids.push(id);
  });
  if (!ids.length || !token) {
    return;
  }

  var listReqs = ids.map(function (fileId) {
    return {
      url:
        'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(fileId) +
        '/permissions?supportsAllDrives=true&fields=' +
        encodeURIComponent('permissions(id,type,role)'),
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };
  });

  var listResps = UrlFetchApp.fetchAll(listReqs);
  var patchReqs = [];

  for (var i = 0; i < ids.length; i++) {
    var resp = listResps[i];
    if (!resp || resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
      continue;
    }
    var body;
    try {
      body = JSON.parse(resp.getContentText() || '{}');
    } catch (eParse) {
      continue;
    }
    var perms = body.permissions || [];
    for (var p = 0; p < perms.length; p++) {
      var perm = perms[p];
      if (!perm || !perm.id) {
        continue;
      }
      var role = String(perm.role || '');
      if (role !== 'writer' && role !== 'commenter') {
        continue;
      }
      patchReqs.push({
        url:
          'https://www.googleapis.com/drive/v3/files/' +
          encodeURIComponent(ids[i]) +
          '/permissions/' +
          encodeURIComponent(perm.id) +
          '?supportsAllDrives=true',
        method: 'patch',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ role: 'reader' }),
        muteHttpExceptions: true
      });
    }
  }

  // fetchAll лимит ~100 за вызов
  for (var off = 0; off < patchReqs.length; off += 100) {
    UrlFetchApp.fetchAll(patchReqs.slice(off, off + 100));
  }
}

/**
 * @deprecated оставлен для совместимости; архивация использует batch.
 * @param {GoogleAppsScript.Drive.File|GoogleAppsScript.Drive.Folder} fileOrFolder
 */
function demoteDriveCollaboratorsToReaders_(fileOrFolder) {
  if (!fileOrFolder) {
    return;
  }
  try {
    demoteDriveFilesCollaboratorsBatch_(
      [fileOrFolder.getId()],
      ScriptApp.getOAuthToken()
    );
  } catch (e) {
    // ignore
  }
}
