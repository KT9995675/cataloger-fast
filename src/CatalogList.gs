/**
 * §2.2 — содержимое виртуальной папки для панели каталога.
 *
 * @param {string=} folderId пусто → корневая виртуальная папка (`CATALOG_VIRTUAL_ROOT_FOLDER_ID`)
 * @returns {{
 *   folderId: string,
 *   parentFolderId: (string|null),
 *   folderName: string,
 *   items: Array<{
 *     kind: 'folder'|'file',
 *     id: string,
 *     name: string,
 *     sizeBytes: (number|null),
 *     modifiedAt: (string|null),
 *     approved: boolean,
 *     approvedBy: string,
 *     approvedByName: string,
 *     isSystem: boolean,
 *     editors: string[],
 *     commenters: string[],
 *     readers: string[]
 *   }>
 * }}
 */
function listFolderContents(folderId) {
  var resolvedFolderId = folderId ? String(folderId).trim() : getVirtualRootFolderId_();
  return listFolderContentsBatch([resolvedFolderId])[0];
}

/**
 * §2.7 — полный снимок каталога для клиента (без Drive).
 * Один round-trip; дальше UI листает локально.
 *
 * @returns {{
 *   virtualRootFolderId: string,
 *   folders: Array<{
 *     id: string,
 *     parentFolderId: (string|null),
 *     name: string,
 *     sizeBytes: number,
 *     modifiedAt: (string|null),
 *     isSystem: boolean,
 *     editors: string[],
 *     commenters: string[],
 *     readers: string[]
 *   }>,
 *   files: Array<{
 *     id: string,
 *     folderId: string,
 *     name: string,
 *     sizeBytes: number,
 *     modifiedAt: (string|null),
 *     approved: boolean,
 *     approvedBy: string,
 *     approvedByName: string,
 *     editors: string[],
 *     commenters: string[],
 *     readers: string[]
 *   }>
 * }}
 */
function getCatalogSnapshot() {
  try {
    return buildCatalogSnapshotPayload_();
  } catch (e) {
    var msg = (e && e.message) || String(e);
    var code = (e && e.name) || 'SNAPSHOT_FAILED';
    if (code === 'Error') {
      code = 'SNAPSHOT_FAILED';
    }
    throw catalogError_(code, 'Снимок каталога: ' + msg);
  }
}

/**
 * Только Tree (+ пустые files) — маленький ответ для старта UI.
 * @returns {Object}
 */
function getCatalogTreeSnapshot() {
  try {
    var built = buildCatalogSnapshotPayload_({ skipFiles: true });
    return {
      virtualRootFolderId: built.virtualRootFolderId,
      folders: built.folders,
      files: [],
      catalogRev: built.catalogRev,
      fileTotal: built.fileTotal != null ? built.fileTotal : 0,
      partial: true,
      build: 'r10'
    };
  } catch (e) {
    var msg = (e && e.message) || String(e);
    throw catalogError_('SNAPSHOT_FAILED', 'Снимок Tree: ' + msg);
  }
}

/**
 * Старт передачи Files через CacheService (короткий ответ).
 * Клиент затем забирает чанки fetchCatalogFilesChunk.
 *
 * @returns {{
 *   ok: true,
 *   transferId: string,
 *   chunks: number,
 *   bytes: number,
 *   fileTotal: number,
 *   catalogRev: number,
 *   build: string
 * }}
 */
function beginCatalogFilesTransfer() {
  assertCatalogReadyLight_();
  var fileRows = readSheetRecords_('Files');
  var files = [];
  for (var i = 0; i < fileRows.length; i++) {
    files.push(buildCatalogFileListItemMinimal_(fileRows[i]));
  }
  var json = JSON.stringify(files);
  var transferId = Utilities.getUuid().replace(/-/g, '');
  var chunkSize = 90000;
  var chunks = Math.ceil(json.length / chunkSize) || 1;
  var cache = CacheService.getDocumentCache();
  for (var c = 0; c < chunks; c++) {
    var part = json.substring(c * chunkSize, (c + 1) * chunkSize);
    cache.put('cf:' + transferId + ':' + c, part, 600);
  }
  cache.put('cf:' + transferId + ':n', String(chunks), 600);
  return {
    ok: true,
    transferId: transferId,
    chunks: chunks,
    bytes: json.length,
    fileTotal: files.length,
    catalogRev: getCatalogRev_(),
    build: 'r10'
  };
}

/**
 * @param {string} transferId
 * @param {number} index
 * @returns {{ ok: boolean, transferId: string, index: number, text: string, build: string }}
 */
function fetchCatalogFilesChunk(transferId, index) {
  transferId = String(transferId || '').trim();
  index = Math.max(0, Number(index) || 0);
  var cache = CacheService.getDocumentCache();
  var text = cache.get('cf:' + transferId + ':' + index);
  return {
    ok: text != null && text !== '',
    transferId: transferId,
    index: index,
    text: text || '',
    build: 'r10'
  };
}

/**
 * Страница файлов (legacy / запасной путь). Компактные поля без ACL.
 *
 * @param {{ offset?: number, limit?: number }|number=} input
 * @param {number=} limitMaybe
 * @returns {{ files: Array, offset: number, nextOffset: number, done: boolean, fileTotal: number }}
 */
function getCatalogFilesSnapshotPage(input, limitMaybe) {
  try {
    var offset = 0;
    var limit = 15;
    if (input && typeof input === 'object') {
      offset = Math.max(0, Number(input.offset) || 0);
      limit = Number(input.limit) || 15;
    } else {
      offset = Math.max(0, Number(input) || 0);
      limit = Number(limitMaybe) || 15;
    }
    limit = Math.min(25, Math.max(5, limit));

    assertCatalogReadyLight_();
    var fileRows = readSheetRecords_('Files');
    var fileTotal = fileRows.length;
    var sliceRows = fileRows.slice(offset, offset + limit);
    var files = sliceRows.map(function (row) {
      return buildCatalogFileListItemMinimal_(row);
    });
    var nextOffset = offset + files.length;
    // Строка надёжнее объекта для google.script.run при пограничном размере.
    return JSON.stringify({
      files: files,
      offset: offset,
      nextOffset: nextOffset,
      done: nextOffset >= fileTotal,
      fileTotal: fileTotal,
      catalogRev: getCatalogRev_(),
      build: 'r10'
    });
  } catch (e) {
    var msg = (e && e.message) || String(e);
    throw catalogError_('SNAPSHOT_FAILED', 'Снимок Files: ' + msg);
  }
}

/**
 * Минимальный файл для порционной загрузки (без ACL-строк).
 * @param {Object} row
 * @returns {Object}
 */
function buildCatalogFileListItemMinimal_(row) {
  var item = {
    id: row.catalog_id,
    folderId: String(row.folder_id || ''),
    name: row.display_name,
    sizeBytes: parseNumber_(row.size_bytes)
  };
  var mime = String(row.mime_type || '').trim();
  if (mime) {
    item.mimeType = mime;
  }
  var modifiedAt = formatCatalogDate_(row.drive_modified_at);
  if (modifiedAt) {
    item.modifiedAt = modifiedAt;
  }
  if (parseBoolean_(row.approved)) {
    item.approved = true;
  }
  var status = String(row.status || 'ready').toLowerCase() || 'ready';
  if (status !== 'ready') {
    item.status = status;
  }
  var shortcutOf = String(row.shortcut_of_catalog_id || '').trim();
  var shortcutOfDrive = String(row.shortcut_of_drive_file_id || '').trim();
  if (shortcutOf || shortcutOfDrive) {
    item.isShortcut = true;
  }
  if (shortcutOfDrive) {
    item.isExternalShortcut = true;
    item.shortcutOfDriveFileId = shortcutOfDrive;
  }
  if (shortcutOf) {
    item.shortcutOfCatalogId = shortcutOf;
  }
  return item;
}

/**
 * @param {{
 *   skipFiles?: boolean,
 *   skipFolders?: boolean,
 *   filesOffset?: number,
 *   filesLimit?: number
 * }=} options
 * @returns {{
 *   virtualRootFolderId: string,
 *   folders: Array,
 *   files: Array,
 *   catalogRev: number,
 *   fileTotal?: number
 * }}
 */
function buildCatalogSnapshotPayload_(options) {
  options = options || {};
  assertCatalogReadyLight_();

  var treeRows = options.skipFolders ? [] : readSheetRecords_('Tree');
  var needFileRows =
    !options.skipFolders || !options.skipFiles || options.filesOffset != null;
  var fileRows = needFileRows ? readSheetRecords_('Files') : [];
  var userRows = readSheetRecords_('Users');
  var displayNameByEmail = {};
  (userRows || []).forEach(function (u) {
    var email = String(u.email || '').trim().toLowerCase();
    if (!email) {
      return;
    }
    var name = String(u.display_name || '').trim();
    displayNameByEmail[email] = name || String(u.email || '').trim();
  });

  var filesByCatalogId = {};
  (fileRows || []).forEach(function (row) {
    filesByCatalogId[String(row.catalog_id || '')] = row;
  });
  var treeById = {};
  (treeRows || []).forEach(function (row) {
    treeById[String(row.folder_id || '')] = row;
  });

  var folderSizes = {};
  var folderFileCounts = {};
  if (!options.skipFolders) {
    var folderStats = buildFolderStatsIndex_(treeRows, fileRows);
    folderSizes = folderStats.sizes;
    folderFileCounts = folderStats.fileCounts;
  }

  function approvedByNameFor_(email) {
    var key = String(email || '')
      .trim()
      .toLowerCase();
    return displayNameByEmail[key] || String(email || '').trim();
  }

  var folders = [];
  if (!options.skipFolders) {
    folders = treeRows.map(function (row) {
      var mirrorOf = String(row.mirror_of_folder_id || '').trim();
      var mirrorOfDrive = String(row.mirror_of_drive_folder_id || '').trim();
      var isExternalMirror = !!mirrorOfDrive;
      var isMirror = !!mirrorOf || isExternalMirror;
      var displayId = mirrorOf ? mirrorOf : row.folder_id;
      var aclRow = mirrorOf && treeById[mirrorOf] ? treeById[mirrorOf] : row;
      var item = {
        id: row.folder_id,
        parentFolderId: row.parent_folder_id ? String(row.parent_folder_id) : null,
        name: row.name,
        sizeBytes: isExternalMirror ? 0 : folderSizes[displayId] || folderSizes[row.folder_id] || 0,
        fileCount: isExternalMirror
          ? 0
          : folderFileCounts[displayId] || folderFileCounts[row.folder_id] || 0,
        modifiedAt: formatCatalogDate_(row.folder_created_at),
        isSystem: parseBoolean_(row.is_system),
        isMirror: isMirror,
        isExternalMirror: isExternalMirror
      };
      if (mirrorOf) {
        item.mirrorOfFolderId = mirrorOf;
      }
      if (mirrorOfDrive) {
        item.mirrorOfDriveFolderId = mirrorOfDrive;
      }
      if (!isExternalMirror) {
        var fe = String(aclRow.acl_editors || '').trim();
        var fc = String(aclRow.acl_commenters || '').trim();
        var fr = String(aclRow.acl_readers || '').trim();
        if (fe) {
          item.editors = fe;
        }
        if (fc) {
          item.commenters = fc;
        }
        if (fr) {
          item.readers = fr;
        }
      }
      return item;
    });
  }

  var files = [];
  var fileTotal = fileRows.length;
  if (!options.skipFiles || options.filesOffset != null) {
    var offset = Math.max(0, Number(options.filesOffset) || 0);
    var limit =
      options.filesLimit != null
        ? Math.min(150, Math.max(1, Number(options.filesLimit) || 80))
        : fileRows.length;
    var sliceRows =
      options.filesOffset != null || options.filesLimit != null
        ? fileRows.slice(offset, offset + limit)
        : fileRows;
    files = sliceRows.map(function (row) {
      return buildCatalogFileListItemFromCache_(row, filesByCatalogId, approvedByNameFor_);
    });
  }

  return {
    virtualRootFolderId: getVirtualRootFolderId_(),
    folders: folders,
    files: files,
    catalogRev: getCatalogRev_(),
    fileTotal: fileTotal
  };
}

/**
 * Мгновенный зонд: пауза Jobs + снять триггеры воркера (без abort по Sheets —
 * abort ждёт lock и сам вешает «Загрузка»). Abort — в resumeCatalogJobsAfterUiLoad.
 * @returns {{ ok: true, build: string, catalogRev: number, jobsPaused: true, removedTriggers: number }}
 */
function getCatalogLoadProbe() {
  assertCatalogReadyLight_();
  setCatalogJobsPaused_(true);
  var removedTriggers = 0;
  try {
    removedTriggers = removeCatalogJobsTriggers_();
  } catch (eTrig) {
    removedTriggers = -1;
  }
  try {
    clearCatalogOpStatus_();
  } catch (eSt) {
    // ignore
  }
  return {
    ok: true,
    build: 'r11',
    catalogRev: getCatalogRev_(),
    jobsPaused: true,
    treeRows: -1,
    fileRows: -1,
    aclRows: -1,
    jobsRows: -1,
    abortedJobs: 0,
    removedTriggers: removedTriggers
  };
}

/**
 * Снять паузу Jobs после загрузки UI. Активные pending/running — abort (лёгкий).
 */
function resumeCatalogJobsAfterUiLoad() {
  var aborted = 0;
  try {
    aborted = abortActiveJobStatusesLight_();
  } catch (eAbort) {
    aborted = -1;
  }
  setCatalogJobsPaused_(false);
  // Триггер только если после abort ещё что-то живо (не должно).
  if (hasActiveCatalogJobs_()) {
    ensureCatalogJobsTrigger_();
    kickCatalogJobsProcessing_();
  }
  return { ok: true, busy: hasActiveCatalogJobs_(), abortedJobs: aborted };
}

/**
 * Явно снять паузу Jobs (импорт / ручной kick), без abort очереди.
 * @returns {{ ok: true, wasPaused: boolean }}
 */
function unpauseCatalogJobs() {
  var was = isCatalogJobsPaused_();
  setCatalogJobsPaused_(false);
  if (hasActiveCatalogJobs_()) {
    ensureCatalogJobsTrigger_();
  }
  return { ok: true, wasPaused: was };
}

/**
 * @param {string[]} folderIds
 * @returns {Array}
 */
function listFolderContentsBatch(folderIds) {
  assertCatalogReadyLight_();

  folderIds = folderIds || [];
  if (!folderIds.length) {
    return [];
  }

  var treeRows = readSheetRecords_('Tree');
  var fileRows = readSheetRecords_('Files');
  var folderById = indexTreeFolders_(treeRows);
  var folderStats = buildFolderStatsIndex_(treeRows, fileRows);
  var folderSizes = folderStats.sizes;
  var folderFileCounts = folderStats.fileCounts;
  var engine = buildAclEngineFromRows_(
    treeRows,
    fileRows,
    readSheetRecords_('ACL'),
    readSheetRecords_('GroupMembers'),
    readSheetRecords_('Users'),
    readSheetRecords_('Groups')
  );

  function approvedByNameFor_(email) {
    return resolveUserLabelFromEngine_(engine, email);
  }

  return folderIds.map(function (folderId) {
    var resolvedFolderId = String(folderId || '').trim() || getVirtualRootFolderId_();
    var current = folderById[resolvedFolderId];
    if (!current) {
      throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + resolvedFolderId);
    }

    var items = [];

    treeRows.forEach(function (row) {
      if (row.parent_folder_id !== resolvedFolderId) {
        return;
      }
      var mirrorOf = String(row.mirror_of_folder_id || '').trim();
      var mirrorOfDrive = String(row.mirror_of_drive_folder_id || '').trim();
      var isExternalMirror = !!mirrorOfDrive;
      var isMirror = !!mirrorOf || isExternalMirror;
      var sizeId = mirrorOf ? mirrorOf : row.folder_id;
      var editors = isExternalMirror ? [] : parseAclCacheField_(row.acl_editors);
      var commenters = isExternalMirror ? [] : parseAclCacheField_(row.acl_commenters);
      var readers = isExternalMirror ? [] : parseAclCacheField_(row.acl_readers);
      if (mirrorOf && !isExternalMirror) {
        var targetDisp = getEffectiveAclDisplayFromEngine_(engine, 'folder', mirrorOf);
        editors = targetDisp.editors || editors;
        commenters = targetDisp.commenters || commenters;
        readers = targetDisp.readers || readers;
      }
      items.push({
        kind: 'folder',
        id: row.folder_id,
        name: row.name,
        sizeBytes: isExternalMirror ? 0 : folderSizes[sizeId] || 0,
        fileCount: isExternalMirror ? 0 : folderFileCounts[sizeId] || 0,
        modifiedAt: formatCatalogDate_(row.folder_created_at),
        approved: false,
        approvedBy: '',
        approvedByName: '',
        isSystem: parseBoolean_(row.is_system),
        isMirror: isMirror,
        isExternalMirror: isExternalMirror,
        mirrorOfFolderId: mirrorOf,
        mirrorOfDriveFolderId: mirrorOfDrive,
        editors: editors,
        commenters: commenters,
        readers: readers
      });
    });

    fileRows.forEach(function (row) {
      if (row.folder_id !== resolvedFolderId) {
        return;
      }
      var fileItem = buildCatalogFileListItem_(row, engine, approvedByNameFor_);
      fileItem.kind = 'file';
      fileItem.isSystem = false;
      items.push(fileItem);
    });

    items.sort(comparePanelItems_);

    return {
      folderId: resolvedFolderId,
      parentFolderId: current.parent_folder_id ? current.parent_folder_id : null,
      folderName: current.name,
      items: items
    };
  });
}

function assertCatalogReadyLight_() {
  var props = PropertiesService.getDocumentProperties();
  if (!props.getProperty(PROP_SCHEMA_VERSION_) || !props.getProperty(PROP_CATALOG_ROOT_FOLDER_ID_)) {
    throw catalogError_('CATALOG_NOT_INITIALIZED', 'Catalog is not initialized.');
  }
}

/**
 * Элемент файла для snapshot / listFolderContentsBatch (§22 ярлыки).
 *
 * @param {Object} row
 * @param {Object} engine
 * @param {function(string): string} approvedByNameFor_
 * @returns {Object}
 */
function buildCatalogFileListItem_(row, engine, approvedByNameFor_) {
  var filesByCatalogId =
    engine && engine.filesByCatalogId ? engine.filesByCatalogId : {};
  return buildCatalogFileListItemFromCache_(row, filesByCatalogId, approvedByNameFor_);
}

/**
 * Файл для UI только из кэша acl_* (без пересчёта effective ACL).
 *
 * @param {Object} row
 * @param {Object.<string, Object>} filesByCatalogId
 * @param {function(string): string} approvedByNameFor_
 * @returns {Object}
 */
function buildCatalogFileListItemFromCache_(row, filesByCatalogId, approvedByNameFor_) {
  var shortcutOf = String(row.shortcut_of_catalog_id || '').trim();
  var shortcutOfDrive = String(row.shortcut_of_drive_file_id || '').trim();
  var isExternalShortcut = !!shortcutOfDrive;
  var isShortcut = !!shortcutOf || isExternalShortcut;

  var display = row;
  var editorsStr = String(row.acl_editors || '').trim();
  var commentersStr = String(row.acl_commenters || '').trim();
  var readersStr = String(row.acl_readers || '').trim();
  var approvedBy = row.approved_by || '';
  var approved = parseBoolean_(row.approved);

  if (isExternalShortcut) {
    editorsStr = '';
    commentersStr = '';
    readersStr = '';
    approved = false;
    approvedBy = '';
  } else if (shortcutOf && filesByCatalogId) {
    var targetId = shortcutOf;
    try {
      targetId = resolveFileShortcutTargetCatalogId_(filesByCatalogId, String(row.catalog_id));
    } catch (eResolve) {
      targetId = shortcutOf;
    }
    var target = filesByCatalogId[targetId];
    if (target) {
      display = target;
      approvedBy = target.approved_by || '';
      approved = parseBoolean_(target.approved);
      editorsStr = String(target.acl_editors || '').trim();
      commentersStr = String(target.acl_commenters || '').trim();
      readersStr = String(target.acl_readers || '').trim();
    }
  }

  var status = String(display.status || row.status || 'ready').toLowerCase() || 'ready';
  var item = {
    id: row.catalog_id,
    folderId: String(row.folder_id || ''),
    name: row.display_name,
    sizeBytes: parseNumber_(display.size_bytes != null ? display.size_bytes : row.size_bytes)
  };
  var mime = display.mime_type || row.mime_type || '';
  if (mime) {
    item.mimeType = mime;
  }
  var modifiedAt = formatCatalogDate_(display.drive_modified_at || row.drive_modified_at);
  if (modifiedAt) {
    item.modifiedAt = modifiedAt;
  }
  if (approved) {
    item.approved = true;
  }
  if (approvedBy) {
    item.approvedBy = approvedBy;
    item.approvedByName = approvedByNameFor_(approvedBy);
  }
  if (status && status !== 'ready') {
    item.status = status;
  }
  if (isShortcut) {
    item.isShortcut = true;
  }
  if (isExternalShortcut) {
    item.isExternalShortcut = true;
  }
  if (shortcutOf) {
    item.shortcutOfCatalogId = shortcutOf;
  }
  if (shortcutOfDrive) {
    item.shortcutOfDriveFileId = shortcutOfDrive;
    item.openUrl = buildCatalogFileOpenUrl_(shortcutOfDrive, row.mime_type || '', '');
  }
  if (editorsStr) {
    item.editors = editorsStr;
  }
  if (commentersStr) {
    item.commenters = commentersStr;
  }
  if (readersStr) {
    item.readers = readersStr;
  }
  return item;
}

/**
 * Дочитывает mime_type с Drive для строк без кэша (один batch UrlFetchApp.fetchAll).
 * Пишет в лист и обновляет fileRows на месте.
 *
 * @param {Object.<string, string>[]} fileRows
 */
function backfillMissingMimeTypes_(fileRows) {
  if (!fileRows || !fileRows.length) {
    return;
  }

  var need = [];
  for (var i = 0; i < fileRows.length; i++) {
    var row = fileRows[i];
    if (String(row.mime_type || '').trim()) {
      continue;
    }
    if (!String(row.file_id || '').trim()) {
      continue;
    }
    need.push(row);
  }
  if (!need.length) {
    return;
  }

  var token = ScriptApp.getOAuthToken();
  var limit = Math.min(need.length, 50);
  var requests = [];
  for (var r = 0; r < limit; r++) {
    requests.push({
      url:
        'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(String(need[r].file_id)) +
        '?fields=mimeType&supportsAllDrives=true',
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
  }

  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    return;
  }

  var updates = [];
  for (var j = 0; j < responses.length; j++) {
    var resp = responses[j];
    if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
      continue;
    }
    try {
      var parsed = JSON.parse(resp.getContentText() || '{}');
      var mime = String(parsed.mimeType || '').trim();
      if (!mime) {
        continue;
      }
      need[j].mime_type = mime;
      updates.push({ catalogId: String(need[j].catalog_id), mimeType: mime });
    } catch (e2) {
      // skip
    }
  }

  if (updates.length) {
    writeMimeTypesToFilesSheet_(updates);
  }
}

/**
 * @param {Array<{ catalogId: string, mimeType: string }>} updates
 */
function writeMimeTypesToFilesSheet_(updates) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet || !updates.length) {
    return;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var catalogCol = headers.indexOf('catalog_id');
  var mimeCol = headers.indexOf('mime_type');
  if (catalogCol < 0 || mimeCol < 0) {
    return;
  }

  var byId = {};
  updates.forEach(function (u) {
    byId[u.catalogId] = u.mimeType;
  });

  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][catalogCol] || '');
    if (byId[id]) {
      sheet.getRange(i + 1, mimeCol + 1).setValue(byId[id]);
    }
  }
}

function assertCatalogReady_() {
  var state = isCatalogInitialized();
  if (!state.initialized) {
    // schema 0.1 → 0.2: soft-upgrade before failing on version mismatch
    try {
      ensureCatalogSchemaUpToDate_();
    } catch (e) {
      // ignore; re-check below
    }
    state = isCatalogInitialized();
  }
  if (!state.initialized) {
    throw catalogError_('CATALOG_NOT_INITIALIZED', 'Catalog is not initialized.');
  }
}

/**
 * @returns {string}
 */
function getVirtualRootFolderId_() {
  var id = PropertiesService.getDocumentProperties().getProperty(
    PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_
  );
  if (!id) {
    throw catalogError_('CATALOG_NOT_CONFIGURED', 'CATALOG_VIRTUAL_ROOT_FOLDER_ID is missing.');
  }
  return id;
}

/**
 * @param {string} sheetName
 * @returns {Object.<string, string>[]}
 */
function readSheetRecords_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ' + sheetName);
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var record = {};
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      record[key] = coerceSheetCellForHeader_(key, values[i][c]);
    }
    rows.push(record);
  }
  return rows;
}

/**
 * Дата в текстовом имени (Sheets из «02.05») → строка. Остальные Date не трогаем.
 *
 * @param {string} header
 * @param {*} value
 * @returns {*}
 */
function coerceSheetCellForHeader_(header, value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) {
    return value;
  }
  var h = String(header || '').toLowerCase();
  if (h === 'display_name' || h === 'name') {
    return formatAccidentalSheetDateAsName_(value);
  }
  return value;
}

/**
 * Восстановление имени вроде «02.05» после автодаты Sheets (иначе в UI — ISO).
 *
 * @param {Date} d
 * @returns {string}
 */
var cachedSpreadsheetTz_ = '';
function getSpreadsheetTimeZoneCached_() {
  if (cachedSpreadsheetTz_) {
    return cachedSpreadsheetTz_;
  }
  try {
    cachedSpreadsheetTz_ =
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || '';
  } catch (eTz) {
    cachedSpreadsheetTz_ = '';
  }
  if (!cachedSpreadsheetTz_) {
    cachedSpreadsheetTz_ = Session.getScriptTimeZone() || 'Etc/GMT';
  }
  return cachedSpreadsheetTz_;
}

function formatAccidentalSheetDateAsName_(d) {
  var tz = getSpreadsheetTimeZoneCached_();
  var hm = Utilities.formatDate(d, tz, 'HH:mm');
  if (hm === '00:00') {
    return Utilities.formatDate(d, tz, 'dd.MM');
  }
  return Utilities.formatDate(d, tz, 'dd.MM.yyyy HH:mm');
}

/**
 * Перед setValues: колонки имён как Plain text (@), иначе «02.05» → Date.
 * Сигнатура: getRange(row, column, numRows, numColumns).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow
 * @param {number} numRows
 * @param {string[]} headers
 * @param {string[]} textHeaders
 */
function prepareSheetPlainTextColumns_(sheet, startRow, numRows, headers, textHeaders) {
  if (!sheet || numRows < 1 || !headers || !textHeaders) {
    return;
  }
  for (var i = 0; i < textHeaders.length; i++) {
    var col = headers.indexOf(textHeaders[i]);
    if (col < 0) {
      continue;
    }
    sheet.getRange(startRow, col + 1, numRows, 1).setNumberFormat('@');
  }
}

/**
 * @param {Object.<string, string>[]} treeRows
 * @returns {Object.<string, Object.<string, string>>}
 */
function indexTreeFolders_(treeRows) {
  var map = {};
  treeRows.forEach(function (row) {
    map[row.folder_id] = row;
  });
  return map;
}

/**
 * Рекурсивные размер и число файлов по каждой папке Tree (из листа Files).
 * Зеркала не разворачиваются: у узла-зеркала в индексе — свои прямые дети (обычно 0);
 * UI берёт stats цели через mirror_of_folder_id.
 *
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @returns {{ sizes: Object.<string, number>, fileCounts: Object.<string, number> }}
 */
function buildFolderStatsIndex_(treeRows, fileRows) {
  var childrenByParent = {};
  treeRows.forEach(function (row) {
    var parentId = String(row.parent_folder_id || '');
    if (!childrenByParent[parentId]) {
      childrenByParent[parentId] = [];
    }
    childrenByParent[parentId].push(row.folder_id);
  });

  var directFileSize = {};
  var directFileCount = {};
  fileRows.forEach(function (row) {
    var folderId = String(row.folder_id || '');
    var size = parseNumber_(row.size_bytes) || 0;
    directFileSize[folderId] = (directFileSize[folderId] || 0) + size;
    directFileCount[folderId] = (directFileCount[folderId] || 0) + 1;
  });

  var memoBytes = {};
  var memoCounts = {};
  var visiting = {};
  function statsOf(folderId) {
    if (memoBytes[folderId] !== undefined) {
      return;
    }
    if (visiting[folderId]) {
      // Цикл в Tree.parent_folder_id — считаем 0, чтобы не зависнуть на snapshot.
      memoBytes[folderId] = 0;
      memoCounts[folderId] = 0;
      return;
    }
    visiting[folderId] = true;
    var totalBytes = directFileSize[folderId] || 0;
    var totalCount = directFileCount[folderId] || 0;
    var children = childrenByParent[folderId] || [];
    for (var i = 0; i < children.length; i++) {
      statsOf(children[i]);
      totalBytes += memoBytes[children[i]] || 0;
      totalCount += memoCounts[children[i]] || 0;
    }
    memoBytes[folderId] = totalBytes;
    memoCounts[folderId] = totalCount;
    delete visiting[folderId];
  }

  treeRows.forEach(function (row) {
    statsOf(row.folder_id);
  });
  return { sizes: memoBytes, fileCounts: memoCounts };
}

/**
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @returns {Object.<string, number>}
 */
function buildFolderSizeIndex_(treeRows, fileRows) {
  return buildFolderStatsIndex_(treeRows, fileRows).sizes;
}

/**
 * @param {*} value
 * @returns {boolean}
 */
function parseBoolean_(value) {
  if (value === true) {
    return true;
  }
  var s = String(value == null ? '' : value).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes';
}

/**
 * @param {*} value
 * @returns {number}
 */
function parseNumber_(value) {
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

/**
 * @param {*} value
 * @returns {(string|null)}
 */
function formatCatalogDate_(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}

/**
 * @param {Object} a
 * @param {Object} b
 * @returns {number}
 */
function comparePanelItems_(a, b) {
  if (a.kind !== b.kind) {
    return a.kind === 'folder' ? -1 : 1;
  }
  return String(a.name).localeCompare(String(b.name), 'ru', { sensitivity: 'base' });
}
