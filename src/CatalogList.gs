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
  assertCatalogReadyLight_();
  ensureCatalogSchemaUpToDate_();

  var treeRows = readSheetRecords_('Tree');
  var fileRows = readSheetRecords_('Files');
  var aclRows = readSheetRecords_('ACL');
  var groupMemberRows = readSheetRecords_('GroupMembers');
  var userRows = readSheetRecords_('Users');
  var groupRows = readSheetRecords_('Groups');

  var engine = buildAclEngineFromRows_(
    treeRows,
    fileRows,
    aclRows,
    groupMemberRows,
    userRows,
    groupRows
  );
  ensureExplicitAclAndCache_(engine, treeRows, fileRows);

  backfillMissingMimeTypes_(fileRows);
  var folderStats = buildFolderStatsIndex_(treeRows, fileRows);
  var folderSizes = folderStats.sizes;
  var folderFileCounts = folderStats.fileCounts;

  function approvedByNameFor_(email) {
    return resolveUserLabelFromEngine_(engine, email);
  }

  var folders = treeRows.map(function (row) {
    var mirrorOf = String(row.mirror_of_folder_id || '').trim();
    var mirrorOfDrive = String(row.mirror_of_drive_folder_id || '').trim();
    var isExternalMirror = !!mirrorOfDrive;
    var isMirror = !!mirrorOf || isExternalMirror;
    var displayId = mirrorOf ? mirrorOf : row.folder_id;
    var aclRow = mirrorOf
      ? treeRows.filter(function (r) {
          return String(r.folder_id) === mirrorOf;
        })[0] || row
      : row;
    var editors = isExternalMirror ? [] : parseAclCacheField_(aclRow.acl_editors);
    var commenters = isExternalMirror ? [] : parseAclCacheField_(aclRow.acl_commenters);
    var readers = isExternalMirror ? [] : parseAclCacheField_(aclRow.acl_readers);
    if (mirrorOf && !isExternalMirror) {
      var targetDisp = getEffectiveAclDisplayFromEngine_(engine, 'folder', mirrorOf);
      editors = targetDisp.editors || editors;
      commenters = targetDisp.commenters || commenters;
      readers = targetDisp.readers || readers;
    }
    return {
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
      isExternalMirror: isExternalMirror,
      mirrorOfFolderId: mirrorOf,
      mirrorOfDriveFolderId: mirrorOfDrive,
      editors: editors,
      commenters: commenters,
      readers: readers
    };
  });

  var files = fileRows.map(function (row) {
    return buildCatalogFileListItem_(row, engine, approvedByNameFor_);
  });

  return {
    virtualRootFolderId: getVirtualRootFolderId_(),
    folders: folders,
    files: files,
    catalogRev: getCatalogRev_()
  };
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
  var shortcutOf = String(row.shortcut_of_catalog_id || '').trim();
  var shortcutOfDrive = String(row.shortcut_of_drive_file_id || '').trim();
  var isExternalShortcut = !!shortcutOfDrive;
  var isShortcut = !!shortcutOf || isExternalShortcut;

  var display = row;
  var editors = parseAclCacheField_(row.acl_editors);
  var commenters = parseAclCacheField_(row.acl_commenters);
  var readers = parseAclCacheField_(row.acl_readers);
  var approvedBy = row.approved_by || '';
  var approved = parseBoolean_(row.approved);

  if (isExternalShortcut) {
    editors = [];
    commenters = [];
    readers = [];
    approved = false;
    approvedBy = '';
  } else if (shortcutOf && engine && engine.filesByCatalogId) {
    var targetId = shortcutOf;
    try {
      targetId = resolveFileShortcutTargetCatalogId_(engine.filesByCatalogId, String(row.catalog_id));
    } catch (eResolve) {
      targetId = shortcutOf;
    }
    var target = engine.filesByCatalogId[targetId];
    if (target) {
      display = target;
      approvedBy = target.approved_by || '';
      approved = parseBoolean_(target.approved);
      var targetAcl = getEffectiveAclDisplayFromEngine_(engine, 'file', targetId);
      editors = targetAcl.editors || [];
      commenters = targetAcl.commenters || [];
      readers = targetAcl.readers || [];
    }
  }

  return {
    id: row.catalog_id,
    folderId: String(row.folder_id || ''),
    name: row.display_name,
    mimeType: display.mime_type || row.mime_type || '',
    sizeBytes: parseNumber_(display.size_bytes != null ? display.size_bytes : row.size_bytes),
    modifiedAt: formatCatalogDate_(display.drive_modified_at || row.drive_modified_at),
    approved: approved,
    approvedBy: approvedBy,
    approvedByName: approvedBy ? approvedByNameFor_(approvedBy) : '',
    status: String(display.status || row.status || 'ready').toLowerCase() || 'ready',
    isShortcut: isShortcut,
    isExternalShortcut: isExternalShortcut,
    shortcutOfCatalogId: shortcutOf,
    shortcutOfDriveFileId: shortcutOfDrive,
    openUrl: isExternalShortcut
      ? buildCatalogFileOpenUrl_(shortcutOfDrive, row.mime_type || '', '')
      : '',
    editors: editors,
    commenters: commenters,
    readers: readers
  };
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
      record[headers[c]] = values[i][c];
    }
    rows.push(record);
  }
  return rows;
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
  function statsOf(folderId) {
    if (memoBytes[folderId] !== undefined) {
      return;
    }
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
