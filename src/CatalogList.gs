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
  backfillMissingMimeTypes_(fileRows);
  var folderSizes = buildFolderSizeIndex_(treeRows, fileRows);
  var aclRows = readSheetRecords_('ACL');
  var engine = null;
  var emptyAcl = { editors: [], commenters: [], readers: [] };

  if (aclRows.length) {
    engine = buildAclEngineFromRows_(
      treeRows,
      fileRows,
      aclRows,
      readSheetRecords_('GroupMembers'),
      readSheetRecords_('Users')
    );
  }

  function aclFor(objectType, objectId) {
    if (!engine) {
      return emptyAcl;
    }
    return getEffectiveAclDisplayFromEngine_(engine, objectType, objectId);
  }

  var folders = treeRows.map(function (row) {
    var acl = aclFor('folder', row.folder_id);
    return {
      id: row.folder_id,
      parentFolderId: row.parent_folder_id ? String(row.parent_folder_id) : null,
      name: row.name,
      sizeBytes: folderSizes[row.folder_id] || 0,
      modifiedAt: formatCatalogDate_(row.folder_created_at),
      isSystem: parseBoolean_(row.is_system),
      editors: acl.editors,
      commenters: acl.commenters,
      readers: acl.readers
    };
  });

  var files = fileRows.map(function (row) {
    var acl = aclFor('file', row.catalog_id);
    return {
      id: row.catalog_id,
      folderId: String(row.folder_id || ''),
      name: row.display_name,
      mimeType: row.mime_type || '',
      sizeBytes: parseNumber_(row.size_bytes),
      modifiedAt: formatCatalogDate_(row.drive_modified_at),
      approved: parseBoolean_(row.approved),
      approvedBy: row.approved_by || '',
      editors: acl.editors,
      commenters: acl.commenters,
      readers: acl.readers
    };
  });

  return {
    virtualRootFolderId: getVirtualRootFolderId_(),
    folders: folders,
    files: files
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
  var folderSizes = buildFolderSizeIndex_(treeRows, fileRows);
  var aclRows = readSheetRecords_('ACL');
  var engine = null;
  var emptyAcl = { editors: [], commenters: [], readers: [] };

  if (aclRows.length) {
    engine = buildAclEngineFromRows_(
      treeRows,
      fileRows,
      aclRows,
      readSheetRecords_('GroupMembers'),
      readSheetRecords_('Users')
    );
  }

  function aclFor(objectType, objectId) {
    if (!engine) {
      return emptyAcl;
    }
    return getEffectiveAclDisplayFromEngine_(engine, objectType, objectId);
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
      var acl = aclFor('folder', row.folder_id);
      items.push({
        kind: 'folder',
        id: row.folder_id,
        name: row.name,
        sizeBytes: folderSizes[row.folder_id] || 0,
        modifiedAt: formatCatalogDate_(row.folder_created_at),
        approved: false,
        approvedBy: '',
        isSystem: parseBoolean_(row.is_system),
        editors: acl.editors,
        commenters: acl.commenters,
        readers: acl.readers
      });
    });

    fileRows.forEach(function (row) {
      if (row.folder_id !== resolvedFolderId) {
        return;
      }
      var fileAcl = aclFor('file', row.catalog_id);
      items.push({
        kind: 'file',
        id: row.catalog_id,
        name: row.display_name,
        mimeType: row.mime_type || '',
        sizeBytes: parseNumber_(row.size_bytes),
        modifiedAt: formatCatalogDate_(row.drive_modified_at),
        approved: parseBoolean_(row.approved),
        approvedBy: row.approved_by || '',
        isSystem: false,
        editors: fileAcl.editors,
        commenters: fileAcl.commenters,
        readers: fileAcl.readers
      });
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
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @returns {Object.<string, number>}
 */
function buildFolderSizeIndex_(treeRows, fileRows) {
  var childrenByParent = {};
  treeRows.forEach(function (row) {
    var parentId = String(row.parent_folder_id || '');
    if (!childrenByParent[parentId]) {
      childrenByParent[parentId] = [];
    }
    childrenByParent[parentId].push(row.folder_id);
  });

  var directFileSize = {};
  fileRows.forEach(function (row) {
    var folderId = String(row.folder_id || '');
    var size = parseNumber_(row.size_bytes) || 0;
    directFileSize[folderId] = (directFileSize[folderId] || 0) + size;
  });

  var memo = {};
  function sizeOf(folderId) {
    if (memo[folderId] !== undefined) {
      return memo[folderId];
    }
    var total = directFileSize[folderId] || 0;
    var children = childrenByParent[folderId] || [];
    for (var i = 0; i < children.length; i++) {
      total += sizeOf(children[i]);
    }
    memo[folderId] = total;
    return total;
  }

  treeRows.forEach(function (row) {
    sizeOf(row.folder_id);
  });
  return memo;
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
