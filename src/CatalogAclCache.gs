/**
 * §4 — явный ACL на каждом объекте + кэш колонок acl_* в Tree/Files.
 * Наследование только при миграции «сирота → явные строки» и при копировании с родителя.
 */

/**
 * Миграция при старте/snapshot: сиротам — явный ACL с родителя; пустой кэш — заполнить.
 *
 * @param {Object} engine
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @returns {{ aclCopied: number, cacheFilled: number }}
 */
function ensureExplicitAclAndCache_(engine, treeRows, fileRows) {
  engine = engine || createAclEngine_();
  treeRows = treeRows || readSheetRecords_('Tree');
  fileRows = fileRows || readSheetRecords_('Files');

  var aclCopied = materializeOrphanExplicitAcls_(engine);
  var cacheFilled = fillEmptyAclCacheColumns_(engine, treeRows, fileRows);
  return { aclCopied: aclCopied, cacheFilled: cacheFilled };
}

/**
 * Объекты без своих ACL-строк, но с наследованием от предка → копируем явно.
 *
 * @param {Object} engine
 * @returns {number} сколько объектов получили ACL
 */
function materializeOrphanExplicitAcls_(engine) {
  var newRows = [];
  var objectCount = 0;

  function copyInherited(objectType, objectId) {
    var key = objectType + ':' + objectId;
    if ((engine.aclByObject[key] || []).length) {
      return;
    }
    var inherited = resolveInheritedAclRows_(engine, objectType, objectId);
    if (!inherited.length) {
      return;
    }
    objectCount++;
    if (!engine.aclByObject[key]) {
      engine.aclByObject[key] = [];
    }
    inherited.forEach(function (row) {
      var principalType = String(row.principal_type || '').trim();
      var principalId = String(row.principal_id || '').trim();
      if (!principalType || !principalId) {
        return;
      }
      var level = normalizePermissionLevel_(row.permission_level);
      var newRow = {
        acl_id: Utilities.getUuid(),
        object_type: objectType,
        object_id: objectId,
        principal_type: principalType,
        principal_id: principalId,
        permission_level: level
      };
      engine.aclByObject[key].push(newRow);
      newRows.push([
        newRow.acl_id,
        objectType,
        objectId,
        principalType,
        principalId,
        level
      ]);
    });
  }

  Object.keys(engine.foldersById || {}).forEach(function (folderId) {
    copyInherited('folder', folderId);
  });
  Object.keys(engine.filesByCatalogId || {}).forEach(function (catalogId) {
    copyInherited('file', catalogId);
  });

  if (!newRows.length) {
    return 0;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ACL');
  }
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, newRows.length, 6).setValues(newRows);
  return objectCount;
}

/**
 * @param {Object.<string, string>[]} aclRows
 * @returns {Array<{ principalType: string, principalId: string, permissionLevel: string }>}
 */
function aclRowsToEntries_(aclRows) {
  var entries = [];
  (aclRows || []).forEach(function (row) {
    var principalType = String(row.principal_type || '').trim();
    var principalId = String(row.principal_id || '').trim();
    if (!principalType || !principalId) {
      return;
    }
    entries.push({
      principalType: principalType,
      principalId: principalId,
      permissionLevel: normalizePermissionLevel_(row.permission_level)
    });
  });
  return entries;
}

/**
 * Заполняет пустые acl_* из явного ACL (группа = один ярлык #Имя).
 *
 * @param {Object} engine
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @returns {number}
 */
function fillEmptyAclCacheColumns_(engine, treeRows, fileRows) {
  var treeUpdates = [];
  (treeRows || []).forEach(function (row) {
    if (!aclCacheFieldsEmpty_(row)) {
      return;
    }
    var folderId = String(row.folder_id || '');
    if (!folderId) {
      return;
    }
    var labels = aclRowsToCacheLabels_(
      engine,
      engine.aclByObject['folder:' + folderId] || [],
      false
    );
    if (!labels.editors.length && !labels.commenters.length && !labels.readers.length) {
      return;
    }
    treeUpdates.push({
      folderId: folderId,
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    });
    row.acl_editors = treeUpdates[treeUpdates.length - 1].aclEditors;
    row.acl_commenters = treeUpdates[treeUpdates.length - 1].aclCommenters;
    row.acl_readers = treeUpdates[treeUpdates.length - 1].aclReaders;
  });

  var fileUpdates = [];
  (fileRows || []).forEach(function (row) {
    if (!aclCacheFieldsEmpty_(row)) {
      return;
    }
    var catalogId = String(row.catalog_id || '');
    if (!catalogId) {
      return;
    }
    var approved = parseBoolean_(row.approved);
    var labels = aclRowsToCacheLabels_(
      engine,
      engine.aclByObject['file:' + catalogId] || [],
      approved
    );
    if (!labels.editors.length && !labels.commenters.length && !labels.readers.length) {
      return;
    }
    fileUpdates.push({
      catalogId: catalogId,
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    });
    row.acl_editors = fileUpdates[fileUpdates.length - 1].aclEditors;
    row.acl_commenters = fileUpdates[fileUpdates.length - 1].aclCommenters;
    row.acl_readers = fileUpdates[fileUpdates.length - 1].aclReaders;
  });

  writeTreeAclCacheBatch_(treeUpdates);
  writeFilesAclCacheBatch_(fileUpdates);
  return treeUpdates.length + fileUpdates.length;
}

/**
 * @param {Object} row
 * @returns {boolean}
 */
function aclCacheFieldsEmpty_(row) {
  return (
    !String(row.acl_editors || '').trim() &&
    !String(row.acl_commenters || '').trim() &&
    !String(row.acl_readers || '').trim()
  );
}

/**
 * Явный ACL → ярлыки для колонок (группа = #Имя, без раскрытия членов).
 *
 * @param {Object} engine
 * @param {Object.<string, string>[]} aclRows
 * @param {boolean} approved
 * @returns {{ editors: string[], commenters: string[], readers: string[] }}
 */
function aclRowsToCacheLabels_(engine, aclRows, approved) {
  var editors = [];
  var commenters = [];
  var readers = [];
  var seen = {};

  (aclRows || []).forEach(function (row) {
    var level = normalizePermissionLevel_(row.permission_level);
    if (!level || level === 'none') {
      return;
    }
    if (approved && level === 'editor') {
      level = 'commenter';
    }

    var principalType = String(row.principal_type || '').trim();
    var principalId = String(row.principal_id || '').trim();
    if (!principalId) {
      return;
    }

    var label = '';
    var dedupe = '';
    if (principalType === 'group') {
      label = formatGroupAclLabel_(engine, principalId);
      dedupe = 'g:' + principalId;
    } else {
      label = resolveUserLabelFromEngine_(engine, principalId);
      dedupe = 'u:' + principalId.toLowerCase();
    }
    if (!label || seen[dedupe]) {
      return;
    }
    seen[dedupe] = true;

    if (level === 'editor') {
      editors.push(label);
    } else if (level === 'commenter') {
      commenters.push(label);
    } else if (level === 'reader') {
      readers.push(label);
    }
  });

  return {
    editors: sortEmails_(editors),
    commenters: sortEmails_(commenters),
    readers: sortEmails_(readers)
  };
}

/**
 * @param {Object} engine
 * @param {string} groupId
 * @returns {string}
 */
function formatGroupAclLabel_(engine, groupId) {
  var name =
    (engine && engine.groupNameById && engine.groupNameById[groupId]) ||
    String(groupId || '').trim();
  if (!name) {
    return '';
  }
  return name.charAt(0) === '#' ? name : '#' + name;
}

/**
 * @param {string[]} labels
 * @returns {string}
 */
function formatAclCacheField_(labels) {
  return (labels || []).join(', ');
}

/**
 * @param {string} value
 * @returns {string[]}
 */
function parseAclCacheField_(value) {
  var raw = String(value || '').trim();
  if (!raw) {
    return [];
  }
  return raw
    .split(',')
    .map(function (part) {
      return String(part || '').trim();
    })
    .filter(Boolean);
}

/**
 * Копирует явный ACL родителя на объект + пишет кэш.
 *
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {string} parentFolderId
 */
function copyExplicitAclFromParentFolder_(engine, objectType, objectId, parentFolderId) {
  var entries = buildAclEntriesFromObject_(engine, 'folder', parentFolderId);
  replaceAclForObjects_(
    [{ objectType: objectType, objectId: objectId }],
    entries,
    engine
  );
  if (engine && engine.aclByObject) {
    engine.aclByObject[objectType + ':' + objectId] = (entries || []).map(function (e) {
      return {
        acl_id: '',
        object_type: objectType,
        object_id: objectId,
        principal_type: e.principalType,
        principal_id: e.principalId,
        permission_level: e.permissionLevel
      };
    });
  }
}

/**
 * Синхронизирует acl_* для списка объектов из entries (после replaceAcl).
 *
 * @param {Array<{ objectType: 'folder'|'file', objectId: string }>} targetObjects
 * @param {Array} entries
 * @param {Object} engine
 */
function syncAclCacheForObjects_(targetObjects, entries, engine) {
  if (!targetObjects || !targetObjects.length) {
    return;
  }
  var treeUpdates = [];
  var fileUpdates = [];

  targetObjects.forEach(function (obj) {
    var approved = false;
    if (obj.objectType === 'file') {
      var file = engine.filesByCatalogId[obj.objectId];
      approved = file && parseBoolean_(file.approved);
    }
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
    var payload = {
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    };
    if (obj.objectType === 'folder') {
      treeUpdates.push({
        folderId: obj.objectId,
        aclEditors: payload.aclEditors,
        aclCommenters: payload.aclCommenters,
        aclReaders: payload.aclReaders
      });
    } else {
      fileUpdates.push({
        catalogId: obj.objectId,
        aclEditors: payload.aclEditors,
        aclCommenters: payload.aclCommenters,
        aclReaders: payload.aclReaders
      });
    }
  });

  writeTreeAclCacheBatch_(treeUpdates);
  writeFilesAclCacheBatch_(fileUpdates);
}

/**
 * Пересобрать кэш для всех объектов, где в ACL есть группа (после rename).
 *
 * @param {string} groupId
 */
function rebuildAclCacheForGroupPrincipal_(groupId) {
  groupId = String(groupId || '').trim();
  if (!groupId) {
    return;
  }
  var engine = createAclEngine_();
  var targets = [];
  Object.keys(engine.aclByObject || {}).forEach(function (key) {
    var rows = engine.aclByObject[key] || [];
    var hit = false;
    for (var i = 0; i < rows.length; i++) {
      if (
        String(rows[i].principal_type || '') === 'group' &&
        String(rows[i].principal_id || '') === groupId
      ) {
        hit = true;
        break;
      }
    }
    if (!hit) {
      return;
    }
    var parts = key.split(':');
    targets.push({ objectType: parts[0], objectId: parts.slice(1).join(':') });
  });

  targets.forEach(function (obj) {
    var entries = buildAclEntriesFromObject_(engine, obj.objectType, obj.objectId);
    syncAclCacheForObjects_([obj], entries, engine);
  });
}

/**
 * @param {Array<{ folderId: string, aclEditors: string, aclCommenters: string, aclReaders: string }>} updates
 */
function writeTreeAclCacheBatch_(updates) {
  if (!updates || !updates.length) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
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
  var idCol = headers.indexOf('folder_id');
  var eCol = headers.indexOf('acl_editors');
  var cCol = headers.indexOf('acl_commenters');
  var rCol = headers.indexOf('acl_readers');
  if (idCol < 0 || eCol < 0 || cCol < 0 || rCol < 0) {
    return;
  }
  var byId = {};
  updates.forEach(function (u) {
    byId[String(u.folderId)] = u;
  });
  var dirty = false;
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][idCol] || '');
    if (!byId[id]) {
      continue;
    }
    values[i][eCol] = byId[id].aclEditors;
    values[i][cCol] = byId[id].aclCommenters;
    values[i][rCol] = byId[id].aclReaders;
    dirty = true;
  }
  if (dirty) {
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  }
}

/**
 * @param {Array<{ catalogId: string, aclEditors: string, aclCommenters: string, aclReaders: string }>} updates
 */
function writeFilesAclCacheBatch_(updates) {
  if (!updates || !updates.length) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
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
  var idCol = headers.indexOf('catalog_id');
  var eCol = headers.indexOf('acl_editors');
  var cCol = headers.indexOf('acl_commenters');
  var rCol = headers.indexOf('acl_readers');
  if (idCol < 0 || eCol < 0 || cCol < 0 || rCol < 0) {
    return;
  }
  var byId = {};
  updates.forEach(function (u) {
    byId[String(u.catalogId)] = u;
  });
  var dirty = false;
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][idCol] || '');
    if (!byId[id]) {
      continue;
    }
    values[i][eCol] = byId[id].aclEditors;
    values[i][cCol] = byId[id].aclCommenters;
    values[i][rCol] = byId[id].aclReaders;
    dirty = true;
  }
  if (dirty) {
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  }
}
