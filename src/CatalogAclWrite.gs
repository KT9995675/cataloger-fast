/**
 * §4.4, §4.5, §3.7.2 — назначение прав на объект.
 * Для папки перезаписывает ACL на всех вложенных папках и файлах (§4.4).
 *
 * @param {{
 *   objectType: 'folder'|'file',
 *   objectId: string,
 *   entries: Array<{
 *     principalType: 'user'|'group',
 *     principalId: string,
 *     permissionLevel: 'none'|'reader'|'commenter'|'editor'
 *   }>
 * }} input
 * @returns {{ ok: true, updatedObjectCount: number }}
 */
function setObjectAcl(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var objectType = String(input.objectType || '').trim();
  var objectId = String(input.objectId || '').trim();
  var entries = input.entries || [];

  if (objectType !== 'folder' && objectType !== 'file') {
    throw catalogError_('INVALID_INPUT', 'objectType must be folder or file.');
  }
  if (!objectId) {
    throw catalogError_('INVALID_INPUT', 'objectId is required.');
  }
  if (!Array.isArray(entries)) {
    throw catalogError_('INVALID_INPUT', 'entries must be an array.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  assertCatalogObjectExists_(engine, objectType, objectId);
  assertCanAssignAcl_(engine, userEmail, loginRole, objectType, objectId);

  var normalizedEntries = validateAndNormalizeAclEntries_(entries, engine);
  var updatedCount = applyObjectAclEntries_(
    engine,
    objectType,
    objectId,
    normalizedEntries
  );
  bumpCatalogRev_();

  return {
    ok: true,
    updatedObjectCount: updatedCount
  };
}

/**
 * §2.2.4a — те же entries на каждый объект из списка (один bump).
 *
 * @param {{
 *   objects: Array<{ objectType: 'folder'|'file', objectId: string }>,
 *   entries: Array
 * }} input
 * @returns {{ ok: true, updatedObjectCount: number, objectCount: number }}
 */
function setObjectsAcl(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var objects = input.objects || [];
  var entries = input.entries || [];

  if (!Array.isArray(objects) || !objects.length) {
    throw catalogError_('INVALID_INPUT', 'objects must be a non-empty array.');
  }
  if (!Array.isArray(entries)) {
    throw catalogError_('INVALID_INPUT', 'entries must be an array.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  var normalizedEntries = validateAndNormalizeAclEntries_(entries, engine);
  var seen = {};
  var targets = [];

  for (var i = 0; i < objects.length; i++) {
    var obj = objects[i] || {};
    var objectType = String(obj.objectType || '').trim();
    var objectId = String(obj.objectId || '').trim();
    if (objectType !== 'folder' && objectType !== 'file') {
      throw catalogError_('INVALID_INPUT', 'objectType must be folder or file.');
    }
    if (!objectId) {
      throw catalogError_('INVALID_INPUT', 'objectId is required.');
    }
    var key = objectType + ':' + objectId;
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    assertCatalogObjectExists_(engine, objectType, objectId);
    assertCanAssignAcl_(engine, userEmail, loginRole, objectType, objectId);
    targets.push({ objectType: objectType, objectId: objectId });
  }

  targets.sort(function (a, b) {
    if (a.objectType === b.objectType) {
      return 0;
    }
    return a.objectType === 'folder' ? -1 : 1;
  });

  var updatedCount = 0;
  targets.forEach(function (t) {
    updatedCount += applyObjectAclEntries_(
      engine,
      t.objectType,
      t.objectId,
      normalizedEntries
    );
  });
  bumpCatalogRev_();

  return {
    ok: true,
    updatedObjectCount: updatedCount,
    objectCount: targets.length
  };
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {Array} normalizedEntries
 * @returns {number}
 */
function applyObjectAclEntries_(engine, objectType, objectId, normalizedEntries) {
  var parent = getAclParentObject_(engine, objectType, objectId);
  var updatedCount = 1;

  if (objectType === 'folder') {
    var subtree = collectFolderSubtreeObjects_(engine, objectId);
    var descendants = [];
    subtree.forEach(function (obj) {
      if (obj.objectType === 'folder' && obj.objectId === objectId) {
        return;
      }
      descendants.push(obj);
    });

    var folderRows;
    if (!parent) {
      folderRows = normalizedEntries.map(function (e) {
        return {
          principalType: e.principalType,
          principalId: e.principalId,
          permissionLevel: e.permissionLevel,
          delta: 'base'
        };
      });
    } else {
      var parentMapForFolder = getEffectiveAclMapFromEngine_(
        engine,
        parent.objectType,
        parent.objectId
      );
      folderRows = diffAclMapsToDeltas_(parentMapForFolder, normalizedEntries);
    }

    replaceAclForObjects_(
      [{ objectType: 'folder', objectId: objectId }],
      folderRows,
      engine,
      { writeDeltas: true, skipCacheSync: true }
    );
    if (descendants.length) {
      clearAclRowsForObjects_(descendants, engine);
    }
    syncAclCacheForObjects_(subtree, normalizedEntries, engine);
    updatedCount = subtree.length;
  } else if (!parent) {
    replaceAclForObjects_(
      [{ objectType: 'file', objectId: objectId }],
      normalizedEntries.map(function (e) {
        return {
          principalType: e.principalType,
          principalId: e.principalId,
          permissionLevel: e.permissionLevel,
          delta: 'base'
        };
      }),
      engine,
      { writeDeltas: true, skipCacheSync: true }
    );
    syncAclCacheForObjects_(
      [{ objectType: 'file', objectId: objectId }],
      normalizedEntries,
      engine
    );
  } else {
    var parentMap = getEffectiveAclMapFromEngine_(
      engine,
      parent.objectType,
      parent.objectId
    );
    var deltas = diffAclMapsToDeltas_(parentMap, normalizedEntries);
    replaceAclForObjects_(
      [{ objectType: 'file', objectId: objectId }],
      deltas,
      engine,
      { writeDeltas: true, skipCacheSync: true }
    );
    syncAclCacheForObjects_(
      [{ objectType: 'file', objectId: objectId }],
      normalizedEntries,
      engine
    );
  }

  return updatedCount;
}

/**
 * §2.2.4, §4.5 — явные записи ACL на объекте и справочники для UI.
 *
 * @param {{
 *   objectType: 'folder'|'file',
 *   objectId: string
 * }} input
 * @returns {{
 *   ok: true,
 *   objectType: 'folder'|'file',
 *   objectId: string,
 *   objectName: string,
 *   approved: boolean,
 *   canEdit: boolean,
 *   propagateToSubtree: boolean,
 *   principals: Array<{
 *     principalType: 'user'|'group',
 *     principalId: string,
 *     label: string,
 *     permissionLevel: 'none'|'reader'|'commenter'|'editor'
 *   }>
 * }}
 */
function getObjectAclForEdit(input) {
  assertCatalogReady_();

  input = input || {};
  var objectType = String(input.objectType || '').trim();
  var objectId = String(input.objectId || '').trim();

  if (objectType !== 'folder' && objectType !== 'file') {
    throw catalogError_('INVALID_INPUT', 'objectType must be folder or file.');
  }
  if (!objectId) {
    throw catalogError_('INVALID_INPUT', 'objectId is required.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  assertCatalogObjectExists_(engine, objectType, objectId);

  var approved = false;
  if (objectType === 'file') {
    var fileRow = engine.filesByCatalogId[objectId];
    approved = fileRow && parseBoolean_(fileRow.approved);
  }

  var groups = [];
  readSheetRecords_('Groups').forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    if (groupId) {
      groups.push({
        groupId: groupId,
        name: String(row.name || '').trim() || groupId
      });
    }
  });
  groups.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru');
  });

  var users = [];
  readSheetRecords_('Users').forEach(function (row) {
    var email = String(row.email || '').trim();
    if (email) {
      users.push({
        email: email,
        displayName: resolveUserDisplayName_(row)
      });
    }
  });
  users.sort(function (a, b) {
    return String(a.displayName || a.email).localeCompare(
      String(b.displayName || b.email),
      'ru'
    );
  });

  var principals = [];
  groups.forEach(function (group) {
    principals.push({
      principalType: 'group',
      principalId: group.groupId,
      label: '#' + group.name,
      permissionLevel: getEffectivePermissionForGroupFromEngine_(
        engine,
        objectType,
        objectId,
        group.groupId
      )
    });
  });
  users.forEach(function (user) {
    principals.push({
      principalType: 'user',
      principalId: user.email,
      label: user.displayName || user.email,
      permissionLevel: getEffectivePermissionForUserFromEngine_(
        engine,
        objectType,
        objectId,
        user.email
      )
    });
  });

  return {
    ok: true,
    objectType: objectType,
    objectId: objectId,
    objectName: resolveCatalogObjectName_(engine, objectType, objectId),
    approved: approved,
    canEdit: canAssignAcl_(engine, userEmail, loginRole, objectType, objectId),
    propagateToSubtree: objectType === 'folder',
    principals: principals
  };
}

/**
 * §2.2.4a — ACL UI для нескольких объектов: radio = max эффективных уровней.
 *
 * @param {{
 *   objects: Array<{ objectType: 'folder'|'file', objectId: string }>
 * }} input
 * @returns {{
 *   ok: true,
 *   objects: Array<{ objectType: string, objectId: string, objectName: string }>,
 *   objectCount: number,
 *   objectName: string,
 *   approved: boolean,
 *   canEdit: boolean,
 *   propagateToSubtree: boolean,
 *   principals: Array
 * }}
 */
function getObjectsAclForEdit(input) {
  assertCatalogReady_();

  input = input || {};
  var rawObjects = input.objects || [];
  if (!Array.isArray(rawObjects) || !rawObjects.length) {
    throw catalogError_('INVALID_INPUT', 'objects must be a non-empty array.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  var seen = {};
  var targets = [];

  for (var i = 0; i < rawObjects.length; i++) {
    var obj = rawObjects[i] || {};
    var objectType = String(obj.objectType || '').trim();
    var objectId = String(obj.objectId || '').trim();
    if (objectType !== 'folder' && objectType !== 'file') {
      throw catalogError_('INVALID_INPUT', 'objectType must be folder or file.');
    }
    if (!objectId) {
      throw catalogError_('INVALID_INPUT', 'objectId is required.');
    }
    var key = objectType + ':' + objectId;
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    assertCatalogObjectExists_(engine, objectType, objectId);
    targets.push({
      objectType: objectType,
      objectId: objectId,
      objectName: resolveCatalogObjectName_(engine, objectType, objectId)
    });
  }

  if (targets.length === 1) {
    var single = getObjectAclForEdit({
      objectType: targets[0].objectType,
      objectId: targets[0].objectId
    });
    single.objects = targets;
    single.objectCount = 1;
    return single;
  }

  var groups = [];
  readSheetRecords_('Groups').forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    if (groupId) {
      groups.push({
        groupId: groupId,
        name: String(row.name || '').trim() || groupId
      });
    }
  });
  groups.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru');
  });

  var users = [];
  readSheetRecords_('Users').forEach(function (row) {
    var email = String(row.email || '').trim();
    if (email) {
      users.push({
        email: email,
        displayName: resolveUserDisplayName_(row)
      });
    }
  });
  users.sort(function (a, b) {
    return String(a.displayName || a.email).localeCompare(
      String(b.displayName || b.email),
      'ru'
    );
  });

  var anyApproved = false;
  var anyFolder = false;
  var canEdit = true;
  targets.forEach(function (t) {
    if (t.objectType === 'folder') {
      anyFolder = true;
    } else {
      var fileRow = engine.filesByCatalogId[t.objectId];
      if (fileRow && parseBoolean_(fileRow.approved)) {
        anyApproved = true;
      }
    }
    if (!canAssignAcl_(engine, userEmail, loginRole, t.objectType, t.objectId)) {
      canEdit = false;
    }
  });

  var principals = [];
  groups.forEach(function (group) {
    var level = 'none';
    targets.forEach(function (t) {
      level = maxPermissionLevel_(
        level,
        getEffectivePermissionForGroupFromEngine_(
          engine,
          t.objectType,
          t.objectId,
          group.groupId
        )
      );
    });
    principals.push({
      principalType: 'group',
      principalId: group.groupId,
      label: '#' + group.name,
      permissionLevel: level
    });
  });
  users.forEach(function (user) {
    var level = 'none';
    targets.forEach(function (t) {
      level = maxPermissionLevel_(
        level,
        getEffectivePermissionForUserFromEngine_(
          engine,
          t.objectType,
          t.objectId,
          user.email
        )
      );
    });
    principals.push({
      principalType: 'user',
      principalId: user.email,
      label: user.displayName || user.email,
      permissionLevel: level
    });
  });

  var folders = 0;
  var files = 0;
  targets.forEach(function (t) {
    if (t.objectType === 'folder') {
      folders += 1;
    } else {
      files += 1;
    }
  });
  var nameParts = [];
  if (folders) {
    nameParts.push(folders + ' пап.');
  }
  if (files) {
    nameParts.push(files + ' файл.');
  }

  return {
    ok: true,
    objects: targets,
    objectCount: targets.length,
    objectType: 'batch',
    objectId: '',
    objectName: nameParts.join(', ') + ' (выбрано: ' + targets.length + ')',
    approved: anyApproved,
    canEdit: canEdit,
    propagateToSubtree: anyFolder,
    principals: principals
  };
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {string}
 */
function resolveCatalogObjectName_(engine, objectType, objectId) {
  if (objectType === 'folder') {
    var folder = engine.foldersById[objectId];
    return folder ? String(folder.name || objectId) : objectId;
  }
  var file = engine.filesByCatalogId[objectId];
  return file ? String(file.display_name || objectId) : objectId;
}

/**
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {Object} engine
 * @returns {Array}
 */
function readExplicitAclEntries_(objectType, objectId, engine) {
  return readSheetRecords_('ACL')
    .filter(function (row) {
      return (
        String(row.object_type || '').trim() === objectType &&
        String(row.object_id || '').trim() === objectId
      );
    })
    .map(function (row) {
      var principalType = String(row.principal_type || '').trim().toLowerCase();
      var principalId = String(row.principal_id || '').trim();
      return {
        principalType: principalType === 'group' ? 'group' : 'user',
        principalId: principalId,
        principalLabel: resolvePrincipalLabel_(principalType, principalId, engine),
        permissionLevel: normalizePermissionLevel_(row.permission_level)
      };
    });
}

/**
 * @param {string} principalType
 * @param {string} principalId
 * @param {Object} engine
 * @returns {string}
 */
function resolvePrincipalLabel_(principalType, principalId, engine) {
  if (String(principalType || '').trim().toLowerCase() === 'group') {
    var groups = readSheetRecords_('Groups');
    for (var i = 0; i < groups.length; i++) {
      if (String(groups[i].group_id || '').trim() === principalId) {
        return String(groups[i].name || '').trim() || principalId;
      }
    }
    return principalId;
  }
  return principalId;
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {boolean}
 */
function canAssignAcl_(engine, userEmail, loginRole, objectType, objectId) {
  try {
    assertCanAssignAcl_(engine, userEmail, loginRole, objectType, objectId);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * @param {string} userEmail
 * @returns {'user'|'manager'|'controller'}
 */
function getLoginRoleForUser_(userEmail) {
  var users = readSheetRecords_('Users');
  var normalized = userEmail.toLowerCase();
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].email || '').toLowerCase() === normalized) {
      var role = String(users[i].login_role || 'user').trim().toLowerCase();
      if (role === 'controller' || role === 'manager') {
        return role;
      }
      return 'user';
    }
  }
  throw catalogError_('NOT_IN_CATALOG', 'Ваш аккаунт не добавлен в каталог.');
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 */
function assertCatalogObjectExists_(engine, objectType, objectId) {
  if (objectType === 'folder') {
    if (!engine.foldersById[objectId]) {
      throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + objectId);
    }
    return;
  }
  if (!engine.filesByCatalogId[objectId]) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + objectId);
  }
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 */
function assertCanAssignAcl_(engine, userEmail, loginRole, objectType, objectId) {
  if (loginRole === 'controller') {
    return;
  }
  if (loginRole !== 'manager') {
    throw catalogError_('NOT_ALLOWED', 'Назначать права могут только Менеджер и Управляющий.');
  }

  var permission = getEffectivePermissionForUserFromEngine_(
    engine,
    objectType,
    objectId,
    userEmail
  );
  if (permission !== 'editor') {
    throw catalogError_('NOT_ALLOWED', 'Нужно право редактор на объект.');
  }
}

/**
 * @param {Array} entries
 * @param {Object} engine
 * @returns {Array<{
 *   principalType: 'user'|'group',
 *   principalId: string,
 *   permissionLevel: 'none'|'reader'|'commenter'|'editor'
 * }>}
 */
function validateAndNormalizeAclEntries_(entries, engine) {
  var users = {};
  readSheetRecords_('Users').forEach(function (row) {
    var email = String(row.email || '').trim();
    if (email) {
      users[email.toLowerCase()] = email;
    }
  });

  var groups = {};
  readSheetRecords_('Groups').forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    if (groupId) {
      groups[groupId] = true;
    }
  });

  var seen = {};
  var normalized = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i] || {};
    var principalType = String(entry.principalType || '').trim().toLowerCase();
    var principalId = String(entry.principalId || '').trim();
    var permissionLevel = normalizePermissionLevel_(entry.permissionLevel);

    if (principalType !== 'user' && principalType !== 'group') {
      throw catalogError_('INVALID_INPUT', 'principalType must be user or group.');
    }
    if (!principalId) {
      throw catalogError_('INVALID_INPUT', 'principalId is required.');
    }

    if (principalType === 'user') {
      var canonicalEmail = users[principalId.toLowerCase()];
      if (!canonicalEmail) {
        throw catalogError_('UNKNOWN_USER', 'User not found in catalog: ' + principalId);
      }
      principalId = canonicalEmail;
    } else if (!groups[principalId]) {
      throw catalogError_('UNKNOWN_GROUP', 'Group not found: ' + principalId);
    }

    var dedupeKey = principalType + ':' + principalId.toLowerCase();
    if (seen[dedupeKey]) {
      throw catalogError_('INVALID_INPUT', 'Duplicate principal in entries: ' + principalId);
    }
    seen[dedupeKey] = true;

    normalized.push({
      principalType: principalType,
      principalId: principalId,
      permissionLevel: permissionLevel
    });
  }

  return normalized;
}

/**
 * @param {Object} engine
 * @param {string} folderId
 * @returns {Array<{ objectType: 'folder'|'file', objectId: string }>}
 */
function collectFolderSubtreeObjects_(engine, folderId) {
  var folderIds = {};
  var queue = [folderId];
  folderIds[folderId] = true;

  while (queue.length) {
    var currentId = queue.shift();
    Object.keys(engine.foldersById).forEach(function (childId) {
      var child = engine.foldersById[childId];
      if (String(child.parent_folder_id || '') === currentId && !folderIds[childId]) {
        folderIds[childId] = true;
        queue.push(childId);
      }
    });
  }

  var objects = Object.keys(folderIds).map(function (id) {
    return { objectType: 'folder', objectId: id };
  });

  Object.keys(engine.filesByCatalogId).forEach(function (catalogId) {
    var file = engine.filesByCatalogId[catalogId];
    var parentFolderId = String(file.folder_id || '');
    if (folderIds[parentFolderId]) {
      objects.push({ objectType: 'file', objectId: catalogId });
    }
  });

  return objects;
}

/**
 * @param {Array<{ objectType: 'folder'|'file', objectId: string }>} targetObjects
 * @param {Array} entries — permissionLevel; опц. delta ('+'|'-'|'base')
 * @param {Object} engine
 * @param {{ writeDeltas?: boolean, skipCacheSync?: boolean }=} options
 */
function replaceAclForObjects_(targetObjects, entries, engine, options) {
  options = options || {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ACL');
  }

  ensureCatalogSchemaUpToDate_();

  var values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw catalogError_('SCHEMA_MISMATCH', 'ACL sheet has no header row.');
  }

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var objectTypeCol = headers.indexOf('object_type');
  var objectIdCol = headers.indexOf('object_id');
  var deltaCol = headers.indexOf('delta');
  if (objectTypeCol < 0 || objectIdCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'ACL sheet headers are invalid.');
  }

  var targetKeys = {};
  targetObjects.forEach(function (obj) {
    targetKeys[obj.objectType + ':' + obj.objectId] = true;
  });

  var kept = [headers];
  for (var i = 1; i < values.length; i++) {
    var rowKey =
      String(values[i][objectTypeCol] || '').trim() +
      ':' +
      String(values[i][objectIdCol] || '').trim();
    if (!targetKeys[rowKey]) {
      kept.push(values[i]);
    }
  }

  targetObjects.forEach(function (obj) {
    var approved = false;
    if (obj.objectType === 'file') {
      var file = engine.filesByCatalogId[obj.objectId];
      approved = file && parseBoolean_(file.approved);
    }

    for (var e = 0; e < entries.length; e++) {
      var entry = entries[e];
      var level = entry.permissionLevel;
      if (approved && level === 'editor' && entry.delta !== '-') {
        level = 'commenter';
      }

      var line = [];
      for (var c = 0; c < headers.length; c++) {
        var h = headers[c];
        if (h === 'acl_id') {
          line.push(Utilities.getUuid());
        } else if (h === 'object_type') {
          line.push(obj.objectType);
        } else if (h === 'object_id') {
          line.push(obj.objectId);
        } else if (h === 'principal_type') {
          line.push(entry.principalType);
        } else if (h === 'principal_id') {
          line.push(entry.principalId);
        } else if (h === 'permission_level') {
          line.push(level);
        } else if (h === 'delta') {
          var d = entry.delta != null ? normalizeAclDelta_(entry.delta) : '';
          if (!d && options.writeDeltas) {
            d = '+';
          }
          if (d === 'base') {
            d = '';
          }
          line.push(d === '+' || d === '-' ? d : '');
        } else {
          line.push('');
        }
      }
      kept.push(line);
    }

    if (engine && engine.aclByObject) {
      engine.aclByObject[obj.objectType + ':' + obj.objectId] = (entries || []).map(
        function (en) {
          var d = en.delta != null ? normalizeAclDelta_(en.delta) : '';
          if (d === 'base') {
            d = '';
          }
          return {
            acl_id: '',
            object_type: obj.objectType,
            object_id: obj.objectId,
            principal_type: en.principalType,
            principal_id: en.principalId,
            permission_level: en.permissionLevel,
            delta: d
          };
        }
      );
    }
  });

  sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);

  var lastRow = sheet.getLastRow();
  if (lastRow > kept.length) {
    sheet
      .getRange(kept.length + 1, 1, lastRow - kept.length, headers.length)
      .clearContent();
  }

  if (!options.skipCacheSync) {
    syncAclCacheForObjects_(targetObjects, entries, engine);
  }
}

/**
 * Удалить все строки ACL для объектов (сброс отклонений).
 *
 * @param {Array<{ objectType: string, objectId: string }>} targetObjects
 * @param {Object} engine
 */
function clearAclRowsForObjects_(targetObjects, engine) {
  if (!targetObjects || !targetObjects.length) {
    return;
  }
  replaceAclForObjects_(targetObjects, [], engine, { skipCacheSync: true });
}

/**
 * Добавить ACL для новых объектов (без удаления старых строк) + один проход кэша.
 * У каждого объекта свой список entries.
 *
 * @param {Array<{
 *   objectType: 'folder'|'file',
 *   objectId: string,
 *   entries: Array<{ principalType: string, principalId: string, permissionLevel: string }>
 * }>} objectAclList
 * @param {Object} engine
 */
function appendAclForNewObjectsBatch_(objectAclList, engine) {
  if (!objectAclList || !objectAclList.length) {
    return;
  }

  ensureCatalogSchemaUpToDate_();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ACL');
  }

  var values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw catalogError_('SCHEMA_MISMATCH', 'ACL sheet has no header row.');
  }

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var kept = values.slice();
  var treeUpdates = [];
  var fileUpdates = [];

  objectAclList.forEach(function (item) {
    var approved = false;
    if (item.objectType === 'file') {
      var file = engine.filesByCatalogId[item.objectId];
      approved = file && parseBoolean_(file.approved);
    }

    var entries = item.entries || [];
    for (var e = 0; e < entries.length; e++) {
      var entry = entries[e];
      var level = entry.permissionLevel;
      if (approved && level === 'editor' && entry.delta !== '-') {
        level = 'commenter';
      }
      var line = [];
      for (var c = 0; c < headers.length; c++) {
        var h = headers[c];
        if (h === 'acl_id') {
          line.push(Utilities.getUuid());
        } else if (h === 'object_type') {
          line.push(item.objectType);
        } else if (h === 'object_id') {
          line.push(item.objectId);
        } else if (h === 'principal_type') {
          line.push(entry.principalType);
        } else if (h === 'principal_id') {
          line.push(entry.principalId);
        } else if (h === 'permission_level') {
          line.push(level);
        } else if (h === 'delta') {
          var d = entry.delta != null ? normalizeAclDelta_(entry.delta) : '';
          if (d === 'base') {
            d = '';
          }
          line.push(d === '+' || d === '-' ? d : '');
        } else {
          line.push('');
        }
      }
      kept.push(line);
    }

    var labels = aclRowsToCacheLabels_(
      engine,
      entries.map(function (en) {
        return {
          principal_type: en.principalType,
          principal_id: en.principalId,
          permission_level: en.permissionLevel
        };
      }),
      approved
    );
    var payload = {
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    };
    if (item.objectType === 'folder') {
      treeUpdates.push({
        folderId: item.objectId,
        aclEditors: payload.aclEditors,
        aclCommenters: payload.aclCommenters,
        aclReaders: payload.aclReaders
      });
    } else {
      fileUpdates.push({
        catalogId: item.objectId,
        aclEditors: payload.aclEditors,
        aclCommenters: payload.aclCommenters,
        aclReaders: payload.aclReaders
      });
    }
  });

  if (kept.length > values.length) {
    sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
    var lastRow = sheet.getLastRow();
    if (lastRow > kept.length) {
      sheet
        .getRange(kept.length + 1, 1, lastRow - kept.length, headers.length)
        .clearContent();
    }
  }

  writeTreeAclCacheBatch_(treeUpdates);
  writeFilesAclCacheBatch_(fileUpdates);
}
