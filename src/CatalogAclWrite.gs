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
  var targetObjects =
    objectType === 'folder'
      ? collectFolderSubtreeObjects_(engine, objectId)
      : [{ objectType: 'file', objectId: objectId }];

  replaceAclForObjects_(targetObjects, normalizedEntries, engine);

  return {
    ok: true,
    updatedObjectCount: targetObjects.length
  };
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
 * @param {Array} entries
 * @param {Object} engine
 */
function replaceAclForObjects_(targetObjects, entries, engine) {
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
  var objectTypeCol = headers.indexOf('object_type');
  var objectIdCol = headers.indexOf('object_id');
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
      if (approved && level === 'editor') {
        level = 'commenter';
      }

      kept.push([
        Utilities.getUuid(),
        obj.objectType,
        obj.objectId,
        entry.principalType,
        entry.principalId,
        level
      ]);
    }
  });

  sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);

  var lastRow = sheet.getLastRow();
  if (lastRow > kept.length) {
    // getRange(row, column, numRows, numColumns)
    sheet
      .getRange(kept.length + 1, 1, lastRow - kept.length, headers.length)
      .clearContent();
  }

  syncAclCacheForObjects_(targetObjects, entries, engine);
}
