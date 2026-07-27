/**
 * §2.2.4, §4, §11 — эффективные права объекта для столбцов UI.
 *
 * @param {'folder'|'file'} objectType
 * @param {string} objectId `folder_id` или `catalog_id`
 * @returns {{
 *   editors: string[],
 *   commenters: string[],
 *   readers: string[]
 * }}
 */
function getEffectiveAclDisplay(objectType, objectId) {
  assertCatalogReady_();

  objectType = String(objectType || '').trim();
  objectId = String(objectId || '').trim();
  if (objectType !== 'folder' && objectType !== 'file') {
    throw catalogError_('INVALID_INPUT', 'objectType must be folder or file.');
  }
  if (!objectId) {
    throw catalogError_('INVALID_INPUT', 'objectId is required.');
  }

  var engine = createAclEngine_();
  return getEffectiveAclDisplayFromEngine_(engine, objectType, objectId);
}

/**
 * @returns {Object}
 */
function createAclEngine_() {
  return buildAclEngineFromRows_(
    readSheetRecords_('Tree'),
    readSheetRecords_('Files'),
    readSheetRecords_('ACL'),
    readSheetRecords_('GroupMembers'),
    readSheetRecords_('Users'),
    readSheetRecords_('Groups')
  );
}

/**
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @param {Object.<string, string>[]} aclRows
 * @param {Object.<string, string>[]} groupMemberRows
 * @param {Object.<string, string>[]=} userRows
 * @param {Object.<string, string>[]=} groupRows
 * @returns {Object}
 */
function buildAclEngineFromRows_(treeRows, fileRows, aclRows, groupMemberRows, userRows, groupRows) {
  var foldersById = {};
  treeRows.forEach(function (row) {
    foldersById[row.folder_id] = row;
  });

  var filesByCatalogId = {};
  fileRows.forEach(function (row) {
    filesByCatalogId[row.catalog_id] = row;
  });

  var aclByObject = {};
  aclRows.forEach(function (row) {
    var objectType = String(row.object_type || '').trim();
    var objectId = String(row.object_id || '').trim();
    if (!objectType || !objectId) {
      return;
    }
    var key = objectType + ':' + objectId;
    if (!aclByObject[key]) {
      aclByObject[key] = [];
    }
    aclByObject[key].push(row);
  });

  var groupMembers = {};
  groupMemberRows.forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    var email = String(row.email || '').trim();
    if (!groupId || !email) {
      return;
    }
    if (!groupMembers[groupId]) {
      groupMembers[groupId] = [];
    }
    groupMembers[groupId].push(email);
  });

  var userDisplayNameByEmail = {};
  (userRows || []).forEach(function (row) {
    var email = String(row.email || '').trim();
    if (!email) {
      return;
    }
    userDisplayNameByEmail[email.toLowerCase()] = resolveUserDisplayName_(row);
  });

  var groupNameById = {};
  (groupRows || []).forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    if (!groupId) {
      return;
    }
    groupNameById[groupId] = String(row.name || '').trim() || groupId;
  });

  return {
    foldersById: foldersById,
    filesByCatalogId: filesByCatalogId,
    aclByObject: aclByObject,
    groupMembers: groupMembers,
    userDisplayNameByEmail: userDisplayNameByEmail,
    groupNameById: groupNameById
  };
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {{
 *   editors: string[],
 *   commenters: string[],
 *   readers: string[]
 * }}
 */
function getEffectiveAclDisplayFromEngine_(engine, objectType, objectId) {
  var approved = false;
  if (objectType === 'file') {
    var file = engine.filesByCatalogId[objectId];
    if (!file) {
      throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + objectId);
    }
    approved = parseBoolean_(file.approved);
  } else if (!engine.foldersById[objectId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + objectId);
  }

  var rows = effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, objectType, objectId));
  return aclRowsToCacheLabels_(engine, rows, approved);
}

/**
 * §4.4a — эффективная карта principal → уровень (мать ⊕ дельты; корень = база).
 *
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {Object=} memo
 * @returns {Object.<string, { principalType: string, principalId: string, level: string }>}
 */
function getEffectiveAclMapFromEngine_(engine, objectType, objectId, memo) {
  memo = memo || {};
  var key = objectType + ':' + objectId;
  if (memo[key]) {
    return cloneAclPrincipalMap_(memo[key]);
  }

  var parent = getAclParentObject_(engine, objectType, objectId);
  var rows = engine.aclByObject[key] || [];
  var map = {};

  if (!parent) {
    rows.forEach(function (row) {
      var delta = normalizeAclDelta_(row);
      if (delta === '-') {
        return;
      }
      var level = normalizePermissionLevel_(row.permission_level);
      if (!level || level === 'none') {
        return;
      }
      var pKey = aclPrincipalMapKey_(row.principal_type, row.principal_id);
      if (!pKey) {
        return;
      }
      map[pKey] = {
        principalType: String(row.principal_type || '').trim() === 'group' ? 'group' : 'user',
        principalId: String(row.principal_id || '').trim(),
        level: level
      };
    });
  } else {
    map = getEffectiveAclMapFromEngine_(
      engine,
      parent.objectType,
      parent.objectId,
      memo
    );
    map = cloneAclPrincipalMap_(map);

    var minusKeys = {};
    rows.forEach(function (row) {
      if (normalizeAclDelta_(row) !== '-') {
        return;
      }
      var pKey = aclPrincipalMapKey_(row.principal_type, row.principal_id);
      if (pKey) {
        minusKeys[pKey] = true;
        delete map[pKey];
      }
    });

    rows.forEach(function (row) {
      var delta = normalizeAclDelta_(row);
      if (delta === '-') {
        return;
      }
      var pKey = aclPrincipalMapKey_(row.principal_type, row.principal_id);
      if (!pKey || minusKeys[pKey]) {
        return;
      }
      var level = normalizePermissionLevel_(row.permission_level);
      if (!level || level === 'none') {
        return;
      }
      map[pKey] = {
        principalType: String(row.principal_type || '').trim() === 'group' ? 'group' : 'user',
        principalId: String(row.principal_id || '').trim(),
        level: level
      };
    });
  }

  memo[key] = map;
  return cloneAclPrincipalMap_(map);
}

/**
 * @param {Object|string} rowOrDelta
 * @returns {'+'|'-'|'base'|''}
 */
function normalizeAclDelta_(rowOrDelta) {
  var raw =
    typeof rowOrDelta === 'string'
      ? rowOrDelta
      : rowOrDelta && (rowOrDelta.delta != null ? rowOrDelta.delta : rowOrDelta.Delta);
  var d = String(raw == null ? '' : raw).trim();
  if (d === '-' || d === '−' || d === 'minus') {
    return '-';
  }
  if (d === '+' || d === 'plus') {
    return '+';
  }
  if (d === 'base') {
    return 'base';
  }
  return '';
}

/**
 * @param {string} principalType
 * @param {string} principalId
 * @returns {string}
 */
function aclPrincipalMapKey_(principalType, principalId) {
  var type = String(principalType || '').trim().toLowerCase() === 'group' ? 'group' : 'user';
  var id = String(principalId || '').trim();
  if (!id) {
    return '';
  }
  if (type === 'user') {
    id = id.toLowerCase();
  }
  return type + ':' + id;
}

/**
 * @param {Object} map
 * @returns {Object}
 */
function cloneAclPrincipalMap_(map) {
  var out = {};
  Object.keys(map || {}).forEach(function (k) {
    var v = map[k];
    out[k] = {
      principalType: v.principalType,
      principalId: v.principalId,
      level: v.level
    };
  });
  return out;
}

/**
 * @param {Object} map
 * @returns {Array<{ principal_type: string, principal_id: string, permission_level: string }>}
 */
function effectiveAclMapToRows_(map) {
  var rows = [];
  Object.keys(map || {}).forEach(function (k) {
    var v = map[k];
    rows.push({
      principal_type: v.principalType,
      principal_id: v.principalId,
      permission_level: v.level
    });
  });
  return rows;
}

/**
 * @param {Object} map
 * @returns {Array<{ principalType: string, principalId: string, permissionLevel: string }>}
 */
function effectiveAclMapToEntries_(map) {
  return effectiveAclMapToRows_(map).map(function (row) {
    return {
      principalType: row.principal_type,
      principalId: row.principal_id,
      permissionLevel: row.permission_level
    };
  });
}

/**
 * Желаемые эффективные entries → дельты относительно карты матери.
 *
 * @param {Object} parentMap
 * @param {Array<{ principalType: string, principalId: string, permissionLevel: string }>} desiredEntries
 * @returns {Array<{ principalType: string, principalId: string, permissionLevel: string, delta: '+'|'-' }>}
 */
function diffAclMapsToDeltas_(parentMap, desiredEntries) {
  var desiredMap = {};
  (desiredEntries || []).forEach(function (entry) {
    var level = normalizePermissionLevel_(entry.permissionLevel);
    if (!level || level === 'none') {
      return;
    }
    var pKey = aclPrincipalMapKey_(entry.principalType, entry.principalId);
    if (!pKey) {
      return;
    }
    desiredMap[pKey] = {
      principalType: entry.principalType === 'group' ? 'group' : 'user',
      principalId: String(entry.principalId || '').trim(),
      level: level
    };
  });

  var deltas = [];
  var seen = {};
  Object.keys(parentMap || {}).forEach(function (k) {
    seen[k] = true;
  });
  Object.keys(desiredMap).forEach(function (k) {
    seen[k] = true;
  });

  Object.keys(seen).forEach(function (k) {
    var parent = parentMap[k];
    var desired = desiredMap[k];
    if (!parent && desired) {
      deltas.push({
        principalType: desired.principalType,
        principalId: desired.principalId,
        permissionLevel: desired.level,
        delta: '+'
      });
      return;
    }
    if (parent && !desired) {
      deltas.push({
        principalType: parent.principalType,
        principalId: parent.principalId,
        permissionLevel: parent.level,
        delta: '-'
      });
      return;
    }
    if (parent && desired && parent.level !== desired.level) {
      deltas.push({
        principalType: parent.principalType,
        principalId: parent.principalId,
        permissionLevel: parent.level,
        delta: '-'
      });
      deltas.push({
        principalType: desired.principalType,
        principalId: desired.principalId,
        permissionLevel: desired.level,
        delta: '+'
      });
    }
  });

  return deltas;
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} sourceType
 * @param {string} sourceId
 * @returns {Object.<string, string>[]}
 */
function resolveInheritedAclRows_(engine, sourceType, sourceId) {
  return effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, sourceType, sourceId));
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {({ objectType: 'folder', objectId: string }|null)}
 */
function getAclParentObject_(engine, objectType, objectId) {
  if (objectType === 'file') {
    var file = engine.filesByCatalogId[objectId];
    if (!file || !file.folder_id) {
      return null;
    }
    return { objectType: 'folder', objectId: String(file.folder_id) };
  }

  var folder = engine.foldersById[objectId];
  if (!folder || !folder.parent_folder_id) {
    return null;
  }
  return { objectType: 'folder', objectId: String(folder.parent_folder_id) };
}

/**
 * @param {Object} engine
 * @param {Object.<string, string>[]} aclRows
 * @param {boolean} approved
 * @returns {{
 *   editors: string[],
 *   commenters: string[],
 *   readers: string[]
 * }}
 */
function aclRowsToDisplay_(engine, aclRows, approved) {
  var personal = {};
  var groupLevels = {};
  var emailDisplay = {};

  aclRows.forEach(function (row) {
    var level = normalizePermissionLevel_(row.permission_level);
    var principalType = String(row.principal_type || '').trim();
    var principalId = String(row.principal_id || '').trim();
    if (!principalId) {
      return;
    }

    if (principalType === 'user') {
      rememberEmail_(emailDisplay, principalId);
      personal[principalId.toLowerCase()] = level;
      return;
    }

    if (principalType === 'group') {
      var members = engine.groupMembers[principalId] || [];
      members.forEach(function (email) {
        rememberEmail_(emailDisplay, email);
        var key = email.toLowerCase();
        groupLevels[key] = maxPermissionLevel_(groupLevels[key], level);
      });
    }
  });

  var editors = [];
  var commenters = [];
  var readers = [];
  var seen = {};

  Object.keys(personal).forEach(function (key) {
    seen[key] = true;
  });
  Object.keys(groupLevels).forEach(function (key) {
    seen[key] = true;
  });

  Object.keys(seen).forEach(function (key) {
    var level = personal.hasOwnProperty(key) ? personal[key] : groupLevels[key] || 'none';
    if (level === 'none') {
      return;
    }
    if (approved && level === 'editor') {
      level = 'commenter';
    }

    var email = emailDisplay[key] || key;
    var label = resolveUserLabelFromEngine_(engine, email);
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
 * @param {string} email
 * @returns {string}
 */
function resolveUserLabelFromEngine_(engine, email) {
  var trimmed = String(email || '').trim();
  if (!trimmed) {
    return '';
  }
  var map = (engine && engine.userDisplayNameByEmail) || {};
  var named = map[trimmed.toLowerCase()];
  return named || trimmed;
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {string} userEmail
 * @returns {'none'|'reader'|'commenter'|'editor'}
 */
function getEffectivePermissionForUserFromEngine_(engine, objectType, objectId, userEmail) {
  var approved = false;
  if (objectType === 'file') {
    var file = engine.filesByCatalogId[objectId];
    if (!file) {
      throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + objectId);
    }
    approved = parseBoolean_(file.approved);
  } else if (!engine.foldersById[objectId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + objectId);
  }

  var aclRows = effectiveAclMapToRows_(
    getEffectiveAclMapFromEngine_(engine, objectType, objectId)
  );
  return resolveUserPermissionFromAclRows_(engine, aclRows, userEmail, approved);
}

/**
 * @param {Object} engine
 * @param {Object.<string, string>[]} aclRows
 * @param {string} userEmail
 * @param {boolean} approved
 * @returns {'none'|'reader'|'commenter'|'editor'}
 */
function resolveUserPermissionFromAclRows_(engine, aclRows, userEmail, approved) {
  var normalizedEmail = String(userEmail || '').trim().toLowerCase();
  if (!normalizedEmail) {
    return 'none';
  }

  var personal = null;
  var groupLevel = 'none';

  aclRows.forEach(function (row) {
    var level = normalizePermissionLevel_(row.permission_level);
    var principalType = String(row.principal_type || '').trim();
    var principalId = String(row.principal_id || '').trim();
    if (!principalId) {
      return;
    }

    if (principalType === 'user') {
      if (principalId.toLowerCase() === normalizedEmail) {
        personal = level;
      }
      return;
    }

    if (principalType === 'group') {
      var members = engine.groupMembers[principalId] || [];
      for (var i = 0; i < members.length; i++) {
        if (String(members[i]).toLowerCase() === normalizedEmail) {
          groupLevel = maxPermissionLevel_(groupLevel, level);
          break;
        }
      }
    }
  });

  var level = personal !== null ? personal : groupLevel;
  if (approved && level === 'editor') {
    level = 'commenter';
  }
  return level;
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {string} groupId
 * @returns {'none'|'reader'|'commenter'|'editor'}
 */
function getEffectivePermissionForGroupFromEngine_(engine, objectType, objectId, groupId) {
  var approved = false;
  if (objectType === 'file') {
    var file = engine.filesByCatalogId[objectId];
    if (!file) {
      throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + objectId);
    }
    approved = parseBoolean_(file.approved);
  } else if (!engine.foldersById[objectId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + objectId);
  }

  var aclRows = effectiveAclMapToRows_(
    getEffectiveAclMapFromEngine_(engine, objectType, objectId)
  );
  var normalizedGroupId = String(groupId || '').trim();
  var level = 'none';

  aclRows.forEach(function (row) {
    if (String(row.principal_type || '').trim() !== 'group') {
      return;
    }
    if (String(row.principal_id || '').trim() !== normalizedGroupId) {
      return;
    }
    level = maxPermissionLevel_(level, normalizePermissionLevel_(row.permission_level));
  });

  if (approved && level === 'editor') {
    level = 'commenter';
  }
  return level;
}

/**
 * @param {Object} engine
 * @param {string} catalogId
 * @returns {string[]}
 */
function getNoAccessEditorsFromEngine_(engine, catalogId) {
  var file = engine.filesByCatalogId[catalogId];
  if (!file) {
    return [];
  }

  var editors = {};
  var fileAcl = getEffectiveAclDisplayFromEngine_(engine, 'file', catalogId);
  fileAcl.editors.forEach(function (email) {
    editors[email.toLowerCase()] = email;
  });

  if (file.folder_id) {
    var folderAcl = getEffectiveAclDisplayFromEngine_(engine, 'folder', String(file.folder_id));
    folderAcl.editors.forEach(function (email) {
      editors[email.toLowerCase()] = email;
    });
  }

  return sortEmails_(
    Object.keys(editors).map(function (key) {
      return editors[key];
    })
  );
}

/**
 * @param {Object.<string, string>} emailDisplay
 * @param {string} email
 */
function rememberEmail_(emailDisplay, email) {
  var trimmed = String(email || '').trim();
  if (!trimmed) {
    return;
  }
  var key = trimmed.toLowerCase();
  if (!emailDisplay[key]) {
    emailDisplay[key] = trimmed;
  }
}

/**
 * @param {*} value
 * @returns {'none'|'reader'|'commenter'|'editor'}
 */
function normalizePermissionLevel_(value) {
  var level = String(value == null ? '' : value).trim().toLowerCase();
  if (level === 'editor' || level === 'commenter' || level === 'reader' || level === 'none') {
    return level;
  }
  return 'none';
}

/**
 * @param {('none'|'reader'|'commenter'|'editor'|undefined)} a
 * @param {('none'|'reader'|'commenter'|'editor')} b
 * @returns {'none'|'reader'|'commenter'|'editor'}
 */
function maxPermissionLevel_(a, b) {
  var rank = { none: 0, reader: 1, commenter: 2, editor: 3 };
  if (!a || rank[b] > rank[a]) {
    return b;
  }
  return a;
}

/**
 * @param {string[]} emails
 * @returns {string[]}
 */
function sortEmails_(emails) {
  return emails.slice().sort(function (a, b) {
    return String(a).localeCompare(String(b), 'ru', { sensitivity: 'base' });
  });
}
