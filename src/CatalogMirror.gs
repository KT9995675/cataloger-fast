/**
 * §20 — зеркало папки: навигационная ссылка на другую папку каталога.
 * Новые права не назначаются; удаление — сразу (не в корзину).
 */

/**
 * Создать зеркало папки target в parent.
 *
 * @param {{
 *   parentFolderId: string,
 *   targetFolderId: string
 * }} input
 * @returns {{ ok: true, folder: Object }}
 */
function createCatalogMirror(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var targetFolderId = String(input.targetFolderId || '').trim();

  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'Выберите папку для зеркала.');
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

  var resolvedTargetId = resolveMirrorTargetFolderId_(engine.foldersById, targetFolderId);
  var target = engine.foldersById[resolvedTargetId];
  if (!target) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Target folder not found: ' + targetFolderId);
  }

  var name = String(target.name || '').trim();
  if (!name) {
    throw catalogError_('INVALID_INPUT', 'У целевой папки нет имени.');
  }

  var folderId = Utilities.getUuid();
  var now = new Date();
  appendTreeFolderRowsBatch_([
    {
      folderId: folderId,
      parentFolderId: parentFolderId,
      name: name,
      folderCreatedAt: now,
      isSystem: false,
      aclEditors: '',
      aclCommenters: '',
      aclReaders: '',
      mirrorOfFolderId: resolvedTargetId,
      mirrorOfDriveFolderId: ''
    }
  ]);

  var targetAcl = getEffectiveAclDisplayFromEngine_(engine, 'folder', resolvedTargetId);
  var folderStats = buildFolderStatsIndex_(
    readSheetRecords_('Tree'),
    readSheetRecords_('Files')
  );
  bumpCatalogRev_();

  return {
    ok: true,
    folder: {
      id: folderId,
      parentFolderId: parentFolderId,
      name: name,
      sizeBytes: folderStats.sizes[resolvedTargetId] || 0,
      fileCount: folderStats.fileCounts[resolvedTargetId] || 0,
      modifiedAt: formatCatalogDate_(now),
      isSystem: false,
      isMirror: true,
      isExternalMirror: false,
      mirrorOfFolderId: resolvedTargetId,
      mirrorOfDriveFolderId: '',
      editors: targetAcl.editors || [],
      commenters: targetAcl.commenters || [],
      readers: targetAcl.readers || []
    }
  };
}

/**
 * Резолв цепочки зеркал → конечная реальная папка.
 *
 * @param {Object.<string, Object>} foldersById
 * @param {string} folderId
 * @returns {string}
 */
function resolveMirrorTargetFolderId_(foldersById, folderId) {
  var id = String(folderId || '').trim();
  var seen = {};
  while (id && foldersById[id]) {
    if (seen[id]) {
      throw catalogError_('INVALID_INPUT', 'Цикл зеркал: ' + id);
    }
    seen[id] = true;
    var mirrorOf = String(foldersById[id].mirror_of_folder_id || '').trim();
    if (!mirrorOf) {
      return id;
    }
    id = mirrorOf;
  }
  if (!id || !foldersById[id]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Target folder not found.');
  }
  return id;
}

/**
 * Жёстко удалить строки Tree (зеркала) + ACL; не в корзину.
 *
 * @param {string[]} folderIds
 * @returns {number} removed count
 */
function deleteCatalogMirrorFolderRows_(folderIds) {
  var ids = {};
  (folderIds || []).forEach(function (id) {
    if (id) {
      ids[String(id)] = true;
    }
  });
  var idList = Object.keys(ids);
  if (!idList.length) {
    return 0;
  }

  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    return 0;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return 0;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('folder_id');
  if (idCol < 0) {
    return 0;
  }

  var keep = [values[0]];
  var removed = 0;
  var removedIds = [];
  for (var r = 1; r < values.length; r++) {
    var fid = String(values[r][idCol] || '');
    if (ids[fid]) {
      removed++;
      removedIds.push(fid);
      continue;
    }
    keep.push(values[r]);
  }
  if (!removed) {
    return 0;
  }
  sheet.clearContents();
  if (keep.length) {
    sheet.getRange(1, 1, keep.length, headers.length).setValues(keep);
  }
  clearAclRowsForObjects_(
    removedIds.map(function (id) {
      return { objectType: 'folder', objectId: id };
    })
  );
  return removed;
}

/**
 * Удалить все зеркала, указывающие на данные folder id (после резолва — точное совпадение поля).
 *
 * @param {string[]} targetFolderIds
 * @returns {number}
 */
function deleteMirrorsPointingToFolders_(targetFolderIds) {
  var targets = {};
  (targetFolderIds || []).forEach(function (id) {
    if (id) {
      targets[String(id)] = true;
    }
  });
  if (!Object.keys(targets).length) {
    return 0;
  }

  ensureCatalogSchemaUpToDate_();
  var rows = readSheetRecords_('Tree');
  var toDelete = [];
  rows.forEach(function (row) {
    var mirrorOf = String(row.mirror_of_folder_id || '').trim();
    if (mirrorOf && targets[mirrorOf]) {
      toDelete.push(String(row.folder_id));
    }
  });
  return deleteCatalogMirrorFolderRows_(toDelete);
}

/**
 * @param {Object} folderRow sheet row map
 * @returns {boolean}
 */
function isMirrorFolderRow_(folderRow) {
  return !!(
    folderRow &&
    (String(folderRow.mirror_of_folder_id || '').trim() ||
      String(folderRow.mirror_of_drive_folder_id || '').trim())
  );
}

/**
 * @param {Object} folderRow sheet row map
 * @returns {boolean}
 */
function isExternalMirrorFolderRow_(folderRow) {
  return !!(folderRow && String(folderRow.mirror_of_drive_folder_id || '').trim());
}
