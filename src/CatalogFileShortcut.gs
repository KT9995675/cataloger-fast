/**
 * §22 — ярлыки файлов (синие в каталоге + красные на Drive).
 */

/**
 * Ярлык файла каталога (синий).
 *
 * @param {{
 *   parentFolderId: string,
 *   targetCatalogId: string
 * }} input
 * @returns {{ ok: true, item: Object }}
 */
function createCatalogFileShortcut(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var targetCatalogId = String(input.targetCatalogId || '').trim();
  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!targetCatalogId) {
    throw catalogError_('INVALID_INPUT', 'Выберите файл для ярлыка.');
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

  var resolvedTargetId = resolveFileShortcutTargetCatalogId_(engine.filesByCatalogId, targetCatalogId);
  var target = engine.filesByCatalogId[resolvedTargetId];
  if (!target) {
    throw catalogError_('FILE_NOT_FOUND', 'Target file not found: ' + targetCatalogId);
  }
  if (isFileShortcutRow_(target)) {
    throw catalogError_('INVALID_INPUT', 'Цель ярлыка должна быть обычным файлом.');
  }

  var name = String(target.display_name || '').trim();
  if (!name) {
    throw catalogError_('INVALID_INPUT', 'У целевого файла нет имени.');
  }

  var catalogId = Utilities.getUuid();
  var now = new Date();
  appendCatalogFileRowsBatch_([
    {
      catalogId: catalogId,
      folderId: parentFolderId,
      fileId: '',
      displayName: name,
      sizeBytes: parseNumber_(target.size_bytes) || 0,
      driveModifiedAt: target.drive_modified_at || now,
      approved: false,
      status: 'ready',
      mimeType: target.mime_type || '',
      aclEditors: '',
      aclCommenters: '',
      aclReaders: '',
      shortcutOfCatalogId: resolvedTargetId,
      shortcutOfDriveFileId: ''
    }
  ]);
  bumpCatalogRev_();

  var targetAcl = getEffectiveAclDisplayFromEngine_(engine, 'file', resolvedTargetId);
  return {
    ok: true,
    item: {
      id: catalogId,
      folderId: parentFolderId,
      name: name,
      mimeType: target.mime_type || '',
      sizeBytes: parseNumber_(target.size_bytes) || 0,
      modifiedAt: formatCatalogDate_(target.drive_modified_at || now),
      approved: false,
      approvedBy: '',
      approvedByName: '',
      status: 'ready',
      isShortcut: true,
      isExternalShortcut: false,
      shortcutOfCatalogId: resolvedTargetId,
      shortcutOfDriveFileId: '',
      editors: targetAcl.editors || [],
      commenters: targetAcl.commenters || [],
      readers: targetAcl.readers || []
    }
  };
}

/**
 * Ярлык внешнего файла Drive (красный).
 *
 * @param {{
 *   parentFolderId: string,
 *   driveFileId: string
 * }} input
 * @returns {{ ok: true, item: Object }}
 */
function createCatalogExternalFileShortcut(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var driveFileId = String(input.driveFileId || '').trim();
  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!driveFileId) {
    throw catalogError_('INVALID_INPUT', 'Выберите файл Drive для ярлыка.');
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

  var token = ScriptApp.getOAuthToken();
  var meta = fetchExternalDriveFileMeta_(token, driveFileId);
  if (!meta || !meta.id) {
    throw catalogError_('FILE_NOT_FOUND', 'Файл Drive не найден или нет доступа.');
  }
  if (String(meta.mimeType || '') === DRIVE_FOLDER_MIME_) {
    throw catalogError_('INVALID_INPUT', 'Для папки используйте ярлык внешней папки.');
  }

  var name = String(meta.name || '').trim() || driveFileId;
  var catalogId = Utilities.getUuid();
  var now = new Date();
  var sizeBytes = meta.size != null && meta.size !== '' ? parseNumber_(meta.size) : 0;
  appendCatalogFileRowsBatch_([
    {
      catalogId: catalogId,
      folderId: parentFolderId,
      fileId: '',
      displayName: name,
      sizeBytes: sizeBytes,
      driveModifiedAt: meta.modifiedTime || now,
      approved: false,
      status: 'ready',
      mimeType: meta.mimeType || '',
      aclEditors: '',
      aclCommenters: '',
      aclReaders: '',
      shortcutOfCatalogId: '',
      shortcutOfDriveFileId: String(meta.id)
    }
  ]);
  bumpCatalogRev_();

  return {
    ok: true,
    item: {
      id: catalogId,
      folderId: parentFolderId,
      name: name,
      mimeType: meta.mimeType || '',
      sizeBytes: sizeBytes,
      modifiedAt: formatCatalogDate_(meta.modifiedTime || now),
      approved: false,
      approvedBy: '',
      approvedByName: '',
      status: 'ready',
      isShortcut: true,
      isExternalShortcut: true,
      shortcutOfCatalogId: '',
      shortcutOfDriveFileId: String(meta.id),
      openUrl: buildCatalogFileOpenUrl_(
        String(meta.id),
        meta.mimeType || '',
        meta.webViewLink || ''
      ),
      editors: [],
      commenters: [],
      readers: []
    }
  };
}

/**
 * @param {Object.<string, Object>} filesByCatalogId
 * @param {string} catalogId
 * @returns {string}
 */
function resolveFileShortcutTargetCatalogId_(filesByCatalogId, catalogId) {
  var id = String(catalogId || '').trim();
  var seen = {};
  while (id && filesByCatalogId[id]) {
    if (seen[id]) {
      throw catalogError_('INVALID_INPUT', 'Цикл ярлыков файлов: ' + id);
    }
    seen[id] = true;
    var shortcutOf = String(filesByCatalogId[id].shortcut_of_catalog_id || '').trim();
    if (!shortcutOf) {
      return id;
    }
    id = shortcutOf;
  }
  if (!id || !filesByCatalogId[id]) {
    throw catalogError_('FILE_NOT_FOUND', 'Target file not found.');
  }
  return id;
}

/**
 * @param {Object} fileRow
 * @returns {boolean}
 */
function isFileShortcutRow_(fileRow) {
  return !!(
    fileRow &&
    (String(fileRow.shortcut_of_catalog_id || '').trim() ||
      String(fileRow.shortcut_of_drive_file_id || '').trim())
  );
}

/**
 * @param {Object} fileRow
 * @returns {boolean}
 */
function isExternalFileShortcutRow_(fileRow) {
  return !!(fileRow && String(fileRow.shortcut_of_drive_file_id || '').trim());
}

/**
 * Жёстко удалить строки Files (ярлыки) + ACL.
 *
 * @param {string[]} catalogIds
 * @returns {number}
 */
function deleteCatalogFileShortcutRows_(catalogIds) {
  var ids = {};
  (catalogIds || []).forEach(function (id) {
    if (id) {
      ids[String(id)] = true;
    }
  });
  var idList = Object.keys(ids);
  if (!idList.length) {
    return 0;
  }

  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
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
  var idCol = headers.indexOf('catalog_id');
  if (idCol < 0) {
    return 0;
  }

  var keep = [values[0]];
  var removed = 0;
  var removedIds = [];
  for (var r = 1; r < values.length; r++) {
    var cid = String(values[r][idCol] || '');
    if (ids[cid]) {
      removed++;
      removedIds.push(cid);
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
      return { objectType: 'file', objectId: id };
    })
  );
  return removed;
}

/**
 * Удалить ярлыки, указывающие на данные catalog_id целей.
 *
 * @param {string[]} targetCatalogIds
 * @returns {number}
 */
function deleteFileShortcutsPointingToCatalogIds_(targetCatalogIds) {
  var targets = {};
  (targetCatalogIds || []).forEach(function (id) {
    if (id) {
      targets[String(id)] = true;
    }
  });
  if (!Object.keys(targets).length) {
    return 0;
  }
  ensureCatalogSchemaUpToDate_();
  var rows = readSheetRecords_('Files');
  var toDelete = [];
  rows.forEach(function (row) {
    var shortcutOf = String(row.shortcut_of_catalog_id || '').trim();
    if (shortcutOf && targets[shortcutOf]) {
      toDelete.push(String(row.catalog_id));
    }
  });
  return deleteCatalogFileShortcutRows_(toDelete);
}

/**
 * @param {string} token
 * @param {string} fileId
 * @returns {Object}
 */
function fetchExternalDriveFileMeta_(token, fileId) {
  return driveImportFetchJson_(
    'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '?fields=' +
      encodeURIComponent('id,name,mimeType,size,modifiedTime,webViewLink') +
      '&supportsAllDrives=true',
    token
  );
}
