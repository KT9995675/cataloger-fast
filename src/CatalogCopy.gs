/**
 * §13.1, §14.7 — копирование файлов и папок между виртуальными папками (F5).
 * На Drive создаётся вторая физическая копия каждого файла (§13.1).
 *
 * @param {{
 *   targetFolderId: string,
 *   items: Array<{ kind: 'folder'|'file', id: string }>
 * }} input
 * @returns {{
 *   ok: true,
 *   copied: Array<{ kind: 'folder'|'file', sourceId: string, newId: string }>
 * }}
 */
function copyCatalogItems(input) {
  assertCatalogReady_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var items = input.items || [];

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (!Array.isArray(items) || !items.length) {
    throw catalogError_('INVALID_INPUT', 'items must be a non-empty array.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var engine = createAclEngine_();
  if (!engine.foldersById[targetFolderId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Target folder not found: ' + targetFolderId);
  }

  var virtualRootFolderId = getVirtualRootFolderId_();
  var normalizedItems = normalizeMoveItems_(items);
  assertCopyPermissions_(engine, userEmail, loginRole, targetFolderId, normalizedItems, virtualRootFolderId);

  normalizedItems.forEach(function (item) {
    if (item.kind === 'folder' && isFolderInside_(engine, item.id, targetFolderId)) {
      throw catalogError_(
        'INVALID_COPY',
        'Cannot copy a folder into itself or its subfolder.'
      );
    }
  });

  var catalogRootFolder = DriveApp.getFolderById(getCatalogRootFolderId_());
  var copied = [];

  normalizedItems.forEach(function (item) {
    if (item.kind === 'file') {
      var fileCopy = copyCatalogFile_(engine, item.id, targetFolderId, catalogRootFolder);
      copied.push({
        kind: 'file',
        sourceId: item.id,
        newId: fileCopy.catalogId
      });
      return;
    }

    var folderCopy = copyCatalogFolder_(engine, item.id, targetFolderId, catalogRootFolder);
    copied.push({
      kind: 'folder',
      sourceId: item.id,
      newId: folderCopy.rootFolderId
    });
  });

  return {
    ok: true,
    copied: copied
  };
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} targetFolderId
 * @param {Array<{ kind: string, id: string }>} items
 * @param {string} virtualRootFolderId
 */
function assertCopyPermissions_(
  engine,
  userEmail,
  loginRole,
  targetFolderId,
  items,
  virtualRootFolderId
) {
  var requiredFolders = {};
  requiredFolders[targetFolderId] = true;

  items.forEach(function (item) {
    if (item.kind === 'file') {
      var file = engine.filesByCatalogId[item.id];
      if (!file) {
        throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + item.id);
      }
      requiredFolders[String(file.folder_id)] = true;
      return;
    }

    var folder = engine.foldersById[item.id];
    if (!folder) {
      throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + item.id);
    }
    if (item.id === virtualRootFolderId) {
      throw catalogError_('NOT_ALLOWED', 'Cannot copy the catalog root folder.');
    }
    if (item.id === TRASH_FOLDER_ID_) {
      throw catalogError_('NOT_ALLOWED', 'Cannot copy the trash folder.');
    }
    if (parseBoolean_(folder.is_system)) {
      throw catalogError_('NOT_ALLOWED', 'Cannot copy a system folder.');
    }
    requiredFolders[item.id] = true;
    if (folder.parent_folder_id) {
      requiredFolders[String(folder.parent_folder_id)] = true;
    }
  });

  Object.keys(requiredFolders).forEach(function (folderId) {
    assertEditorOnFolderForMove_(engine, userEmail, loginRole, folderId);
  });
}

/**
 * @param {Object} engine
 * @param {string} sourceCatalogId
 * @param {string} targetFolderId
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {{ catalogId: string }}
 */
function copyCatalogFile_(engine, sourceCatalogId, targetFolderId, catalogRootFolder) {
  var source = engine.filesByCatalogId[sourceCatalogId];
  if (!source) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + sourceCatalogId);
  }

  var sourceDriveFileId = String(source.file_id || '').trim();
  if (!sourceDriveFileId) {
    throw catalogError_('FILE_NOT_READY', 'Source file is not on Drive yet.');
  }
  if (String(source.status || 'ready').toLowerCase() === 'pending') {
    throw catalogError_('FILE_NOT_READY', 'Source file is still pending.');
  }

  var sourceDriveFile = DriveApp.getFileById(sourceDriveFileId);
  var driveCopy = sourceDriveFile.makeCopy(String(source.display_name || sourceDriveFile.getName()), catalogRootFolder);
  var newCatalogId = Utilities.getUuid();

  appendCatalogFileCopyRow_(source, {
    catalogId: newCatalogId,
    folderId: targetFolderId,
    fileId: driveCopy.getId(),
    displayName: String(source.display_name || driveCopy.getName()),
    sizeBytes: driveCopy.getSize(),
    driveModifiedAt: driveCopy.getLastUpdated(),
    sourceFileId: sourceDriveFileId,
    mimeType: getDriveFileMimeType_(driveCopy) || source.mime_type || ''
  });

  engine.filesByCatalogId[newCatalogId] = {
    catalog_id: newCatalogId,
    folder_id: targetFolderId,
    file_id: driveCopy.getId(),
    display_name: String(source.display_name || driveCopy.getName()),
    approved: source.approved,
    approved_by: source.approved_by
  };

  copyEffectiveAcl_(engine, 'file', sourceCatalogId, 'file', newCatalogId);

  return { catalogId: newCatalogId };
}

/**
 * @param {Object} engine
 * @param {string} sourceFolderId
 * @param {string} targetParentFolderId
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {{ rootFolderId: string }}
 */
function copyCatalogFolder_(engine, sourceFolderId, targetParentFolderId, catalogRootFolder) {
  var sourceFolder = engine.foldersById[sourceFolderId];
  if (!sourceFolder) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + sourceFolderId);
  }

  var subtree = collectFolderSubtreeObjects_(engine, sourceFolderId);
  var folderIdMap = {};
  var now = new Date();

  subtree.forEach(function (obj) {
    if (obj.objectType !== 'folder') {
      return;
    }
    folderIdMap[obj.objectId] = Utilities.getUuid();
  });

  var treeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!treeSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  subtree.forEach(function (obj) {
    if (obj.objectType !== 'folder') {
      return;
    }
    var source = engine.foldersById[obj.objectId];
    var newFolderId = folderIdMap[obj.objectId];
    var newParentId =
      obj.objectId === sourceFolderId
        ? targetParentFolderId
        : folderIdMap[String(source.parent_folder_id)];

    treeSheet.appendRow([
      newFolderId,
      newParentId,
      source.name,
      now,
      false
    ]);

    engine.foldersById[newFolderId] = {
      folder_id: newFolderId,
      parent_folder_id: newParentId,
      name: source.name,
      is_system: false
    };

    copyEffectiveAcl_(engine, 'folder', obj.objectId, 'folder', newFolderId);
  });

  subtree.forEach(function (obj) {
    if (obj.objectType !== 'file') {
      return;
    }
    var targetFolderId = folderIdMap[String(engine.filesByCatalogId[obj.objectId].folder_id)];
    copyCatalogFile_(engine, obj.objectId, targetFolderId, catalogRootFolder);
  });

  return { rootFolderId: folderIdMap[sourceFolderId] };
}

/**
 * @param {Object} sourceRow
 * @param {{
 *   catalogId: string,
 *   folderId: string,
 *   fileId: string,
 *   displayName: string,
 *   sizeBytes: number,
 *   driveModifiedAt: Date,
 *   sourceFileId: string
 * }} target
 */
function appendCatalogFileCopyRow_(sourceRow, target) {
  appendCatalogFileRow_({
    catalogId: target.catalogId,
    folderId: target.folderId,
    fileId: target.fileId,
    displayName: target.displayName,
    sizeBytes: target.sizeBytes,
    driveModifiedAt: target.driveModifiedAt,
    approved: parseBoolean_(sourceRow.approved),
    approvedBy: sourceRow.approved_by || '',
    approvedAt: sourceRow.approved_at || '',
    sourceFileId: target.sourceFileId || '',
    mimeType: target.mimeType || sourceRow.mime_type || ''
  });
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} sourceType
 * @param {string} sourceId
 * @param {'folder'|'file'} targetType
 * @param {string} targetId
 */
function copyEffectiveAcl_(engine, sourceType, sourceId, targetType, targetId) {
  var aclRows = resolveInheritedAclRows_(engine, sourceType, sourceId);
  if (!aclRows.length) {
    return;
  }

  var entries = [];
  for (var i = 0; i < aclRows.length; i++) {
    entries.push({
      principalType: String(aclRows[i].principal_type || '').trim(),
      principalId: String(aclRows[i].principal_id || '').trim(),
      permissionLevel: normalizePermissionLevel_(aclRows[i].permission_level)
    });
  }

  replaceAclForObjects_([{ objectType: targetType, objectId: targetId }], entries, engine);
}
