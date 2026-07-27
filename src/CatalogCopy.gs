/**
 * §13.1, §14.7 — копирование файлов и папок между виртуальными папками (F5).
 * На Drive создаётся вторая физическая копия каждого файла (§13.1).
 * ≤1 файл — sync (batch Tree/Files/ACL); ≥2 файлов — Jobs `copy_catalog` (`Копирование: N/M`).
 *
 * @param {{
 *   targetFolderId: string,
 *   items: Array<{ kind: 'folder'|'file', id: string }>
 * }} input
 * @returns {{
 *   ok: true,
 *   queued?: boolean,
 *   jobId?: string,
 *   fileCount?: number,
 *   copied: Array<{ kind: 'folder'|'file', sourceId: string, newId: string }>,
 *   created?: {
 *     folders: Array<Object>,
 *     files: Array<Object>
 *   }
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
  var fileCount = countCopyFileTargets_(engine, normalizedItems);

  if (fileCount > COPY_SYNC_MAX_FILES_) {
    return enqueueCopyCatalogJob_(engine, targetFolderId, normalizedItems, userEmail);
  }

  return copyCatalogItemsSync_(engine, targetFolderId, normalizedItems, catalogRootFolder);
}

/** @const {number} Больше этого числа файлов → Jobs. */
var COPY_SYNC_MAX_FILES_ = 1;

/**
 * @param {Object} engine
 * @param {Array<{ kind: string, id: string }>} items
 * @returns {number}
 */
function countCopyFileTargets_(engine, items) {
  var n = 0;
  items.forEach(function (item) {
    if (item.kind === 'file') {
      n += 1;
      return;
    }
    collectFolderSubtreeObjects_(engine, item.id).forEach(function (obj) {
      if (obj.objectType === 'file') {
        n += 1;
      }
    });
  });
  return n;
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
 * @returns {{
 *   treeRows: Array<Object>,
 *   fileRows: Array<Object>,
 *   aclList: Array<{ objectType: string, objectId: string, entries: Array }>
 * }}
 */
function beginCopyWriteContext_() {
  return {
    treeRows: [],
    fileRows: [],
    aclList: []
  };
}

/**
 * @param {Object} ctx
 * @param {Object} engine
 */
function commitCopyWriteContext_(ctx, engine) {
  appendTreeFolderRowsBatch_(ctx.treeRows);
  appendCatalogFileRowsBatch_(ctx.fileRows);
  appendAclForNewObjectsBatch_(ctx.aclList, engine);
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} sourceType
 * @param {string} sourceId
 * @returns {Array<{ principalType: string, principalId: string, permissionLevel: string }>}
 */
function buildCopyAclEntries_(engine, sourceType, sourceId) {
  var aclRows = resolveInheritedAclRows_(engine, sourceType, sourceId);
  var entries = [];
  for (var i = 0; i < aclRows.length; i++) {
    entries.push({
      principalType: String(aclRows[i].principal_type || '').trim(),
      principalId: String(aclRows[i].principal_id || '').trim(),
      permissionLevel: normalizePermissionLevel_(aclRows[i].permission_level)
    });
  }
  return entries;
}

/**
 * @param {Object} ctx
 * @param {Object} engine
 * @param {'folder'|'file'} sourceType
 * @param {string} sourceId
 * @param {'folder'|'file'} targetType
 * @param {string} targetId
 * @returns {Array}
 */
function queueCopyAcl_(ctx, engine, sourceType, sourceId, targetType, targetId) {
  var entries = buildCopyAclEntries_(engine, sourceType, sourceId);
  if (entries.length) {
    ctx.aclList.push({
      objectType: targetType,
      objectId: targetId,
      entries: entries
    });
  }
  return entries;
}

/**
 * @param {Object} engine
 * @param {Array} entries
 * @param {boolean} approved
 * @returns {{ editors: string[], commenters: string[], readers: string[] }}
 */
function uiAclFromCopyEntries_(engine, entries, approved) {
  var rows = (entries || []).map(function (e) {
    return {
      principal_type: e.principalType,
      principal_id: e.principalId,
      permission_level: e.permissionLevel
    };
  });
  var labels = aclRowsToCacheLabels_(engine, rows, !!approved);
  return {
    editors: labels.editors || [],
    commenters: labels.commenters || [],
    readers: labels.readers || []
  };
}

/**
 * Sync-копирование (≤ COPY_SYNC_MAX_FILES_ файлов): Drive makeCopy + один batch листов.
 *
 * @param {Object} engine
 * @param {string} targetFolderId
 * @param {Array<{ kind: string, id: string }>} items
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {Object}
 */
function copyCatalogItemsSync_(engine, targetFolderId, items, catalogRootFolder) {
  var ctx = beginCopyWriteContext_();
  var copied = [];
  var created = { folders: [], files: [] };

  items.forEach(function (item) {
    if (item.kind === 'file') {
      var fileCopy = prepareCopyFileReady_(
        engine,
        ctx,
        item.id,
        targetFolderId,
        catalogRootFolder
      );
      copied.push({
        kind: 'file',
        sourceId: item.id,
        newId: fileCopy.catalogId
      });
      created.files.push(fileCopy.uiItem);
      return;
    }

    var folderCopy = prepareCopyFolderReady_(
      engine,
      ctx,
      item.id,
      targetFolderId,
      catalogRootFolder
    );
    copied.push({
      kind: 'folder',
      sourceId: item.id,
      newId: folderCopy.rootFolderId
    });
    created.folders = created.folders.concat(folderCopy.folders);
    created.files = created.files.concat(folderCopy.files);
  });

  commitCopyWriteContext_(ctx, engine);
  bumpCatalogRev_();

  return {
    ok: true,
    queued: false,
    copied: copied,
    created: created
  };
}

/**
 * ≥2 файлов: папки + pending Files + ACL сразу; Drive makeCopy — в Jobs.
 *
 * @param {Object} engine
 * @param {string} targetFolderId
 * @param {Array<{ kind: string, id: string }>} items
 * @param {string} userEmail
 * @returns {Object}
 */
function enqueueCopyCatalogJob_(engine, targetFolderId, items, userEmail) {
  assertNoActiveCatalogJobs_();

  var ctx = beginCopyWriteContext_();
  var copied = [];
  var created = { folders: [], files: [] };
  var workItems = [];

  items.forEach(function (item) {
    if (item.kind === 'file') {
      var filePlan = prepareCopyFilePending_(engine, ctx, item.id, targetFolderId);
      copied.push({
        kind: 'file',
        sourceId: item.id,
        newId: filePlan.catalogId
      });
      created.files.push(filePlan.uiItem);
      workItems.push(filePlan.workItem);
      return;
    }

    var folderPlan = prepareCopyFolderPending_(engine, ctx, item.id, targetFolderId);
    copied.push({
      kind: 'folder',
      sourceId: item.id,
      newId: folderPlan.rootFolderId
    });
    created.folders = created.folders.concat(folderPlan.folders);
    created.files = created.files.concat(folderPlan.files);
    workItems = workItems.concat(folderPlan.workItems);
  });

  commitCopyWriteContext_(ctx, engine);

  var jobId = enqueueCatalogJob_(
    'copy_catalog',
    {
      targetFolderId: targetFolderId,
      items: workItems
    },
    userEmail,
    ''
  );
  kickCatalogJobsProcessing_();
  bumpCatalogRev_();

  return {
    ok: true,
    queued: true,
    jobId: jobId,
    fileCount: workItems.length,
    copied: copied,
    created: created
  };
}

/**
 * @param {Object} engine
 * @param {Object} ctx
 * @param {string} sourceCatalogId
 * @param {string} targetFolderId
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {{ catalogId: string, uiItem: Object }}
 */
function prepareCopyFileReady_(engine, ctx, sourceCatalogId, targetFolderId, catalogRootFolder) {
  var source = assertCopySourceFileReady_(engine, sourceCatalogId);
  var sourceDriveFileId = String(source.file_id || '').trim();
  var sourceDriveFile = DriveApp.getFileById(sourceDriveFileId);
  var driveCopy = sourceDriveFile.makeCopy(
    String(source.display_name || sourceDriveFile.getName()),
    catalogRootFolder
  );
  var newCatalogId = Utilities.getUuid();
  var displayName = String(source.display_name || driveCopy.getName());
  var mimeType = getDriveFileMimeType_(driveCopy) || source.mime_type || '';
  var approved = parseBoolean_(source.approved);
  var sizeBytes = driveCopy.getSize();
  var modifiedAt = driveCopy.getLastUpdated();

  ctx.fileRows.push({
    catalogId: newCatalogId,
    folderId: targetFolderId,
    fileId: driveCopy.getId(),
    displayName: displayName,
    sizeBytes: sizeBytes,
    driveModifiedAt: modifiedAt,
    approved: approved,
    approvedBy: source.approved_by || '',
    approvedAt: source.approved_at || '',
    status: 'ready',
    sourceFileId: sourceDriveFileId,
    mimeType: mimeType
  });

  engine.filesByCatalogId[newCatalogId] = {
    catalog_id: newCatalogId,
    folder_id: targetFolderId,
    file_id: driveCopy.getId(),
    display_name: displayName,
    approved: source.approved,
    approved_by: source.approved_by
  };

  var entries = queueCopyAcl_(ctx, engine, 'file', sourceCatalogId, 'file', newCatalogId);
  var acl = uiAclFromCopyEntries_(engine, entries, approved);

  return {
    catalogId: newCatalogId,
    uiItem: {
      id: newCatalogId,
      folderId: targetFolderId,
      name: displayName,
      mimeType: mimeType,
      sizeBytes: sizeBytes,
      modifiedAt: formatCatalogDate_(modifiedAt),
      approved: approved,
      approvedBy: source.approved_by || '',
      approvedByName: '',
      editors: acl.editors,
      commenters: acl.commenters,
      readers: acl.readers
    }
  };
}

/**
 * @param {Object} engine
 * @param {Object} ctx
 * @param {string} sourceCatalogId
 * @param {string} targetFolderId
 * @returns {{
 *   catalogId: string,
 *   uiItem: Object,
 *   workItem: Object
 * }}
 */
function prepareCopyFilePending_(engine, ctx, sourceCatalogId, targetFolderId) {
  var source = assertCopySourceFileReady_(engine, sourceCatalogId);
  var sourceDriveFileId = String(source.file_id || '').trim();
  var newCatalogId = Utilities.getUuid();
  var displayName = String(source.display_name || '');
  var mimeType = String(source.mime_type || '');
  var approved = parseBoolean_(source.approved);

  ctx.fileRows.push({
    catalogId: newCatalogId,
    folderId: targetFolderId,
    fileId: '',
    displayName: displayName,
    sizeBytes: 0,
    driveModifiedAt: '',
    approved: approved,
    approvedBy: source.approved_by || '',
    approvedAt: source.approved_at || '',
    status: 'pending',
    sourceFileId: sourceDriveFileId,
    mimeType: mimeType
  });

  engine.filesByCatalogId[newCatalogId] = {
    catalog_id: newCatalogId,
    folder_id: targetFolderId,
    file_id: '',
    display_name: displayName,
    approved: source.approved,
    approved_by: source.approved_by
  };

  var entries = queueCopyAcl_(ctx, engine, 'file', sourceCatalogId, 'file', newCatalogId);
  var acl = uiAclFromCopyEntries_(engine, entries, approved);

  return {
    catalogId: newCatalogId,
    uiItem: {
      id: newCatalogId,
      folderId: targetFolderId,
      name: displayName,
      mimeType: mimeType,
      sizeBytes: 0,
      modifiedAt: '',
      approved: approved,
      approvedBy: source.approved_by || '',
      approvedByName: '',
      editors: acl.editors,
      commenters: acl.commenters,
      readers: acl.readers
    },
    workItem: {
      catalogId: newCatalogId,
      sourceDriveFileId: sourceDriveFileId,
      displayName: displayName,
      done: false
    }
  };
}

/**
 * @param {Object} engine
 * @param {string} sourceCatalogId
 * @returns {Object}
 */
function assertCopySourceFileReady_(engine, sourceCatalogId) {
  var source = engine.filesByCatalogId[sourceCatalogId];
  if (!source) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + sourceCatalogId);
  }
  if (!String(source.file_id || '').trim()) {
    throw catalogError_('FILE_NOT_READY', 'Source file is not on Drive yet.');
  }
  if (String(source.status || 'ready').toLowerCase() === 'pending') {
    throw catalogError_('FILE_NOT_READY', 'Source file is still pending.');
  }
  return source;
}

/**
 * @param {Object} engine
 * @param {Object} ctx
 * @param {string} sourceFolderId
 * @param {string} targetParentFolderId
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {{ rootFolderId: string, folders: Array, files: Array }}
 */
function prepareCopyFolderReady_(
  engine,
  ctx,
  sourceFolderId,
  targetParentFolderId,
  catalogRootFolder
) {
  var mapped = mapCopyFolderTree_(engine, ctx, sourceFolderId, targetParentFolderId);
  var files = [];

  mapped.subtree.forEach(function (obj) {
    if (obj.objectType !== 'file') {
      return;
    }
    var targetFolderId = mapped.folderIdMap[String(engine.filesByCatalogId[obj.objectId].folder_id)];
    var fileCopy = prepareCopyFileReady_(
      engine,
      ctx,
      obj.objectId,
      targetFolderId,
      catalogRootFolder
    );
    files.push(fileCopy.uiItem);
  });

  return {
    rootFolderId: mapped.rootFolderId,
    folders: mapped.folders,
    files: files
  };
}

/**
 * @param {Object} engine
 * @param {Object} ctx
 * @param {string} sourceFolderId
 * @param {string} targetParentFolderId
 * @returns {{
 *   rootFolderId: string,
 *   folders: Array,
 *   files: Array,
 *   workItems: Array
 * }}
 */
function prepareCopyFolderPending_(engine, ctx, sourceFolderId, targetParentFolderId) {
  var mapped = mapCopyFolderTree_(engine, ctx, sourceFolderId, targetParentFolderId);
  var files = [];
  var workItems = [];

  mapped.subtree.forEach(function (obj) {
    if (obj.objectType !== 'file') {
      return;
    }
    var targetFolderId = mapped.folderIdMap[String(engine.filesByCatalogId[obj.objectId].folder_id)];
    var filePlan = prepareCopyFilePending_(engine, ctx, obj.objectId, targetFolderId);
    files.push(filePlan.uiItem);
    workItems.push(filePlan.workItem);
  });

  return {
    rootFolderId: mapped.rootFolderId,
    folders: mapped.folders,
    files: files,
    workItems: workItems
  };
}

/**
 * Создаёт новые folder_id для поддерева, ставит Tree-строки и ACL в ctx.
 *
 * @param {Object} engine
 * @param {Object} ctx
 * @param {string} sourceFolderId
 * @param {string} targetParentFolderId
 * @returns {{
 *   rootFolderId: string,
 *   folderIdMap: Object.<string, string>,
 *   subtree: Array,
 *   folders: Array
 * }}
 */
function mapCopyFolderTree_(engine, ctx, sourceFolderId, targetParentFolderId) {
  var sourceFolder = engine.foldersById[sourceFolderId];
  if (!sourceFolder) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + sourceFolderId);
  }

  var subtree = collectFolderSubtreeObjects_(engine, sourceFolderId);
  var folderIdMap = {};
  var now = new Date();
  var folders = [];

  subtree.forEach(function (obj) {
    if (obj.objectType !== 'folder') {
      return;
    }
    folderIdMap[obj.objectId] = Utilities.getUuid();
  });

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

    ctx.treeRows.push({
      folderId: newFolderId,
      parentFolderId: newParentId,
      name: source.name,
      folderCreatedAt: now,
      isSystem: false
    });

    engine.foldersById[newFolderId] = {
      folder_id: newFolderId,
      parent_folder_id: newParentId,
      name: source.name,
      is_system: false
    };

    var entries = queueCopyAcl_(ctx, engine, 'folder', obj.objectId, 'folder', newFolderId);
    var acl = uiAclFromCopyEntries_(engine, entries, false);
    folders.push({
      id: newFolderId,
      parentFolderId: newParentId,
      name: source.name,
      sizeBytes: 0,
      modifiedAt: formatCatalogDate_(now),
      isSystem: false,
      editors: acl.editors,
      commenters: acl.commenters,
      readers: acl.readers
    });
  });

  return {
    rootFolderId: folderIdMap[sourceFolderId],
    folderIdMap: folderIdMap,
    subtree: subtree,
    folders: folders
  };
}
