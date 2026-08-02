/** @const {string} */
var TRASH_FOLDER_ID_ = '__TRASH__';

/**
 * §13.1, §14.6, §14.7 — перемещение файлов и папок между виртуальными папками (F6, F8 → корзина).
 *
 * @param {{
 *   targetFolderId: string,
 *   items: Array<{ kind: 'folder'|'file', id: string }>
 * }} input
 * @returns {{ ok: true, moved: Array<{ kind: 'folder'|'file', id: string }> }}
 */
function moveCatalogItems(input) {
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
  var requiredFolders = { targetFolderId: true };

  normalizedItems.forEach(function (item) {
    if (item.kind === 'file') {
      var file = engine.filesByCatalogId[item.id];
      if (!file) {
        throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + item.id);
      }
      if (isFileShortcutRow_(file)) {
        // Ярлык: только удаление ссылки — право редактора на родителя.
        if (file.folder_id) {
          requiredFolders[String(file.folder_id)] = true;
        }
        return;
      }
      requiredFolders[String(file.folder_id)] = true;
      return;
    }

    var folder = engine.foldersById[item.id];
    if (!folder) {
      throw catalogError_('FOLDER_NOT_FOUND', 'Folder not found: ' + item.id);
    }
    if (item.id === virtualRootFolderId) {
      throw catalogError_('NOT_ALLOWED', 'Cannot move the catalog root folder.');
    }
    if (item.id === TRASH_FOLDER_ID_) {
      throw catalogError_('NOT_ALLOWED', 'Cannot move the trash folder.');
    }
    if (parseBoolean_(folder.is_system)) {
      throw catalogError_('NOT_ALLOWED', 'Cannot move a system folder.');
    }
    if (isMirrorFolderRow_(folder)) {
      // Зеркало: только удаление ссылки — право редактора на родителя.
      if (folder.parent_folder_id) {
        requiredFolders[String(folder.parent_folder_id)] = true;
      }
      return;
    }
    if (isFolderInside_(engine, item.id, targetFolderId)) {
      throw catalogError_('INVALID_MOVE', 'Cannot move a folder into itself or its subfolder.');
    }
    requiredFolders[item.id] = true;
    if (folder.parent_folder_id) {
      requiredFolders[String(folder.parent_folder_id)] = true;
    }
  });

  Object.keys(requiredFolders).forEach(function (folderId) {
    assertEditorOnFolderForMove_(engine, userEmail, loginRole, folderId);
  });

  var fileUpdates = [];
  var folderUpdates = [];
  var moved = [];
  var mirrorsToDelete = [];
  var fileShortcutsToDelete = [];
  var foldersMovedToTrash = [];
  var filesMovedToTrash = [];

  normalizedItems.forEach(function (item) {
    if (item.kind === 'file') {
      var fileRow = engine.filesByCatalogId[item.id];
      if (isFileShortcutRow_(fileRow)) {
        // §22: ярлык файла — только жёсткое удаление (как зеркало папки).
        fileShortcutsToDelete.push(item.id);
        moved.push(item);
        return;
      }
      if (String(fileRow.folder_id) === targetFolderId) {
        moved.push(item);
        return;
      }
      fileUpdates.push({ catalogId: item.id, folderId: targetFolderId });
      moved.push(item);
      if (targetFolderId === TRASH_FOLDER_ID_) {
        filesMovedToTrash.push(item.id);
      }
      return;
    }

    var treeFolder = engine.foldersById[item.id];
    if (isMirrorFolderRow_(treeFolder)) {
      mirrorsToDelete.push(item.id);
      moved.push(item);
      return;
    }

    if (String(treeFolder.parent_folder_id || '') === targetFolderId) {
      moved.push(item);
      return;
    }
    folderUpdates.push({ folderId: item.id, parentFolderId: targetFolderId });
    moved.push(item);
    if (targetFolderId === TRASH_FOLDER_ID_) {
      foldersMovedToTrash.push(item.id);
    }
  });

  if (mirrorsToDelete.length) {
    deleteCatalogMirrorFolderRows_(mirrorsToDelete);
  }
  if (fileShortcutsToDelete.length) {
    deleteCatalogFileShortcutRows_(fileShortcutsToDelete);
  }

  var aclMoved = moved.filter(function (m) {
    return (
      mirrorsToDelete.indexOf(m.id) < 0 && fileShortcutsToDelete.indexOf(m.id) < 0
    );
  });

  // Только ярлыки / no-op — без Jobs.
  if (!fileUpdates.length && !folderUpdates.length) {
    if (mirrorsToDelete.length || fileShortcutsToDelete.length) {
      bumpCatalogRev_();
    }
    return {
      ok: true,
      queued: false,
      moved: moved,
      deletedMirrors: mirrorsToDelete,
      deletedFileShortcuts: fileShortcutsToDelete
    };
  }

  // §14.7 (2026-08-02): location + лёгкий кэш ACL sync; дельты не сбрасываем.
  // Jobs не нужны для обычного F6/F8 (убрали тяжёлый applyTargetFolderAclAfterMove_).
  applyFileFolderUpdates_(fileUpdates, folderUpdates);

  var trashCascadeTargets = [];
  foldersMovedToTrash.forEach(function (fid) {
    trashCascadeTargets.push(fid);
    collectFolderSubtreeObjects_(engine, fid).forEach(function (obj) {
      if (obj.objectType === 'folder') {
        trashCascadeTargets.push(obj.objectId);
      }
      if (obj.objectType === 'file') {
        filesMovedToTrash.push(obj.objectId);
      }
    });
  });

  if (trashCascadeTargets.length) {
    deleteMirrorsPointingToFolders_(trashCascadeTargets);
  }
  if (filesMovedToTrash.length) {
    deleteFileShortcutsPointingToCatalogIds_(filesMovedToTrash);
  }

  // Дельты ACL не трогаем. Кэш acl_* не переписываем здесь:
  // N× rewrite Tree/Files блокировал таблицу → «Загрузка» на F5; optimistic
  // раньше затирал права целью. Столбцы прав остаются как были; эффективные
  // при открытии/Доступ = новая мать ⊕ дельты (§14.7).
  bumpCatalogRev_();

  return {
    ok: true,
    queued: false,
    fileCount: fileUpdates.length + folderUpdates.length,
    moved: moved,
    deletedMirrors: mirrorsToDelete,
    deletedFileShortcuts: fileShortcutsToDelete
  };
}

/**
 * §14.7 — опц. лёгкий кэш после move (батч). Сейчас move его не вызывает.
 *
 * @param {Object} engine
 * @param {Array<{ kind: 'folder'|'file', id: string }>} movedItems
 */
function refreshAclCacheAfterMoveKeepDeltas_(engine, movedItems) {
  if (!movedItems || !movedItems.length) {
    return;
  }

  var targetKeys = {};
  var targets = [];

  function addTarget(objectType, objectId) {
    var key = objectType + ':' + objectId;
    if (targetKeys[key]) {
      return;
    }
    targetKeys[key] = true;
    targets.push({ objectType: objectType, objectId: objectId });
  }

  movedItems.forEach(function (item) {
    if (item.kind === 'file') {
      addTarget('file', item.id);
      return;
    }
    collectFolderSubtreeObjects_(engine, item.id).forEach(function (obj) {
      addTarget(obj.objectType, obj.objectId);
    });
  });

  var treeUpdates = [];
  var fileUpdatesAcl = [];
  targets.forEach(function (obj) {
    var entries = effectiveAclMapToEntries_(
      getEffectiveAclMapFromEngine_(engine, obj.objectType, obj.objectId)
    );
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
      fileUpdatesAcl.push({
        catalogId: obj.objectId,
        aclEditors: payload.aclEditors,
        aclCommenters: payload.aclCommenters,
        aclReaders: payload.aclReaders
      });
    }
  });

  writeTreeAclCacheBatch_(treeUpdates);
  writeFilesAclCacheBatch_(fileUpdatesAcl);
}

/**
 * @deprecated §14.7 2026-08-02 — сброс под цель убран; оставлен для старых Jobs move_catalog в очереди.
 * @param {Object} engine
 * @param {string} targetFolderId
 * @param {Array<{ kind: 'folder'|'file', id: string }>} movedItems
 */
function applyTargetFolderAclAfterMove_(engine, targetFolderId, movedItems) {
  refreshAclCacheAfterMoveKeepDeltas_(engine, movedItems);
}
/* OLD: applyTargetFolderAclAfterMove_ — clearAclRowsForObjects_ + sync кэша = эффективные целевой папки. */

/**
 * Эффективные ACL объекта → entries для кэша / UI.
 *
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {Array<{ principalType: string, principalId: string, permissionLevel: string }>}
 */
function buildAclEntriesFromObject_(engine, objectType, objectId) {
  return effectiveAclMapToEntries_(
    getEffectiveAclMapFromEngine_(engine, objectType, objectId)
  );
}

/**
 * @param {'user'|'manager'|'controller'} loginRole
 */
function assertCanRunCatalogOperations_(loginRole) {
  if (loginRole === 'manager' || loginRole === 'controller') {
    return;
  }
  throw catalogError_('NOT_ALLOWED', 'Операция доступна только Менеджеру и Управляющему.');
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} folderId
 */
function assertEditorOnFolderForMove_(engine, userEmail, loginRole, folderId) {
  if (loginRole === 'controller') {
    return;
  }
  if (!folderId) {
    return;
  }
  var permission = getEffectivePermissionForUserFromEngine_(
    engine,
    'folder',
    folderId,
    userEmail
  );
  if (permission !== 'editor') {
    throw catalogError_('NOT_ALLOWED', 'Нужно право редактор на папку.');
  }
}

/**
 * @param {Array} items
 * @returns {Array<{ kind: 'folder'|'file', id: string }>}
 */
function normalizeMoveItems_(items) {
  var normalized = [];
  var seen = {};

  for (var i = 0; i < items.length; i++) {
    var item = items[i] || {};
    var kind = String(item.kind || '').trim();
    var id = String(item.id || '').trim();
    if (kind !== 'folder' && kind !== 'file') {
      throw catalogError_('INVALID_INPUT', 'item.kind must be folder or file.');
    }
    if (!id) {
      throw catalogError_('INVALID_INPUT', 'item.id is required.');
    }
    var key = kind + ':' + id;
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    normalized.push({ kind: kind, id: id });
  }

  return normalized;
}

/**
 * @param {Object} engine
 * @param {string} ancestorFolderId
 * @param {string} folderId
 * @returns {boolean}
 */
function isFolderInside_(engine, ancestorFolderId, folderId) {
  if (!ancestorFolderId || !folderId) {
    return false;
  }
  if (ancestorFolderId === folderId) {
    return true;
  }

  var queue = [ancestorFolderId];
  var seen = {};

  while (queue.length) {
    var currentId = queue.shift();
    if (seen[currentId]) {
      continue;
    }
    seen[currentId] = true;

    var childIds = Object.keys(engine.foldersById);
    for (var i = 0; i < childIds.length; i++) {
      var childId = childIds[i];
      var child = engine.foldersById[childId];
      if (String(child.parent_folder_id || '') !== currentId) {
        continue;
      }
      if (childId === folderId) {
        return true;
      }
      queue.push(childId);
    }
  }

  return false;
}

/**
 * @param {Array<{ catalogId: string, folderId: string }>} fileUpdates
 * @param {Array<{ folderId: string, parentFolderId: string }>} folderUpdates
 */
function applyFileFolderUpdates_(fileUpdates, folderUpdates) {
  if (fileUpdates.length) {
    var filesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
    if (!filesSheet) {
      throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
    }
    var fileValues = filesSheet.getDataRange().getValues();
    if (fileValues.length < 2) {
      return;
    }
    var fileHeaders = fileValues[0].map(function (h) {
      return String(h).trim();
    });
    var catalogCol = fileHeaders.indexOf('catalog_id');
    var folderCol = fileHeaders.indexOf('folder_id');
    if (catalogCol < 0 || folderCol < 0) {
      throw catalogError_('SCHEMA_MISMATCH', 'Files sheet headers are invalid.');
    }

    var fileUpdateMap = {};
    fileUpdates.forEach(function (update) {
      fileUpdateMap[update.catalogId] = update.folderId;
    });

    var filePatches = [];
    for (var i = 1; i < fileValues.length; i++) {
      var catalogId = String(fileValues[i][catalogCol]);
      if (fileUpdateMap.hasOwnProperty(catalogId)) {
        filePatches.push({ row: i + 1, value: fileUpdateMap[catalogId] });
      }
    }
    if (filePatches.length) {
      if (filePatches.length > 50) {
        // Батч-путь: пишем весь столбец одним вызовом (быстрее для массовых обновлений).
        var folderColData = [];
        for (var fi = 1; fi < fileValues.length; fi++) {
          var cId = String(fileValues[fi][catalogCol]);
          folderColData.push([fileUpdateMap.hasOwnProperty(cId) ? fileUpdateMap[cId] : fileValues[fi][folderCol]]);
        }
        // getRange(row, column, numRows, numColumns) — 3-й/4-й аргументы = размеры, не lastRow/lastCol
        filesSheet
          .getRange(2, folderCol + 1, folderColData.length, 1)
          .setValues(folderColData);
      } else {
        // Точечный путь: setValue только для изменившихся строк.
        filePatches.forEach(function (p) {
          filesSheet.getRange(p.row, folderCol + 1).setValue(p.value);
        });
      }
    }
  }

  if (folderUpdates.length) {
    var treeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
    if (!treeSheet) {
      throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
    }
    var treeValues = treeSheet.getDataRange().getValues();
    if (treeValues.length < 2) {
      return;
    }
    var treeHeaders = treeValues[0].map(function (h) {
      return String(h).trim();
    });
    var folderIdCol = treeHeaders.indexOf('folder_id');
    var parentCol = treeHeaders.indexOf('parent_folder_id');
    if (folderIdCol < 0 || parentCol < 0) {
      throw catalogError_('SCHEMA_MISMATCH', 'Tree sheet headers are invalid.');
    }

    var folderUpdateMap = {};
    folderUpdates.forEach(function (update) {
      folderUpdateMap[update.folderId] = update.parentFolderId;
    });

    var treePatches = [];
    for (var j = 1; j < treeValues.length; j++) {
      var folderId = String(treeValues[j][folderIdCol]);
      if (folderUpdateMap.hasOwnProperty(folderId)) {
        treePatches.push({ row: j + 1, value: folderUpdateMap[folderId] });
      }
    }
    if (treePatches.length) {
      if (treePatches.length > 50) {
        // Батч-путь: весь столбец одним вызовом.
        var parentColData = [];
        for (var ti = 1; ti < treeValues.length; ti++) {
          var fId = String(treeValues[ti][folderIdCol]);
          parentColData.push([folderUpdateMap.hasOwnProperty(fId) ? folderUpdateMap[fId] : treeValues[ti][parentCol]]);
        }
        treeSheet
          .getRange(2, parentCol + 1, parentColData.length, 1)
          .setValues(parentColData);
      } else {
        // Точечный путь: setValue только для изменившихся строк.
        treePatches.forEach(function (p) {
          treeSheet.getRange(p.row, parentCol + 1).setValue(p.value);
        });
      }
    }
  }
}
