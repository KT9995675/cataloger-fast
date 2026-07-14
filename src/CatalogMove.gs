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

  normalizedItems.forEach(function (item) {
    if (item.kind === 'file') {
      var fileRow = engine.filesByCatalogId[item.id];
      if (String(fileRow.folder_id) === targetFolderId) {
        moved.push(item);
        return;
      }
      fileUpdates.push({ catalogId: item.id, folderId: targetFolderId });
      moved.push(item);
      return;
    }

    var treeFolder = engine.foldersById[item.id];
    if (String(treeFolder.parent_folder_id || '') === targetFolderId) {
      moved.push(item);
      return;
    }
    folderUpdates.push({ folderId: item.id, parentFolderId: targetFolderId });
    moved.push(item);
  });

  applyFileFolderUpdates_(fileUpdates, folderUpdates);

  return {
    ok: true,
    moved: moved
  };
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

    var folderColData = [];
    var changed = false;
    for (var i = 1; i < fileValues.length; i++) {
      var catalogId = String(fileValues[i][catalogCol]);
      var value = fileValues[i][folderCol];
      if (fileUpdateMap.hasOwnProperty(catalogId)) {
        value = fileUpdateMap[catalogId];
        changed = true;
      }
      folderColData.push([value]);
    }
    if (changed) {
      filesSheet
        .getRange(2, folderCol + 1, fileValues.length, folderCol + 1)
        .setValues(folderColData);
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

    var parentColData = [];
    var parentChanged = false;
    for (var j = 1; j < treeValues.length; j++) {
      var folderId = String(treeValues[j][folderIdCol]);
      var parentValue = treeValues[j][parentCol];
      if (folderUpdateMap.hasOwnProperty(folderId)) {
        parentValue = folderUpdateMap[folderId];
        parentChanged = true;
      }
      parentColData.push([parentValue]);
    }
    if (parentChanged) {
      treeSheet
        .getRange(2, parentCol + 1, treeValues.length, parentCol + 1)
        .setValues(parentColData);
    }
  }
}
