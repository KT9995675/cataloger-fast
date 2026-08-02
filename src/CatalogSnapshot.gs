/**
 * §24.1 — экспорт снимка метаданных каталога в новую Google Таблицу.
 * Листы Tree/Files/ACL/Users/Groups/GroupMembers + Config; Jobs не включаем.
 */

/** @const {string[]} */
var SNAPSHOT_SHEET_NAMES_ = [
  'Tree',
  'Files',
  'ACL',
  'Users',
  'Groups',
  'GroupMembers'
];

/**
 * @param {{ targetFolderId?: string }=} input
 * @returns {{
 *   ok: true,
 *   spreadsheetId: string,
 *   url: string,
 *   title: string,
 *   sheetNames: string[]
 * }}
 */
function exportCatalogSnapshot(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  input = input || {};

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var targetFolderId = String(input.targetFolderId || 'root').trim() || 'root';
  if (targetFolderId !== 'root') {
    assertDriveFolderAccessible_(targetFolderId);
  }

  var tz = Session.getScriptTimeZone() || 'Europe/Moscow';
  var title =
    'Каталог — снимок ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  var ss = SpreadsheetApp.create(title);
  var spreadsheetId = ss.getId();
  moveDriveFileToFolder_(spreadsheetId, targetFolderId);

  var created = [];
  SNAPSHOT_SHEET_NAMES_.forEach(function (name) {
    copyCatalogSheetIntoSpreadsheet_(ss, name);
    created.push(name);
  });
  writeCatalogSnapshotConfigSheet_(ss, userEmail);
  created.push('Config');

  removeDefaultEmptySheet_(ss);

  return {
    ok: true,
    spreadsheetId: spreadsheetId,
    url: ss.getUrl(),
    title: title,
    sheetNames: created
  };
}

/**
 * @param {string} folderId
 */
function assertDriveFolderAccessible_(folderId) {
  var token = ScriptApp.getOAuthToken();
  var meta = driveImportFetchJson_(
    'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(folderId) +
      '?fields=' +
      encodeURIComponent('id,name,mimeType') +
      '&supportsAllDrives=true',
    token
  );
  if (!meta || !meta.id) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Папка Drive не найдена.');
  }
  if (String(meta.mimeType || '') !== 'application/vnd.google-apps.folder') {
    throw catalogError_('INVALID_INPUT', 'Цель экспорта должна быть папкой Drive.');
  }
}

/**
 * @param {string} fileId
 * @param {string} targetFolderId 'root' или id папки
 */
function moveDriveFileToFolder_(fileId, targetFolderId) {
  var file = DriveApp.getFileById(fileId);
  var dest =
    !targetFolderId || targetFolderId === 'root'
      ? DriveApp.getRootFolder()
      : DriveApp.getFolderById(targetFolderId);

  dest.addFile(file);
  var parents = file.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    if (parent.getId() !== dest.getId()) {
      try {
        parent.removeFile(file);
      } catch (eRem) {
        // ignore
      }
    }
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 */
function copyCatalogSheetIntoSpreadsheet_(ss, sheetName) {
  var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!src) {
    throw catalogError_('SCHEMA_MISMATCH', 'Нет листа: ' + sheetName);
  }
  var dst = ss.insertSheet(sheetName);
  var range = src.getDataRange();
  var values = range.getValues();
  if (!values.length) {
    var schema = getCatalogSheetSchema_()[sheetName];
    if (schema && schema.length) {
      dst.getRange(1, 1, 1, schema.length).setValues([schema]);
      dst.setFrozenRows(1);
    }
    return;
  }
  dst.getRange(1, 1, values.length, values[0].length).setValues(values);
  if (src.getFrozenRows() > 0) {
    dst.setFrozenRows(src.getFrozenRows());
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} exportedBy
 */
function writeCatalogSnapshotConfigSheet_(ss, exportedBy) {
  var props = PropertiesService.getDocumentProperties();
  var rows = [
    ['key', 'value'],
    ['SCHEMA_VERSION', props.getProperty(PROP_SCHEMA_VERSION_) || ''],
    ['CATALOG_ROOT_FOLDER_ID', props.getProperty(PROP_CATALOG_ROOT_FOLDER_ID_) || ''],
    [
      'CATALOG_VIRTUAL_ROOT_FOLDER_ID',
      props.getProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_) || ''
    ],
    ['CONTROLLER_EMAIL', props.getProperty(PROP_CONTROLLER_EMAIL_) || ''],
    ['SETUP_AT', props.getProperty(PROP_SETUP_AT_) || ''],
    ['CATALOG_REV', props.getProperty(PROP_CATALOG_REV_) || ''],
    ['EXPORTED_AT', new Date().toISOString()],
    ['EXPORTED_BY', String(exportedBy || '')],
    [
      'SOURCE_SPREADSHEET_ID',
      SpreadsheetApp.getActiveSpreadsheet().getId()
    ],
    ['SNAPSHOT_KIND', 'catalog_metadata_v1']
  ];
  var sheet = ss.insertSheet('Config');
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 360);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function removeDefaultEmptySheet_(ss) {
  var sheets = ss.getSheets();
  if (sheets.length < 2) {
    return;
  }
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var name = sh.getName();
    if (name === 'Sheet1' || name === 'Лист1') {
      try {
        ss.deleteSheet(sh);
      } catch (eDel) {
        // ignore
      }
      return;
    }
  }
}

/**
 * §24.2 — импорт снимка: очистка метаданных + заливка листов + сверка Drive.
 *
 * @param {{ spreadsheetId: string }} input
 * @returns {{
 *   ok: true,
 *   title: string,
 *   missingCount: number,
 *   missingSample: string[],
 *   extrasToTrash: number,
 *   extrasCopied: number,
 *   controllerMismatch: (string|null),
 *   message: string
 * }}
 */
function importCatalogSnapshot(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  input = input || {};

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var spreadsheetId = String(input.spreadsheetId || '').trim();
  if (!spreadsheetId) {
    throw catalogError_('INVALID_INPUT', 'Укажите таблицу-снимок.');
  }

  var snapSs;
  try {
    snapSs = SpreadsheetApp.openById(spreadsheetId);
  } catch (eOpen) {
    throw catalogError_(
      'INVALID_INPUT',
      'Не удалось открыть снимок. Проверьте доступ к таблице.'
    );
  }

  var config = readCatalogSnapshotConfig_(snapSs);
  if (String(config.SNAPSHOT_KIND || '') !== 'catalog_metadata_v1') {
    throw catalogError_(
      'INVALID_INPUT',
      'Это не снимок каталога (нет Config / SNAPSHOT_KIND).'
    );
  }

  SNAPSHOT_SHEET_NAMES_.forEach(function (name) {
    if (!snapSs.getSheetByName(name)) {
      throw catalogError_('INVALID_INPUT', 'В снимке нет листа: ' + name);
    }
  });

  var props = PropertiesService.getDocumentProperties();
  var currentRootId = props.getProperty(PROP_CATALOG_ROOT_FOLDER_ID_) || '';
  var currentController = props.getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var snapController = String(config.CONTROLLER_EMAIL || '').trim();
  var controllerMismatch = null;
  if (
    snapController &&
    currentController &&
    snapController.toLowerCase() !== currentController.toLowerCase()
  ) {
    controllerMismatch =
      'В снимке Управляющий «' +
      snapController +
      '», в каталоге сейчас «' +
      currentController +
      '». Роль входа остаётся за владельцем таблицы.';
  }

  // Предварительная очистка метаданных (файлы на Drive не трогаем).
  SNAPSHOT_SHEET_NAMES_.forEach(function (name) {
    clearCatalogSheetDataRows_(name);
  });
  clearCatalogSheetDataRows_('Jobs');

  SNAPSHOT_SHEET_NAMES_.forEach(function (name) {
    replaceCatalogSheetFromSnapshot_(snapSs, name);
  });
  ensureCatalogSchemaUpToDate_();

  if (config.CATALOG_VIRTUAL_ROOT_FOLDER_ID) {
    props.setProperty(
      PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_,
      String(config.CATALOG_VIRTUAL_ROOT_FOLDER_ID)
    );
  }
  if (config.SCHEMA_VERSION) {
    props.setProperty(PROP_SCHEMA_VERSION_, String(config.SCHEMA_VERSION));
  }
  // Плоская папка Drive — текущая (файлы на месте).
  if (currentRootId) {
    props.setProperty(PROP_CATALOG_ROOT_FOLDER_ID_, currentRootId);
  }
  if (currentController) {
    props.setProperty(PROP_CONTROLLER_EMAIL_, currentController);
    writeControllerUser_(currentController, new Date());
  }

  ensureTrashFolderRowAfterSnapshotImport_();
  var reconcile = reconcileCatalogDriveAfterSnapshotImport_(currentController);
  bumpCatalogRev_();

  var parts = [];
  if (reconcile.missingCount) {
    parts.push('не найдено на Drive: ' + reconcile.missingCount);
  }
  if (reconcile.extrasToTrash) {
    parts.push('лишних в корзину каталога: ' + reconcile.extrasToTrash);
  }
  if (reconcile.extrasCopied) {
    parts.push('скопировано (чужой владелец): ' + reconcile.extrasCopied);
  }
  if (controllerMismatch) {
    parts.push('расхождение Управляющего');
  }

  return {
    ok: true,
    title: snapSs.getName(),
    missingCount: reconcile.missingCount,
    missingSample: reconcile.missingSample,
    extrasToTrash: reconcile.extrasToTrash,
    extrasCopied: reconcile.extrasCopied,
    controllerMismatch: controllerMismatch,
    message: parts.length ? parts.join('; ') : 'Снимок загружен, сверка OK'
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} snapSs
 * @returns {Object.<string, string>}
 */
function readCatalogSnapshotConfig_(snapSs) {
  var sheet = snapSs.getSheetByName('Config');
  var map = {};
  if (!sheet) {
    return map;
  }
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) {
      continue;
    }
    map[key] = String(values[i][1] != null ? values[i][1] : '');
  }
  return map;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} snapSs
 * @param {string} sheetName
 */
function replaceCatalogSheetFromSnapshot_(snapSs, sheetName) {
  var src = snapSs.getSheetByName(sheetName);
  var dst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!src || !dst) {
    throw catalogError_('SCHEMA_MISMATCH', 'Лист недоступен: ' + sheetName);
  }
  var values = src.getDataRange().getValues();
  var schema = getCatalogSheetSchema_()[sheetName] || [];
  dst.clearContents();
  if (!values.length) {
    if (schema.length) {
      dst.getRange(1, 1, 1, schema.length).setValues([schema]);
      dst.setFrozenRows(1);
    }
    return;
  }
  dst.getRange(1, 1, values.length, values[0].length).setValues(values);
  dst.setFrozenRows(1);
}

/**
 * Если в снимке нет ## Корзина — создать строку Tree.
 */
function ensureTrashFolderRowAfterSnapshotImport_() {
  var rows = readSheetRecords_('Tree');
  var hasTrash = false;
  var rootId = '';
  rows.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (id === '__TRASH__') {
      hasTrash = true;
    }
    var parent = String(row.parent_folder_id || '').trim();
    if (!parent && id && id !== '__TRASH__') {
      rootId = id;
    }
  });
  if (hasTrash) {
    return;
  }
  if (!rootId) {
    rootId =
      PropertiesService.getDocumentProperties().getProperty(
        PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_
      ) || '';
  }
  if (!rootId) {
    return;
  }
  appendTreeFolderRowsBatch_([
    {
      folderId: '__TRASH__',
      parentFolderId: rootId,
      name: '## Корзина',
      folderCreatedAt: new Date(),
      isSystem: true
    }
  ]);
}

/**
 * Сверка Files ↔ плоская папка Drive после импорта снимка.
 *
 * @param {string} controllerEmail
 * @returns {{
 *   missingCount: number,
 *   missingSample: string[],
 *   extrasToTrash: number,
 *   extrasCopied: number
 * }}
 */
function reconcileCatalogDriveAfterSnapshotImport_(controllerEmail) {
  var fileRows = readSheetRecords_('Files');
  var inCatalog = {};
  var missingSample = [];
  var missingCount = 0;

  fileRows.forEach(function (row) {
    var fileId = String(row.file_id || '').trim();
    if (!fileId) {
      return;
    }
    inCatalog[fileId] = true;
    try {
      DriveApp.getFileById(fileId);
    } catch (eMiss) {
      missingCount += 1;
      if (missingSample.length < 15) {
        missingSample.push(String(row.display_name || fileId));
      }
    }
  });

  var extrasToTrash = 0;
  var extrasCopied = 0;
  var rootId = '';
  try {
    rootId = getCatalogRootFolderId_();
  } catch (eRoot) {
    return {
      missingCount: missingCount,
      missingSample: missingSample,
      extrasToTrash: 0,
      extrasCopied: 0
    };
  }

  var folder = DriveApp.getFolderById(rootId);
  var it = folder.getFiles();
  var guard = 0;
  var trashRows = [];
  var controllerLc = String(controllerEmail || '').toLowerCase();

  while (it.hasNext() && guard < 20000) {
    guard += 1;
    var file = it.next();
    var id = file.getId();
    if (inCatalog[id]) {
      continue;
    }

    var useFile = file;
    try {
      var ownerEmail = '';
      try {
        var owner = file.getOwner();
        ownerEmail = owner ? String(owner.getEmail() || '').toLowerCase() : '';
      } catch (eOwn) {
        ownerEmail = '';
      }
      if (ownerEmail && controllerLc && ownerEmail !== controllerLc) {
        useFile = file.makeCopy(file.getName(), folder);
        try {
          moveDriveFileToTrash_(id);
        } catch (eTrashOrig) {
          // leave original if trash failed
        }
        extrasCopied += 1;
      }
    } catch (eCopy) {
      // if copy failed, try to register original in trash
      useFile = file;
    }

      var mimeType = getDriveFileMimeType_(useFile) || '';
      trashRows.push({
        catalogId: Utilities.getUuid(),
        folderId: '__TRASH__',
        fileId: useFile.getId(),
        displayName: useFile.getName(),
        sizeBytes: resolveDriveFileSizeBytes_(useFile, mimeType),
        driveModifiedAt: useFile.getLastUpdated(),
        sourceFileId: '',
        mimeType: mimeType,
        status: 'ready'
      });
    extrasToTrash += 1;
    inCatalog[useFile.getId()] = true;
  }

  if (trashRows.length) {
    appendCatalogFileRowsBatch_(trashRows);
  }

  return {
    missingCount: missingCount,
    missingSample: missingSample,
    extrasToTrash: extrasToTrash,
    extrasCopied: extrasCopied
  };
}

/**
 * §24.2a — дополнить каталог из снимка без дублей и без предварительной очистки.
 *
 * @param {{ spreadsheetId: string }} input
 * @returns {{
 *   ok: true,
 *   title: string,
 *   added: Object.<string, number>,
 *   skipped: Object.<string, number>,
 *   missingCount: number,
 *   missingSample: string[],
 *   controllerMismatch: (string|null),
 *   message: string
 * }}
 */
function mergeCatalogSnapshot(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  input = input || {};

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);
  assertNoActiveCatalogJobs_();

  var spreadsheetId = String(input.spreadsheetId || '').trim();
  if (!spreadsheetId) {
    throw catalogError_('INVALID_INPUT', 'Укажите таблицу-снимок.');
  }

  var snapSs;
  try {
    snapSs = SpreadsheetApp.openById(spreadsheetId);
  } catch (eOpen) {
    throw catalogError_(
      'INVALID_INPUT',
      'Не удалось открыть снимок. Проверьте доступ к таблице.'
    );
  }

  var config = readCatalogSnapshotConfig_(snapSs);
  if (String(config.SNAPSHOT_KIND || '') !== 'catalog_metadata_v1') {
    throw catalogError_(
      'INVALID_INPUT',
      'Это не снимок каталога (нет Config / SNAPSHOT_KIND).'
    );
  }

  SNAPSHOT_SHEET_NAMES_.forEach(function (name) {
    if (!snapSs.getSheetByName(name)) {
      throw catalogError_('INVALID_INPUT', 'В снимке нет листа: ' + name);
    }
  });

  var props = PropertiesService.getDocumentProperties();
  var currentController = props.getProperty(PROP_CONTROLLER_EMAIL_) || '';
  var snapController = String(config.CONTROLLER_EMAIL || '').trim();
  var controllerMismatch = null;
  if (
    snapController &&
    currentController &&
    snapController.toLowerCase() !== currentController.toLowerCase()
  ) {
    controllerMismatch =
      'В снимке Управляющий «' +
      snapController +
      '», в каталоге сейчас «' +
      currentController +
      '». Роль входа не меняется.';
  }

  var added = {
    users: 0,
    groups: 0,
    groupMembers: 0,
    tree: 0,
    files: 0,
    acl: 0
  };
  var skipped = {
    users: 0,
    groups: 0,
    groupMembers: 0,
    tree: 0,
    files: 0,
    acl: 0,
    orphans: 0
  };

  var controllerLc = String(currentController || '').toLowerCase();

  // Users
  var existingUsers = {};
  readSheetRecords_('Users').forEach(function (row) {
    var em = String(row.email || '')
      .trim()
      .toLowerCase();
    if (em) {
      existingUsers[em] = true;
    }
  });
  var snapUsers = readSnapshotSheetRecords_(snapSs, 'Users');
  var usersToAdd = [];
  snapUsers.forEach(function (row) {
    var em = String(row.email || '')
      .trim()
      .toLowerCase();
    if (!em) {
      return;
    }
    if (existingUsers[em]) {
      skipped.users += 1;
      return;
    }
    var role = String(row.login_role || 'user')
      .trim()
      .toLowerCase();
    if (role === 'controller' && em !== controllerLc) {
      row.login_role = 'manager';
    }
    usersToAdd.push(row);
    existingUsers[em] = true;
  });
  added.users = appendRawRecordsToCatalogSheet_('Users', usersToAdd);

  // Groups
  var existingGroups = {};
  readSheetRecords_('Groups').forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    if (gid) {
      existingGroups[gid] = true;
    }
  });
  var snapGroups = readSnapshotSheetRecords_(snapSs, 'Groups');
  var groupsToAdd = [];
  snapGroups.forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    if (!gid) {
      return;
    }
    if (existingGroups[gid]) {
      skipped.groups += 1;
      return;
    }
    groupsToAdd.push(row);
    existingGroups[gid] = true;
  });
  added.groups = appendRawRecordsToCatalogSheet_('Groups', groupsToAdd);

  // GroupMembers
  var existingMembers = {};
  readSheetRecords_('GroupMembers').forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    var em = String(row.email || '')
      .trim()
      .toLowerCase();
    if (gid && em) {
      existingMembers[gid + '\t' + em] = true;
    }
  });
  var snapMembers = readSnapshotSheetRecords_(snapSs, 'GroupMembers');
  var membersToAdd = [];
  snapMembers.forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    var em = String(row.email || '')
      .trim()
      .toLowerCase();
    if (!gid || !em) {
      return;
    }
    var key = gid + '\t' + em;
    if (existingMembers[key]) {
      skipped.groupMembers += 1;
      return;
    }
    if (!existingGroups[gid] || !existingUsers[em]) {
      skipped.orphans += 1;
      return;
    }
    membersToAdd.push(row);
    existingMembers[key] = true;
  });
  added.groupMembers = appendRawRecordsToCatalogSheet_('GroupMembers', membersToAdd);

  // Tree
  var existingFolders = {};
  readSheetRecords_('Tree').forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (id) {
      existingFolders[id] = true;
    }
  });
  var snapTree = readSnapshotSheetRecords_(snapSs, 'Tree');
  var snapFolderIds = {};
  snapTree.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (id) {
      snapFolderIds[id] = true;
    }
  });
  var treeToAdd = [];
  snapTree.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (!id) {
      return;
    }
    if (existingFolders[id]) {
      skipped.tree += 1;
      return;
    }
    var parent = String(row.parent_folder_id || '').trim();
    if (parent && !existingFolders[parent] && !snapFolderIds[parent]) {
      skipped.orphans += 1;
      return;
    }
    treeToAdd.push(row);
    existingFolders[id] = true;
  });
  added.tree = appendRawRecordsToCatalogSheet_('Tree', treeToAdd);

  // Files (key = file_id; also skip duplicate catalog_id)
  var existingFileIds = {};
  var existingCatalogIds = {};
  readSheetRecords_('Files').forEach(function (row) {
    var fid = String(row.file_id || '').trim();
    var cid = String(row.catalog_id || '').trim();
    if (fid) {
      existingFileIds[fid] = true;
    }
    if (cid) {
      existingCatalogIds[cid] = true;
    }
  });
  var snapFiles = readSnapshotSheetRecords_(snapSs, 'Files');
  var filesToAdd = [];
  var newlyAddedFileIds = [];
  snapFiles.forEach(function (row) {
    var fid = String(row.file_id || '').trim();
    var cid = String(row.catalog_id || '').trim();
    var folderId = String(row.folder_id || '').trim();
    if (!fid && !cid) {
      return;
    }
    if ((fid && existingFileIds[fid]) || (cid && existingCatalogIds[cid])) {
      skipped.files += 1;
      return;
    }
    if (!folderId || !existingFolders[folderId]) {
      skipped.orphans += 1;
      return;
    }
    filesToAdd.push(row);
    if (fid) {
      existingFileIds[fid] = true;
      newlyAddedFileIds.push(fid);
    }
    if (cid) {
      existingCatalogIds[cid] = true;
    }
  });
  added.files = appendRawRecordsToCatalogSheet_('Files', filesToAdd);

  // ACL
  var existingAclIds = {};
  var existingAclComposite = {};
  readSheetRecords_('ACL').forEach(function (row) {
    var aid = String(row.acl_id || '').trim();
    if (aid) {
      existingAclIds[aid] = true;
    }
    existingAclComposite[aclCompositeKey_(row)] = true;
  });
  var snapAcl = readSnapshotSheetRecords_(snapSs, 'ACL');
  var aclToAdd = [];
  snapAcl.forEach(function (row) {
    var aid = String(row.acl_id || '').trim();
    var objectType = String(row.object_type || '')
      .trim()
      .toLowerCase();
    var objectId = String(row.object_id || '').trim();
    if (!objectId) {
      return;
    }
    if (aid && existingAclIds[aid]) {
      skipped.acl += 1;
      return;
    }
    var comp = aclCompositeKey_(row);
    if (existingAclComposite[comp]) {
      skipped.acl += 1;
      return;
    }
    var objectOk = false;
    if (objectType === 'folder') {
      objectOk = !!existingFolders[objectId];
    } else if (objectType === 'file') {
      objectOk = !!existingCatalogIds[objectId];
    }
    if (!objectOk) {
      skipped.orphans += 1;
      return;
    }
    aclToAdd.push(row);
    if (aid) {
      existingAclIds[aid] = true;
    }
    existingAclComposite[comp] = true;
  });
  added.acl = appendRawRecordsToCatalogSheet_('ACL', aclToAdd);

  ensureTrashFolderRowAfterSnapshotImport_();

  var missingSample = [];
  var missingCount = 0;
  newlyAddedFileIds.forEach(function (fileId) {
    try {
      DriveApp.getFileById(fileId);
    } catch (eMiss) {
      missingCount += 1;
      if (missingSample.length < 15) {
        var name = fileId;
        for (var i = 0; i < filesToAdd.length; i++) {
          if (String(filesToAdd[i].file_id || '').trim() === fileId) {
            name = String(filesToAdd[i].display_name || fileId);
            break;
          }
        }
        missingSample.push(name);
      }
    }
  });

  bumpCatalogRev_();

  var parts = [];
  parts.push(
    'добавлено: Tree ' +
      added.tree +
      ', Files ' +
      added.files +
      ', Users ' +
      added.users +
      ', Groups ' +
      added.groups +
      ', ACL ' +
      added.acl
  );
  if (skipped.orphans) {
    parts.push('сирот пропущено: ' + skipped.orphans);
  }
  if (missingCount) {
    parts.push('новых файлов нет на Drive: ' + missingCount);
  }
  if (controllerMismatch) {
    parts.push('расхождение Управляющего');
  }

  return {
    ok: true,
    title: snapSs.getName(),
    added: added,
    skipped: skipped,
    missingCount: missingCount,
    missingSample: missingSample,
    controllerMismatch: controllerMismatch,
    message: parts.join('; ')
  };
}

/**
 * @param {Object} row
 * @returns {string}
 */
function aclCompositeKey_(row) {
  return [
    String(row.object_type || '')
      .trim()
      .toLowerCase(),
    String(row.object_id || '').trim(),
    String(row.principal_type || '')
      .trim()
      .toLowerCase(),
    String(row.principal_id || '')
      .trim()
      .toLowerCase(),
    String(row.permission_level || '')
      .trim()
      .toLowerCase(),
    String(row.delta || '').trim()
  ].join('\t');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} snapSs
 * @param {string} sheetName
 * @returns {Object[]}
 */
function readSnapshotSheetRecords_(snapSs, sheetName) {
  var sheet = snapSs.getSheetByName(sheetName);
  if (!sheet) {
    return [];
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var record = {};
    var empty = true;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) {
        continue;
      }
      record[headers[c]] = values[i][c];
      if (values[i][c] !== '' && values[i][c] != null) {
        empty = false;
      }
    }
    if (!empty) {
      rows.push(record);
    }
  }
  return rows;
}

/**
 * @param {string} sheetName
 * @param {Object[]} records
 * @returns {number}
 */
function appendRawRecordsToCatalogSheet_(sheetName, records) {
  if (!records || !records.length) {
    return 0;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: ' + sheetName);
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  while (headers.length && !headers[headers.length - 1]) {
    headers.pop();
  }
  if (!headers.length) {
    var schema = getCatalogSheetSchema_()[sheetName] || [];
    headers = schema.slice();
    if (!headers.length) {
      throw catalogError_('SCHEMA_MISMATCH', sheetName + ' sheet has no headers.');
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  var matrix = [];
  for (var r = 0; r < records.length; r++) {
    var rec = records[r];
    var line = [];
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      line.push(rec[key] != null ? rec[key] : '');
    }
    matrix.push(line);
  }
  var startRow = Math.max(sheet.getLastRow(), 1) + 1;
  if (sheet.getLastRow() < 1) {
    startRow = 2;
  }
  sheet
    .getRange(startRow, 1, startRow + matrix.length - 1, headers.length)
    .setValues(matrix);
  return records.length;
}
