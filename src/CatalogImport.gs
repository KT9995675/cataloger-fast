/** Sync import safety limit — OLD; импорт всегда через Jobs (§9.6). */
var IMPORT_DRIVE_MAX_FILES_ = 30;
/* <!-- OLD: жёсткий отказ при >30; теперь enqueue без sync-лимита, потолок IMPORT_DRIVE_JOB_MAX_FILES_ --> */

/**
 * §9.10–§9.12 — импорт с Drive.
 * Сценарии: `drive` (файлы+папки, поиск/ссылки), legacy `files` / `folder`.
 * Всегда Jobs (фон); при активной очереди — JOBS_BUSY.
 *
 * @param {{
 *   targetFolderId: string,
 *   scenario?: 'drive'|'files'|'folder',
 *   driveUrl?: string,
 *   driveUrls?: string[],
 *   driveFileUrl?: string,
 *   driveFolderUrl?: string,
 *   items?: Array<{ kind?: 'file'|'folder', id?: string, url?: string }>,
 *   mode?: 'copy'|'move'
 * }} input
 * @returns {Object}
 */
function importFromDrive(input) {
  input = input || {};
  var scenario = String(input.scenario || 'drive').trim().toLowerCase();
  if (scenario === 'drive' || scenario === 'selection') {
    return importDriveSelection(input);
  }
  if (scenario === 'folder') {
    return importDriveFolder({
      targetFolderId: input.targetFolderId,
      driveFolderUrl: input.driveFolderUrl || input.driveUrl || input.driveFileUrl,
      mode: input.mode
    });
  }
  if (scenario === 'files') {
    var urls = normalizeImportDriveUrls_(input);
    return importDriveFiles({
      targetFolderId: input.targetFolderId,
      driveUrls: urls,
      mode: input.mode
    });
  }
  throw catalogError_('INVALID_INPUT', 'scenario must be drive, files or folder.');
}

/**
 * @param {Object} input
 * @returns {string[]}
 */
function normalizeImportDriveUrls_(input) {
  var urls = [];
  if (Array.isArray(input.driveUrls)) {
    input.driveUrls.forEach(function (u) {
      var s = String(u || '').trim();
      if (s) {
        urls.push(s);
      }
    });
  }
  if (!urls.length) {
    var single = String(input.driveUrl || input.driveFileUrl || '').trim();
    if (single) {
      single.split(/\r?\n/).forEach(function (line) {
        var s = String(line || '').trim();
        if (s) {
          urls.push(s);
        }
      });
    }
  }
  if (!urls.length) {
    throw catalogError_('INVALID_INPUT', 'Укажите хотя бы одну ссылку на файл.');
  }
  return urls;
}

/**
 * §9.1 / §9.12 — импорт одного или нескольких файлов → Jobs (фаза права → copy).
 *
 * @param {{
 *   targetFolderId: string,
 *   driveUrls: string[],
 *   mode?: 'copy'|'move'
 * }} input
 * @returns {{
 *   ok: true,
 *   queued: true,
 *   kind: 'files',
 *   jobId: string,
 *   fileCount: number,
 *   mode: 'copy'|'move'
 * }}
 */
function importDriveFiles(input) {
  assertCatalogReady_();
  assertNoActiveCatalogJobs_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var mode = String(input.mode || 'copy').trim().toLowerCase();
  var driveUrls = Array.isArray(input.driveUrls) ? input.driveUrls : [];

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (mode !== 'copy' && mode !== 'move') {
    throw catalogError_('INVALID_INPUT', 'mode must be copy or move.');
  }

  var urls = [];
  driveUrls.forEach(function (u) {
    var s = String(u || '').trim();
    if (s) {
      urls.push(s);
    }
  });
  if (!urls.length) {
    throw catalogError_('INVALID_INPUT', 'Укажите хотя бы одну ссылку на файл.');
  }
  if (urls.length > IMPORT_DRIVE_OPERATION_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов (' +
        urls.length +
        '). Максимум за одну операцию: ' +
        IMPORT_DRIVE_OPERATION_MAX_FILES_ +
        '.'
    );
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
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, targetFolderId);

  for (var k = 0; k < urls.length; k++) {
    if (resolveDriveImportKind_(urls[k]) === 'folder') {
      throw catalogError_(
        'INVALID_INPUT',
        'В сценарии «Файлы» нужна ссылка на файл, не на папку. Для папки выберите сценарий «Папка».'
      );
    }
  }

  var items = [];
  var fileRows = [];
  for (var i = 0; i < urls.length; i++) {
    var sourceFileId = parseDriveFileId_(urls[i]);
    var sourceFile;
    try {
      sourceFile = DriveApp.getFileById(sourceFileId);
    } catch (e) {
      throw catalogError_('INVALID_FILE', 'Drive file not found or not accessible: ' + urls[i]);
    }
    var built = buildPendingImportFileItem_(sourceFileId, targetFolderId, sourceFile.getName());
    items.push(built.item);
    fileRows.push(built.row);
  }
  var chainIdFiles = Utilities.getUuid();
  fileRows.forEach(function (row) {
    row.importChainId = chainIdFiles;
  });
  appendCatalogFileRowsBatch_(fileRows);

  var queuedFiles = enqueueImportDriveJobsChain_(
    mode,
    targetFolderId,
    items,
    userEmail,
    'files',
    chainIdFiles
  );
  ensureCatalogJobsTrigger_();
  kickCatalogJobsProcessing_();
  bumpCatalogRev_();

  return {
    ok: true,
    queued: true,
    kind: 'files',
    jobId: queuedFiles.jobId,
    chainId: queuedFiles.chainId,
    fileCount: queuedFiles.fileCount,
    jobParts: queuedFiles.jobParts,
    mode: mode,
    displayName:
      queuedFiles.fileCount === 1
        ? items[0].displayName
        : queuedFiles.fileCount + ' файл(ов)'
  };
}

/**
 * §9.1, §9.2 — импорт одного файла (обёртка над importDriveFiles).
 *
 * @param {{
 *   targetFolderId: string,
 *   driveFileUrl: string,
 *   mode?: 'copy'|'move'
 * }} input
 * @returns {Object}
 */
function importDriveFile(input) {
  input = input || {};
  return importDriveFiles({
    targetFolderId: input.targetFolderId,
    driveUrls: [String(input.driveFileUrl || input.driveUrl || '').trim()],
    mode: input.mode
  });
}

/**
 * §9.8, §9.11 — импорт папки Drive → Tree сразу, файлы в Jobs.
 *
 * @param {{
 *   targetFolderId: string,
 *   driveFolderUrl: string,
 *   mode?: 'copy'|'move'
 * }} input
 * @returns {{
 *   ok: true,
 *   queued: true,
 *   kind: 'folder',
 *   jobId: string,
 *   rootFolderId: string,
 *   displayName: string,
 *   folderCount: number,
 *   fileCount: number,
 *   skippedShortcuts: number,
 *   mode: 'copy'|'move'
 * }}
 */
function importDriveFolder(input) {
  assertCatalogReady_();
  assertNoActiveCatalogJobs_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var driveFolderUrl = String(input.driveFolderUrl || input.driveUrl || '').trim();
  var mode = String(input.mode || 'copy').trim().toLowerCase();

  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (!driveFolderUrl) {
    throw catalogError_('INVALID_INPUT', 'driveFolderUrl is required.');
  }
  if (mode !== 'copy' && mode !== 'move') {
    throw catalogError_('INVALID_INPUT', 'mode must be copy or move.');
  }
  if (resolveDriveImportKind_(driveFolderUrl) === 'file') {
    throw catalogError_(
      'INVALID_INPUT',
      'В сценарии «Папка» нужна ссылка на папку. Для файлов выберите сценарий «Файлы».'
    );
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
  assertEditorOnFolderForMove_(engine, userEmail, loginRole, targetFolderId);

  var sourceFolderId = parseDriveFolderId_(driveFolderUrl);
  var sourceFolder;
  try {
    sourceFolder = DriveApp.getFolderById(sourceFolderId);
  } catch (e) {
    throw catalogError_('INVALID_FOLDER', 'Drive folder not found or not accessible.');
  }

  var walk = walkDriveFolderForImport_(sourceFolder, IMPORT_DRIVE_OPERATION_MAX_FILES_);
  if (walk.fileCount > IMPORT_DRIVE_OPERATION_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов в папке (' +
        walk.fileCount +
        '). Максимум за одну операцию: ' +
        IMPORT_DRIVE_OPERATION_MAX_FILES_ +
        '.'
    );
  }

  var treeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!treeSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  var now = new Date();
  var driveToVirtual = {};
  var folderCount = 0;
  var treeRows = [];

  walk.folders.forEach(function (node) {
    var virtualId = Utilities.getUuid();
    driveToVirtual[node.driveFolderId] = virtualId;
    var parentVirtualId =
      node.parentDriveFolderId === null
        ? targetFolderId
        : driveToVirtual[node.parentDriveFolderId];
    if (!parentVirtualId) {
      throw catalogError_('IMPORT_WALK_FAILED', 'Parent virtual folder missing during import.');
    }
    treeRows.push([virtualId, parentVirtualId, node.name, now, false, '', '', '']);
    folderCount++;
  });
  if (treeRows.length) {
    treeSheet.getRange(treeSheet.getLastRow() + 1, 1, treeRows.length, 8).setValues(treeRows);
  }

  var rootVirtualId = driveToVirtual[sourceFolderId];
  applyDriveFolderAclToCatalogFolder_(sourceFolder, rootVirtualId, userEmail);

  var aclEngine = createAclEngine_();
  walk.folders.forEach(function (node) {
    if (node.parentDriveFolderId === null) {
      return;
    }
    var virtualId = driveToVirtual[node.driveFolderId];
    var parentVirtualId = driveToVirtual[node.parentDriveFolderId];
    if (!virtualId || !parentVirtualId) {
      return;
    }
    copyExplicitAclFromParentFolder_(aclEngine, 'folder', virtualId, parentVirtualId);
  });

  var items = [];
  var fileRows = [];
  walk.files.forEach(function (item) {
    var virtualParentId = driveToVirtual[item.parentDriveFolderId];
    if (!virtualParentId) {
      throw catalogError_('IMPORT_WALK_FAILED', 'Virtual folder missing for file parent.');
    }
    var built = buildPendingImportFileItem_(
      item.file.getId(),
      virtualParentId,
      item.file.getName()
    );
    items.push(built.item);
    fileRows.push(built.row);
  });
  var chainIdFolder = Utilities.getUuid();
  fileRows.forEach(function (row) {
    row.importChainId = chainIdFolder;
  });
  appendCatalogFileRowsBatch_(fileRows);

  var queuedFolder = enqueueImportDriveJobsChain_(
    mode,
    targetFolderId,
    items,
    userEmail,
    'folder',
    chainIdFolder
  );
  ensureCatalogJobsTrigger_();
  kickCatalogJobsProcessing_();
  bumpCatalogRev_();

  return {
    ok: true,
    queued: true,
    kind: 'folder',
    jobId: queuedFolder.jobId,
    chainId: queuedFolder.chainId,
    rootFolderId: rootVirtualId,
    displayName: sourceFolder.getName(),
    folderCount: folderCount,
    fileCount: queuedFolder.fileCount,
    jobParts: queuedFolder.jobParts,
    skippedShortcuts: walk.skippedShortcuts,
    mode: mode
  };
}

/**
 * При mode=move: свой файл → move, чужой → copy (§9.12).
 *
 * @param {'copy'|'move'} requestedMode
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} controllerEmail
 * @returns {'copy'|'move'}
 */
function resolveDriveImportPlaceMode_(requestedMode, driveFile, controllerEmail) {
  if (requestedMode !== 'move') {
    return 'copy';
  }
  return isDriveFileOwnedByEmail_(driveFile, controllerEmail) ? 'move' : 'copy';
}

/**
 * @param {string} urlOrId
 * @returns {'file'|'folder'}
 */
function resolveDriveImportKind_(urlOrId) {
  var value = String(urlOrId || '').trim();
  if (/\/folders\//.test(value)) {
    return 'folder';
  }
  if (
    /\/file\/d\//.test(value) ||
    /\/(?:document|spreadsheets|presentation)\//.test(value)
  ) {
    return 'file';
  }

  var id = '';
  try {
    id = parseDriveFolderId_(value);
  } catch (e1) {
    try {
      id = parseDriveFileId_(value);
    } catch (e2) {
      throw catalogError_('INVALID_DRIVE_URL', 'Cannot parse Drive file or folder from URL.');
    }
  }

  try {
    DriveApp.getFolderById(id);
    return 'folder';
  } catch (eFolder) {
    try {
      DriveApp.getFileById(id);
      return 'file';
    } catch (eFile) {
      throw catalogError_('INVALID_DRIVE_URL', 'Drive item not found or not accessible.');
    }
  }
}

/**
 * Depth-first walk: folders (root first) + files. Shortcuts skipped.
 *
 * @param {GoogleAppsScript.Drive.Folder} rootFolder
 * @param {number} maxFiles
 * @returns {{
 *   folders: Array<{ driveFolderId: string, parentDriveFolderId: (string|null), name: string }>,
 *   files: Array<{ file: GoogleAppsScript.Drive.File, parentDriveFolderId: string }>,
 *   fileCount: number,
 *   skippedShortcuts: number
 * }}
 */
function walkDriveFolderForImport_(rootFolder, maxFiles) {
  var folders = [];
  var files = [];
  var skippedShortcuts = 0;
  var queue = [{ folder: rootFolder, parentDriveFolderId: null }];

  while (queue.length) {
    var current = queue.shift();
    var driveFolder = current.folder;
    var driveFolderId = driveFolder.getId();
    folders.push({
      driveFolderId: driveFolderId,
      parentDriveFolderId: current.parentDriveFolderId,
      name: driveFolder.getName()
    });

    var fileIt = driveFolder.getFiles();
    while (fileIt.hasNext()) {
      var file = fileIt.next();
      var mime = '';
      try {
        mime = String(file.getMimeType() || '');
      } catch (eMime) {
        mime = '';
      }
      if (mime === 'application/vnd.google-apps.shortcut') {
        skippedShortcuts++;
        continue;
      }
      files.push({
        file: file,
        parentDriveFolderId: driveFolderId
      });
      if (files.length > maxFiles) {
        return {
          folders: folders,
          files: files,
          fileCount: files.length,
          skippedShortcuts: skippedShortcuts
        };
      }
    }

    var folderIt = driveFolder.getFolders();
    while (folderIt.hasNext()) {
      queue.push({
        folder: folderIt.next(),
        parentDriveFolderId: driveFolderId
      });
    }
  }

  return {
    folders: folders,
    files: files,
    fileCount: files.length,
    skippedShortcuts: skippedShortcuts
  };
}

/**
 * @param {Array<{ file: GoogleAppsScript.Drive.File }>} fileItems
 * @param {string} controllerEmail
 */
/**
 * §9.8 — ACL с корня исходной папки Drive → явные права на корень импортированной ветки.
 *
 * @param {GoogleAppsScript.Drive.Folder} driveFolder
 * @param {string} catalogFolderId
 * @param {string} addedBy
 */
function applyDriveFolderAclToCatalogFolder_(driveFolder, catalogFolderId, addedBy) {
  var levelsByEmail = {};
  var namesByEmail = {};
  var emailByKey = {};

  function remember(users, level) {
    if (!users) {
      return;
    }
    var list = [];
    if (typeof users.hasNext === 'function') {
      while (users.hasNext()) {
        list.push(users.next());
      }
    } else if (Array.isArray(users)) {
      list = users;
    }
    for (var i = 0; i < list.length; i++) {
      var user = list[i];
      if (!user) {
        continue;
      }
      var email = '';
      try {
        email = String(user.getEmail() || '').trim();
      } catch (e) {
        continue;
      }
      if (!email) {
        continue;
      }
      var key = email.toLowerCase();
      emailByKey[key] = email;
      namesByEmail[key] = resolveDriveUserDisplayName_(user) || email;
      levelsByEmail[key] = maxPermissionLevel_(levelsByEmail[key], level);
    }
  }

  remember(driveFolder.getEditors(), 'editor');
  try {
    remember(driveFolder.getCommenters(), 'commenter');
  } catch (eComment) {
    // optional on some folder types
  }
  remember(driveFolder.getViewers(), 'reader');

  Object.keys(levelsByEmail).forEach(function (key) {
    var level = levelsByEmail[key];
    if (!level || level === 'none') {
      return;
    }
    var email = emailByKey[key] || key;
    appendOrEnsureUserRow_({
      email: email,
      loginRole: 'user',
      addedBy: addedBy || '',
      displayName: namesByEmail[key] || email
    });
    appendExplicitUserAclRow_('folder', catalogFolderId, email, level);
  });

  var folderEngine = createAclEngine_();
  var entries = buildAclEntriesFromObject_(folderEngine, 'folder', catalogFolderId);
  syncAclCacheForObjects_(
    [{ objectType: 'folder', objectId: catalogFolderId }],
    entries,
    folderEngine
  );
}

/**
 * @returns {string}
 */
function getCatalogRootFolderId_() {
  var id = PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_ROOT_FOLDER_ID_);
  if (!id) {
    throw catalogError_('CATALOG_NOT_CONFIGURED', 'CATALOG_ROOT_FOLDER_ID is missing.');
  }
  return id;
}

/**
 * @param {string} urlOrId
 * @returns {string}
 */
function parseDriveFileId_(urlOrId) {
  var value = String(urlOrId || '').trim();
  var fileMatch = value.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) {
    return fileMatch[1];
  }
  var gdocMatch = value.match(/\/(?:document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]+)/);
  if (gdocMatch) {
    return gdocMatch[1];
  }
  var idMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch) {
    return idMatch[1];
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) {
    return value;
  }
  throw catalogError_('INVALID_FILE_URL', 'Cannot parse Drive file id from URL.');
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} email
 * @returns {boolean}
 */
function isDriveFileOwnedByEmail_(driveFile, email) {
  if (!email) {
    return false;
  }
  try {
    var owner = driveFile.getOwner();
    if (!owner) {
      return false;
    }
    var ownerEmail = owner.getEmail();
    return !!ownerEmail && ownerEmail.toLowerCase() === email.toLowerCase();
  } catch (e) {
    return false;
  }
}

/**
 * @param {GoogleAppsScript.Drive.File} sourceFile
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @param {'copy'|'move'} mode
 * @returns {GoogleAppsScript.Drive.File}
 */
function placeFileInCatalogRoot_(sourceFile, catalogRootFolder, mode) {
  if (mode === 'move') {
    sourceFile.moveTo(catalogRootFolder);
    return sourceFile;
  }
  return sourceFile.makeCopy(sourceFile.getName(), catalogRootFolder);
}

/**
 * §9.4 — добавляет пользователей с доступом к исходному файлу в `Users`.
 *
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @param {string} addedBy
 */
function ensureUsersFromDriveFile_(driveFile, addedBy) {
  var participants = collectDriveFileParticipants_(driveFile);
  if (!participants.length) {
    return;
  }

  participants.forEach(function (p) {
    appendOrEnsureUserRow_({
      email: p.email,
      loginRole: 'user',
      addedBy: addedBy || '',
      displayName: p.displayName || p.email
    });
  });
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @returns {Array<{ email: string, displayName: string }>}
 */
function collectDriveFileParticipants_(driveFile) {
  var byEmail = {};

  function rememberUsers(users) {
    if (!users) {
      return;
    }

    if (typeof users.hasNext === 'function') {
      while (users.hasNext()) {
        rememberUser_(users.next());
      }
      return;
    }

    if (Array.isArray(users)) {
      for (var i = 0; i < users.length; i++) {
        rememberUser_(users[i]);
      }
    }
  }

  function rememberUser_(user) {
    if (!user) {
      return;
    }
    var email = '';
    try {
      email = String(user.getEmail() || '').trim();
    } catch (e) {
      return;
    }
    if (!email) {
      return;
    }
    var key = email.toLowerCase();
    var displayName = resolveDriveUserDisplayName_(user);
    if (!byEmail[key] || (displayName && displayName.toLowerCase() !== key)) {
      byEmail[key] = {
        email: email,
        displayName: displayName || email
      };
    }
  }

  rememberUsers(driveFile.getEditors());
  rememberUsers(driveFile.getViewers());
  try {
    rememberUsers(driveFile.getCommenters());
  } catch (e) {
    // commenters may be unavailable for some mime types
  }

  return Object.keys(byEmail).map(function (k) {
    return byEmail[k];
  });
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @returns {string[]}
 */
function collectDriveFileParticipantEmails_(driveFile) {
  return collectDriveFileParticipants_(driveFile).map(function (p) {
    return p.email;
  });
}


/**
 * @param {{
 *   catalogId: string,
 *   folderId: string,
 *   fileId: string,
 *   displayName: string,
 *   sizeBytes: number,
 *   driveModifiedAt: Date|string,
 *   sourceFileId: string,
 *   mimeType?: string,
 *   approved?: boolean,
 *   approvedBy?: string,
 *   approvedAt?: (Date|string),
 *   status?: string,
 *   lastError?: string
 * }} row
 */
function appendCatalogFileRow_(row) {
  appendCatalogFileRowsBatch_([row]);
}

/**
 * Пакетная запись строк Files (один setValues).
 *
 * @param {Array<Object>} rows
 */
function appendCatalogFileRowsBatch_(rows) {
  if (!rows || !rows.length) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }

  ensureCatalogSchemaUpToDate_();

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
    throw catalogError_('SCHEMA_MISMATCH', 'Files sheet has no headers.');
  }

  var lines = rows.map(function (row) {
    var byHeader = {
      catalog_id: row.catalogId,
      folder_id: row.folderId,
      file_id: row.fileId,
      display_name: row.displayName,
      size_bytes: row.sizeBytes,
      drive_modified_at: row.driveModifiedAt,
      approved: row.approved === true,
      approved_by: row.approvedBy || '',
      approved_at: row.approvedAt || '',
      status: row.status || 'ready',
      last_error: row.lastError || '',
      source_file_id: row.sourceFileId || '',
      mime_type: row.mimeType || '',
      acl_editors: row.aclEditors || '',
      acl_commenters: row.aclCommenters || '',
      acl_readers: row.aclReaders || '',
      import_chain_id: row.importChainId || '',
      shortcut_of_catalog_id: row.shortcutOfCatalogId || '',
      shortcut_of_drive_file_id: row.shortcutOfDriveFileId || ''
    };
    var line = [];
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      line.push(Object.prototype.hasOwnProperty.call(byHeader, key) ? byHeader[key] : '');
    }
    return line;
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, lines.length, headers.length).setValues(lines);
  SpreadsheetApp.flush();
}

/**
 * Пакетная запись строк Tree (один setValues).
 *
 * @param {Array<{
 *   folderId: string,
 *   parentFolderId: string,
 *   name: string,
 *   folderCreatedAt: *,
 *   isSystem?: boolean,
 *   aclEditors?: string,
 *   aclCommenters?: string,
 *   aclReaders?: string
 * }>} rows
 */
function appendTreeFolderRowsBatch_(rows) {
  if (!rows || !rows.length) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  ensureCatalogSchemaUpToDate_();

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
    throw catalogError_('SCHEMA_MISMATCH', 'Tree sheet has no headers.');
  }

  var lines = rows.map(function (row) {
    var byHeader = {
      folder_id: row.folderId,
      parent_folder_id: row.parentFolderId,
      name: row.name,
      folder_created_at: row.folderCreatedAt,
      is_system: row.isSystem === true,
      acl_editors: row.aclEditors || '',
      acl_commenters: row.aclCommenters || '',
      acl_readers: row.aclReaders || '',
      mirror_of_folder_id: row.mirrorOfFolderId || '',
      mirror_of_drive_folder_id: row.mirrorOfDriveFolderId || ''
    };
    var line = [];
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      line.push(Object.prototype.hasOwnProperty.call(byHeader, key) ? byHeader[key] : '');
    }
    return line;
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, lines.length, headers.length).setValues(lines);
  SpreadsheetApp.flush();
}

/**
 * @param {GoogleAppsScript.Drive.File} driveFile
 * @returns {string}
 */
function getDriveFileMimeType_(driveFile) {
  try {
    return String(driveFile.getMimeType() || '');
  } catch (e) {
    return '';
  }
}
