/**
 * §9.12 — поиск на Drive по имени + единый импорт файлов и папок.
 */

/** @const {number} */
var DRIVE_IMPORT_SEARCH_PAGE_SIZE_ = 20;

/**
 * Поиск файлов и папок Drive по подстроке имени.
 *
 * @param {{ query: string, pageToken?: string }} input
 * @returns {{
 *   ok: true,
 *   query: string,
 *   items: Array<{ id: string, name: string, kind: 'file'|'folder', mimeType: string }>,
 *   nextPageToken: (string|null)
 * }}
 */
function searchDriveForImport(input) {
  assertCatalogReady_();

  input = input || {};
  var query = String(input.query || '').trim();
  var pageToken = String(input.pageToken || '').trim();
  if (!query) {
    throw catalogError_('INVALID_INPUT', 'Введите текст для поиска.');
  }
  if (query.length > 200) {
    throw catalogError_('INVALID_INPUT', 'Слишком длинный поисковый запрос.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var escaped = query.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var q = "trashed = false and name contains '" + escaped + "'";
  var params = [
    'q=' + encodeURIComponent(q),
    'pageSize=' + DRIVE_IMPORT_SEARCH_PAGE_SIZE_,
    'fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType)'),
    'supportsAllDrives=true',
    'includeItemsFromAllDrives=true',
    'orderBy=' + encodeURIComponent('folder,name')
  ];
  if (pageToken) {
    params.push('pageToken=' + encodeURIComponent(pageToken));
  }

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?' + params.join('&'),
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );
  var code = response.getResponseCode();
  var bodyText = response.getContentText() || '{}';
  var body;
  try {
    body = JSON.parse(bodyText);
  } catch (e) {
    throw catalogError_('DRIVE_SEARCH_FAILED', 'Не удалось разобрать ответ Drive.');
  }
  if (code < 200 || code >= 300) {
    var errMsg =
      (body.error && body.error.message) || 'Drive search failed (' + code + ')';
    throw catalogError_('DRIVE_SEARCH_FAILED', errMsg);
  }

  var items = [];
  (body.files || []).forEach(function (f) {
    if (!f || !f.id) {
      return;
    }
    var mime = String(f.mimeType || '');
    if (mime === 'application/vnd.google-apps.shortcut') {
      return;
    }
    var kind = mime === 'application/vnd.google-apps.folder' ? 'folder' : 'file';
    items.push({
      id: String(f.id),
      name: String(f.name || f.id),
      kind: kind,
      mimeType: mime
    });
  });

  return {
    ok: true,
    query: query,
    items: items,
    nextPageToken: body.nextPageToken ? String(body.nextPageToken) : null
  };
}

/**
 * Подсчёт файлов и объёма до подтверждения импорта (обход Drive, без записи в каталог).
 *
 * @param {{
 *   items?: Array<{ kind?: 'file'|'folder', id?: string, url?: string }>,
 *   driveUrls?: string[]
 * }} input
 * @returns {{
 *   ok: true,
 *   fileCount: number,
 *   totalBytes: number,
 *   folderCount: number
 * }}
 */
function preflightDriveImport(input) {
  assertCatalogReady_();

  input = input || {};
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var selections = normalizeDriveImportSelections_(input);
  if (!selections.length) {
    throw catalogError_('INVALID_INPUT', 'Выберите файлы/папки или укажите ссылки.');
  }

  var fileCount = 0;
  var totalBytes = 0;
  var folderCount = 0;
  var seenSourceIds = {};

  selections.forEach(function (sel) {
    if (sel.kind === 'folder') {
      var sourceFolder;
      try {
        sourceFolder = DriveApp.getFolderById(sel.driveId);
      } catch (eFolder) {
        throw catalogError_('INVALID_FOLDER', 'Drive folder not found or not accessible.');
      }
      var walk = walkDriveFolderForImport_(sourceFolder, IMPORT_DRIVE_JOB_MAX_FILES_);
      if (walk.fileCount > IMPORT_DRIVE_JOB_MAX_FILES_) {
        throw catalogError_(
          'IMPORT_TOO_LARGE',
          'Слишком много файлов в папке «' +
            sourceFolder.getName() +
            '» (' +
            walk.fileCount +
            '). Максимум: ' +
            IMPORT_DRIVE_JOB_MAX_FILES_ +
            '.'
        );
      }
      folderCount += walk.folders.length;
      walk.files.forEach(function (item) {
        var sourceFileId = item.file.getId();
        if (seenSourceIds[sourceFileId]) {
          return;
        }
        seenSourceIds[sourceFileId] = true;
        fileCount++;
        try {
          totalBytes += Number(item.file.getSize()) || 0;
        } catch (eSize) {
          // ignore
        }
      });
      return;
    }

    if (seenSourceIds[sel.driveId]) {
      return;
    }
    seenSourceIds[sel.driveId] = true;
    var sourceFile;
    try {
      sourceFile = DriveApp.getFileById(sel.driveId);
    } catch (eFile) {
      throw catalogError_('INVALID_FILE', 'Drive file not found: ' + sel.driveId);
    }
    fileCount++;
    try {
      totalBytes += Number(sourceFile.getSize()) || 0;
    } catch (eSize2) {
      // ignore
    }
  });

  if (fileCount > IMPORT_DRIVE_JOB_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов (' +
        fileCount +
        '). Максимум за одну очередь: ' +
        IMPORT_DRIVE_JOB_MAX_FILES_ +
        '.'
    );
  }

  return {
    ok: true,
    fileCount: fileCount,
    totalBytes: totalBytes,
    folderCount: folderCount
  };
}

/**
 * Единый импорт с Drive: файлы и/или папки в одной очереди Jobs.
 *
 * @param {{
 *   targetFolderId: string,
 *   mode?: 'copy'|'move',
 *   items?: Array<{ kind?: 'file'|'folder', id?: string, url?: string }>,
 *   driveUrls?: string[]
 * }} input
 * @returns {{
 *   ok: true,
 *   queued: true,
 *   kind: 'drive',
 *   jobId: string,
 *   fileCount: number,
 *   folderCount: number,
 *   mode: 'copy'|'move'
 * }}
 */
function importDriveSelection(input) {
  assertCatalogReady_();
  assertNoActiveCatalogJobs_();

  input = input || {};
  var targetFolderId = String(input.targetFolderId || '').trim();
  var mode = String(input.mode || 'copy').trim().toLowerCase();
  if (!targetFolderId) {
    throw catalogError_('INVALID_INPUT', 'targetFolderId is required.');
  }
  if (mode !== 'copy' && mode !== 'move') {
    throw catalogError_('INVALID_INPUT', 'mode must be copy or move.');
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

  var selections = normalizeDriveImportSelections_(input);
  if (!selections.length) {
    throw catalogError_('INVALID_INPUT', 'Выберите файлы/папки или укажите ссылки.');
  }

  var treeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!treeSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  var jobItems = [];
  var fileRows = [];
  var folderCount = 0;
  var seenSourceIds = {};

  selections.forEach(function (sel) {
    if (sel.kind === 'folder') {
      var folderResult = appendImportDriveFolderToJob_(
        sel.driveId,
        targetFolderId,
        userEmail,
        treeSheet,
        seenSourceIds
      );
      folderCount += folderResult.folderCount;
      folderResult.items.forEach(function (it) {
        jobItems.push(it);
      });
      (folderResult.fileRows || []).forEach(function (row) {
        fileRows.push(row);
      });
      return;
    }
    if (seenSourceIds[sel.driveId]) {
      return;
    }
    seenSourceIds[sel.driveId] = true;
    var built = buildPendingImportFileItem_(sel.driveId, targetFolderId);
    jobItems.push(built.item);
    fileRows.push(built.row);
  });
  appendCatalogFileRowsBatch_(fileRows);

  if (jobItems.length > IMPORT_DRIVE_JOB_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов (' +
        jobItems.length +
        '). Максимум за одну очередь: ' +
        IMPORT_DRIVE_JOB_MAX_FILES_ +
        '.'
    );
  }
  if (!jobItems.length && !folderCount) {
    throw catalogError_('INVALID_INPUT', 'Нечего импортировать.');
  }
  if (!jobItems.length) {
    // только пустые папки — Jobs не нужен
    return {
      ok: true,
      queued: false,
      kind: 'drive',
      jobId: '',
      fileCount: 0,
      folderCount: folderCount,
      mode: mode,
      displayName: 'Папки без файлов'
    };
  }

  var jobId = enqueueCatalogJob_(
    'import_drive',
    {
      scenario: 'drive',
      mode: mode,
      targetFolderId: targetFolderId,
      phase: 'work',
      items: jobItems
    },
    userEmail
  );
  ensureCatalogJobsTrigger_();
  kickCatalogJobsProcessing_();

  return {
    ok: true,
    queued: true,
    kind: 'drive',
    jobId: jobId,
    fileCount: jobItems.length,
    folderCount: folderCount,
    mode: mode,
    displayName: jobItems.length + ' файл(ов)'
  };
}

/**
 * @param {Object} input
 * @returns {Array<{ kind: 'file'|'folder', driveId: string }>}
 */
function normalizeDriveImportSelections_(input) {
  var out = [];
  var seen = {};

  function add(kind, driveId) {
    driveId = String(driveId || '').trim();
    if (!driveId || seen[driveId]) {
      return;
    }
    seen[driveId] = true;
    out.push({ kind: kind, driveId: driveId });
  }

  (input.items || []).forEach(function (it) {
    if (!it) {
      return;
    }
    var kind = String(it.kind || '').toLowerCase();
    var id = String(it.id || '').trim();
    var url = String(it.url || '').trim();
    if (!id && url) {
      kind = resolveDriveImportKind_(url);
      id = kind === 'folder' ? parseDriveFolderId_(url) : parseDriveFileId_(url);
    }
    if (!id) {
      return;
    }
    if (kind !== 'file' && kind !== 'folder') {
      kind = resolveDriveImportKind_(id);
    }
    add(kind, id);
  });

  var urls = [];
  if (Array.isArray(input.driveUrls)) {
    input.driveUrls.forEach(function (u) {
      var s = String(u || '').trim();
      if (s) {
        urls.push(s);
      }
    });
  }
  var raw = String(input.driveUrl || '').trim();
  if (raw) {
    raw.split(/\r?\n/).forEach(function (line) {
      var s = String(line || '').trim();
      if (s) {
        urls.push(s);
      }
    });
  }
  urls.forEach(function (url) {
    var kind = resolveDriveImportKind_(url);
    var id = kind === 'folder' ? parseDriveFolderId_(url) : parseDriveFileId_(url);
    add(kind, id);
  });

  return out;
}

/**
 * @param {string} sourceFolderId
 * @param {string} targetFolderId
 * @param {string} userEmail
 * @param {GoogleAppsScript.Spreadsheet.Sheet} treeSheet
 * @param {Object.<string, boolean>} seenSourceIds
 * @returns {{ folderCount: number, items: Array }}
 */
function appendImportDriveFolderToJob_(
  sourceFolderId,
  targetFolderId,
  userEmail,
  treeSheet,
  seenSourceIds
) {
  var sourceFolder;
  try {
    sourceFolder = DriveApp.getFolderById(sourceFolderId);
  } catch (e) {
    throw catalogError_('INVALID_FOLDER', 'Drive folder not found or not accessible.');
  }

  var walk = walkDriveFolderForImport_(sourceFolder, IMPORT_DRIVE_JOB_MAX_FILES_);
  if (walk.fileCount > IMPORT_DRIVE_JOB_MAX_FILES_) {
    throw catalogError_(
      'IMPORT_TOO_LARGE',
      'Слишком много файлов в папке «' +
        sourceFolder.getName() +
        '» (' +
        walk.fileCount +
        '). Максимум: ' +
        IMPORT_DRIVE_JOB_MAX_FILES_ +
        '.'
    );
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
    treeRows.push([virtualId, parentVirtualId, node.name, now, false]);
    folderCount++;
  });
  if (treeRows.length) {
    treeSheet.getRange(treeSheet.getLastRow() + 1, 1, treeRows.length, 5).setValues(treeRows);
  }

  var rootVirtualId = driveToVirtual[sourceFolderId];
  applyDriveFolderAclToCatalogFolder_(sourceFolder, rootVirtualId, userEmail);

  var items = [];
  var fileRows = [];
  walk.files.forEach(function (item) {
    var sourceFileId = item.file.getId();
    if (seenSourceIds[sourceFileId]) {
      return;
    }
    seenSourceIds[sourceFileId] = true;
    var virtualParentId = driveToVirtual[item.parentDriveFolderId];
    if (!virtualParentId) {
      throw catalogError_('IMPORT_WALK_FAILED', 'Virtual folder missing for file parent.');
    }
    var built = buildPendingImportFileItem_(sourceFileId, virtualParentId, item.file.getName());
    items.push(built.item);
    fileRows.push(built.row);
  });

  return { folderCount: folderCount, items: items, fileRows: fileRows };
}

/**
 * Собирает pending item + строку Files без записи на лист.
 *
 * @param {string} sourceFileId
 * @param {string} parentFolderId
 * @param {string=} displayName
 * @returns {{ item: Object, row: Object }}
 */
function buildPendingImportFileItem_(sourceFileId, parentFolderId, displayName) {
  var name = displayName || '';
  if (!name) {
    try {
      name = DriveApp.getFileById(sourceFileId).getName();
    } catch (e) {
      throw catalogError_('INVALID_FILE', 'Drive file not found: ' + sourceFileId);
    }
  }
  var catalogId = Utilities.getUuid();
  return {
    item: {
      catalogId: catalogId,
      sourceFileId: sourceFileId,
      parentFolderId: parentFolderId,
      displayName: name,
      aclDone: false,
      copyDone: false,
      error: ''
    },
    row: {
      catalogId: catalogId,
      folderId: parentFolderId,
      fileId: '',
      displayName: name,
      sizeBytes: 0,
      driveModifiedAt: '',
      sourceFileId: sourceFileId,
      mimeType: '',
      status: 'pending'
    }
  };
}

/**
 * @deprecated используйте buildPendingImportFileItem_ + appendCatalogFileRowsBatch_
 */
function createPendingImportFileItem_(sourceFileId, parentFolderId, displayName) {
  var built = buildPendingImportFileItem_(sourceFileId, parentFolderId, displayName);
  appendCatalogFileRowsBatch_([built.row]);
  return built.item;
}
