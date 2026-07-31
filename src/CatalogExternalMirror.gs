/**
 * §20.1 — красное зеркало: ссылка на внешнюю папку Drive.
 * Просмотр — live Drive глазами текущего пользователя; в Sheet только строка-ссылка.
 */

/** @const {number} Размер страницы списка во внешней зоне панели. */
var EXTERNAL_MIRROR_PAGE_SIZE_ = 40;

/**
 * Создать красное зеркало внешней папки Drive в parent.
 *
 * @param {{
 *   parentFolderId: string,
 *   driveFolderId: string
 * }} input
 * @returns {{ ok: true, folder: Object }}
 */
function createCatalogExternalMirror(input) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var driveFolderId = String(input.driveFolderId || '').trim();

  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!driveFolderId) {
    throw catalogError_('INVALID_INPUT', 'Выберите папку Drive для зеркала.');
  }
  if (
    driveFolderId === 'root' ||
    driveFolderId === 'sharedWithMe' ||
    driveFolderId === 'sharedDrives'
  ) {
    throw catalogError_('INVALID_INPUT', 'Выберите конкретную папку, не корень навигатора.');
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
  var meta = fetchExternalDriveFolderMeta_(token, driveFolderId);
  if (!meta || !meta.id) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Папка Drive не найдена или нет доступа.');
  }
  if (String(meta.mimeType || '') !== DRIVE_FOLDER_MIME_) {
    throw catalogError_('INVALID_INPUT', 'Можно зеркалить только папку Drive.');
  }

  var name = String(meta.name || '').trim() || driveFolderId;
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
      mirrorOfFolderId: '',
      mirrorOfDriveFolderId: String(meta.id)
    }
  ]);
  bumpCatalogRev_();

  return {
    ok: true,
    folder: {
      id: folderId,
      parentFolderId: parentFolderId,
      name: name,
      sizeBytes: 0,
      fileCount: 0,
      modifiedAt: formatCatalogDate_(now),
      isSystem: false,
      isMirror: true,
      isExternalMirror: true,
      mirrorOfFolderId: '',
      mirrorOfDriveFolderId: String(meta.id),
      editors: [],
      commenters: [],
      readers: []
    }
  };
}

/**
 * Список содержимого внешней папки в контексте красного зеркала.
 * Любой пользователь каталога; Drive — от текущего OAuth.
 *
 * @param {{
 *   mirrorCatalogId: string,
 *   driveFolderId?: string,
 *   pageToken?: string
 * }} input
 * @returns {Object}
 */
function listExternalMirrorFolder(input) {
  assertCatalogReadyLight_();
  ensureCatalogSchemaUpToDate_();

  input = input || {};
  var mirrorCatalogId = String(input.mirrorCatalogId || '').trim();
  var pageToken = String(input.pageToken || '').trim();
  if (!mirrorCatalogId) {
    throw catalogError_('INVALID_INPUT', 'mirrorCatalogId is required.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertActiveCatalogUser_(userEmail);

  var engine = createAclEngine_();
  var mirrorRow = engine.foldersById[mirrorCatalogId];
  if (!mirrorRow || !isExternalMirrorFolderRow_(mirrorRow)) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Красное зеркало не найдено.');
  }

  var rootDriveId = String(mirrorRow.mirror_of_drive_folder_id || '').trim();
  var driveFolderId = String(input.driveFolderId || '').trim() || rootDriveId;
  var token = ScriptApp.getOAuthToken();

  var rootMeta = fetchExternalDriveFolderMetaSafe_(token, rootDriveId);
  if (rootMeta && rootMeta.code === 'GONE') {
    deleteCatalogMirrorFolderRows_([mirrorCatalogId]);
    bumpCatalogRev_();
    return {
      ok: false,
      code: 'GONE',
      message: 'Больше нет такой папки, удаляю зеркало',
      deletedMirrorId: mirrorCatalogId,
      parentFolderId: mirrorRow.parent_folder_id ? String(mirrorRow.parent_folder_id) : null
    };
  }
  if (rootMeta && rootMeta.code === 'NO_ACCESS') {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа'
    };
  }
  if (!rootMeta || !rootMeta.id) {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа'
    };
  }

  if (driveFolderId !== rootDriveId) {
    var under = isDriveFolderUnderExternalRoot_(token, driveFolderId, rootDriveId);
    if (under === 'GONE') {
      return {
        ok: false,
        code: 'NO_ACCESS',
        message: 'Нет доступа'
      };
    }
    if (under === 'NO_ACCESS') {
      return {
        ok: false,
        code: 'NO_ACCESS',
        message: 'Нет доступа'
      };
    }
    if (!under) {
      throw catalogError_('NOT_ALLOWED', 'Нельзя выйти за пределы зеркалируемой ветки.');
    }
  }

  var folderMeta =
    driveFolderId === rootDriveId
      ? rootMeta
      : fetchExternalDriveFolderMetaSafe_(token, driveFolderId);
  if (folderMeta && folderMeta.code === 'GONE') {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа'
    };
  }
  if (folderMeta && folderMeta.code === 'NO_ACCESS') {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа'
    };
  }
  if (!folderMeta || !folderMeta.id) {
    return {
      ok: false,
      code: 'NO_ACCESS',
      message: 'Нет доступа'
    };
  }

  var parentDriveId = null;
  if (driveFolderId === rootDriveId) {
    parentDriveId = null;
  } else if (folderMeta.parents && folderMeta.parents.length) {
    parentDriveId = String(folderMeta.parents[0]);
  }

  var pathParts = buildExternalMirrorPathParts_(token, driveFolderId, rootDriveId, rootMeta.name);
  var pathNames = pathParts.names;
  var pathIds = pathParts.ids;

  var listed = listExternalDriveChildrenPage_(token, {
    folderId: String(folderMeta.id),
    pageToken: pageToken,
    driveId: String(folderMeta.driveId || '').trim()
  });

  return {
    ok: true,
    mirrorCatalogId: mirrorCatalogId,
    mirrorName: String(mirrorRow.name || rootMeta.name || ''),
    rootDriveFolderId: rootDriveId,
    driveFolderId: String(folderMeta.id),
    driveFolderName: String(folderMeta.name || folderMeta.id),
    parentDriveFolderId: parentDriveId,
    pathNames: pathNames,
    pathIds: pathIds,
    catalogParentFolderId: mirrorRow.parent_folder_id
      ? String(mirrorRow.parent_folder_id)
      : null,
    items: listed.items,
    nextPageToken: listed.nextPageToken
  };
}

/**
 * @param {string} token
 * @param {string} folderId
 * @returns {Object}
 */
function fetchExternalDriveFolderMeta_(token, folderId) {
  return driveImportFetchJson_(
    'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(folderId) +
      '?fields=' +
      encodeURIComponent('id,name,mimeType,parents,driveId') +
      '&supportsAllDrives=true',
    token
  );
}

/**
 * @param {string} token
 * @param {string} folderId
 * @returns {Object|null} meta or { code: 'GONE'|'NO_ACCESS' }
 */
function fetchExternalDriveFolderMetaSafe_(token, folderId) {
  var response = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(folderId) +
      '?fields=' +
      encodeURIComponent('id,name,mimeType,parents,driveId') +
      '&supportsAllDrives=true',
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
    return { code: 'NO_ACCESS' };
  }
  if (code === 404) {
    return { code: 'GONE' };
  }
  if (code === 403 || code === 401) {
    return { code: 'NO_ACCESS' };
  }
  if (code < 200 || code >= 300) {
    return { code: 'NO_ACCESS' };
  }
  return body;
}

/**
 * @param {string} token
 * @param {string} folderId
 * @param {string} rootId
 * @returns {boolean|'GONE'|'NO_ACCESS'}
 */
function isDriveFolderUnderExternalRoot_(token, folderId, rootId) {
  var id = String(folderId || '').trim();
  var root = String(rootId || '').trim();
  if (!id || !root) {
    return false;
  }
  if (id === root) {
    return true;
  }
  var seen = {};
  var guard = 0;
  while (id && !seen[id] && guard < 64) {
    seen[id] = true;
    guard++;
    if (id === root) {
      return true;
    }
    var meta = fetchExternalDriveFolderMetaSafe_(token, id);
    if (meta && meta.code === 'GONE') {
      return 'GONE';
    }
    if (meta && meta.code === 'NO_ACCESS') {
      return 'NO_ACCESS';
    }
    if (!meta || !meta.parents || !meta.parents.length) {
      return false;
    }
    id = String(meta.parents[0]);
  }
  return false;
}

/**
 * @param {string} token
 * @param {string} folderId
 * @param {string} rootId
 * @param {string} rootName
 * @returns {{ names: string[], ids: string[] }}
 */
function buildExternalMirrorPathParts_(token, folderId, rootId, rootName) {
  var names = [];
  var ids = [];
  var id = String(folderId || '').trim();
  var root = String(rootId || '').trim();
  var seen = {};
  var guard = 0;
  while (id && !seen[id] && guard < 64) {
    seen[id] = true;
    guard++;
    var meta =
      id === root
        ? { id: root, name: rootName || root }
        : fetchExternalDriveFolderMetaSafe_(token, id);
    if (!meta || meta.code) {
      break;
    }
    names.unshift(String(meta.name || id));
    ids.unshift(String(meta.id || id));
    if (id === root) {
      break;
    }
    if (!meta.parents || !meta.parents.length) {
      break;
    }
    id = String(meta.parents[0]);
  }
  if (!names.length) {
    names.push(String(rootName || rootId || 'Drive'));
    ids.push(String(rootId || ''));
  }
  return { names: names, ids: ids };
}

/**
 * @param {string} token
 * @param {string} folderId
 * @param {string} rootId
 * @param {string} rootName
 * @returns {string[]}
 */
function buildExternalMirrorPathNames_(token, folderId, rootId, rootName) {
  return buildExternalMirrorPathParts_(token, folderId, rootId, rootName).names;
}

/**
 * @param {string} token
 * @param {{ folderId: string, pageToken: string, driveId: string }} opts
 * @returns {{ items: Array, nextPageToken: (string|null) }}
 */
function listExternalDriveChildrenPage_(token, opts) {
  var safeId = String(opts.folderId || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var q = "'" + safeId + "' in parents and trashed = false";
  var params = [
    'q=' + encodeURIComponent(q),
    'pageSize=' + EXTERNAL_MIRROR_PAGE_SIZE_,
    'fields=' +
      encodeURIComponent(
        'nextPageToken,files(id,name,mimeType,size,modifiedTime,webViewLink)'
      ),
    'supportsAllDrives=true',
    'includeItemsFromAllDrives=true',
    'orderBy=' + encodeURIComponent('folder,name')
  ];
  if (opts.driveId) {
    params.push('corpora=drive');
    params.push('driveId=' + encodeURIComponent(String(opts.driveId)));
  } else {
    params.push('corpora=allDrives');
  }
  if (opts.pageToken) {
    params.push('pageToken=' + encodeURIComponent(String(opts.pageToken)));
  }

  var body = driveImportFetchJson_(
    'https://www.googleapis.com/drive/v3/files?' + params.join('&'),
    token
  );

  var items = [];
  (body.files || []).forEach(function (f) {
    if (!f || !f.id) {
      return;
    }
    var mime = String(f.mimeType || '');
    if (mime === 'application/vnd.google-apps.shortcut') {
      return;
    }
    var kind = mime === DRIVE_FOLDER_MIME_ ? 'folder' : 'file';
    var sizeBytes = null;
    if (kind === 'file' && f.size != null && f.size !== '') {
      sizeBytes = parseNumber_(f.size);
    }
    items.push({
      kind: kind,
      id: String(f.id),
      name: String(f.name || f.id),
      mimeType: mime,
      sizeBytes: sizeBytes,
      modifiedAt: f.modifiedTime ? formatCatalogDate_(f.modifiedTime) : null,
      openUrl:
        kind === 'file'
          ? buildCatalogFileOpenUrl_(String(f.id), mime, f.webViewLink || '')
          : '',
      editors: [],
      commenters: [],
      readers: [],
      approved: false,
      approvedBy: '',
      approvedByName: '',
      isSystem: false,
      isMirror: false,
      isExternalMirror: false,
      isExternalItem: true
    });
  });

  return {
    items: items,
    nextPageToken: body.nextPageToken ? String(body.nextPageToken) : null
  };
}
