/**
 * §0.4a — общие примитивы длинных операций (не фреймворк).
 * Ошибки: файл → failed, остальные в цепочке продолжаем (не рвём всю очередь).
 */

/** @const Правило ошибок длинных ops (§0.4a). */
var LONG_OPS_ERROR_POLICY_ = 'continue_chain';

/**
 * @param {string} message
 */
function longOpsSetStatus_(message) {
  setCatalogOpStatus_(message);
}

function longOpsClearStatus_() {
  clearCatalogOpStatus_();
}

/**
 * Нативные Google-файлы (Docs/Sheets/…) — getSize() часто stub 0/1 (§19.9.1 / B2).
 *
 * @param {string} mimeType
 * @returns {boolean}
 */
function isGoogleNativeMimeType_(mimeType) {
  return /^application\/vnd\.google-apps\./i.test(String(mimeType || ''));
}

/**
 * Не кэшировать stub: для Google MIME значения 0 и 1 — не надёжный размер.
 *
 * @param {number} sizeBytes
 * @param {string} mimeType
 * @returns {boolean}
 */
function isDriveSizeStub_(sizeBytes, mimeType) {
  var n = Number(sizeBytes) || 0;
  if (!isGoogleNativeMimeType_(mimeType)) {
    return false;
  }
  return n <= 1;
}

/**
 * size / quotaBytesUsed одним GET (только когда DriveApp stub).
 * При исчерпании urlfetch → 0 (не 1).
 *
 * @param {string} driveFileId
 * @returns {number}
 */
function fetchDriveQuotaOrSizeBytes_(driveFileId) {
  var id = String(driveFileId || '').trim();
  if (!id) {
    return 0;
  }
  try {
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(id) +
        '?supportsAllDrives=true&fields=' +
        encodeURIComponent('size,quotaBytesUsed'),
      {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      return 0;
    }
    var body = {};
    try {
      body = JSON.parse(resp.getContentText() || '{}');
    } catch (eParse) {
      return 0;
    }
    var raw =
      body.quotaBytesUsed != null && body.quotaBytesUsed !== ''
        ? body.quotaBytesUsed
        : body.size;
    var n = parseDriveSizeBytes_(raw);
    return n > 1 ? n : 0;
  } catch (eFetch) {
    return 0;
  }
}

/**
 * Размер для Files.size_bytes: DriveApp.getSize(), для Google stub — quotaBytesUsed;
 * никогда не пишем 1 у Google MIME.
 *
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string=} mimeTypeOpt
 * @returns {number}
 */
function resolveDriveFileSizeBytes_(file, mimeTypeOpt) {
  if (!file) {
    return 0;
  }
  var mime = String(mimeTypeOpt || '') || getDriveFileMimeType_(file) || '';
  var size = 0;
  try {
    size = Number(file.getSize()) || 0;
  } catch (eSize) {
    size = 0;
  }
  if (!isDriveSizeStub_(size, mime)) {
    return size > 0 ? size : 0;
  }
  var fromApi = 0;
  try {
    fromApi = fetchDriveQuotaOrSizeBytes_(file.getId());
  } catch (eApi) {
    fromApi = 0;
  }
  return fromApi > 1 ? fromApi : 0;
}

/**
 * Нормализация уже известного кэша (без Drive): stub 1 у Google → 0.
 *
 * @param {number} sizeBytes
 * @param {string} mimeType
 * @returns {number}
 */
function normalizeCatalogSizeBytes_(sizeBytes, mimeType) {
  var n = Number(sizeBytes) || 0;
  if (isDriveSizeStub_(n, mimeType)) {
    return 0;
  }
  return n > 0 ? n : 0;
}

/**
 * Meta файлов через DriveApp (+ quotaBytesUsed при Google stub).
 *
 * @param {string[]} driveFileIds
 * @returns {Object.<string, { ok: boolean, sizeBytes: number, mimeType: string, driveModifiedAt: (Date|string), error: string }>}
 */
function longOpsFetchDriveMetaBatch_(driveFileIds) {
  var ids = driveFileIds || [];
  var out = {};
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i] || '');
    if (!id) {
      continue;
    }
    try {
      var f = DriveApp.getFileById(id);
      var mimeType = getDriveFileMimeType_(f) || '';
      out[id] = {
        ok: true,
        sizeBytes: resolveDriveFileSizeBytes_(f, mimeType),
        mimeType: mimeType,
        driveModifiedAt: f.getLastUpdated(),
        error: ''
      };
    } catch (eMeta) {
      out[id] = {
        ok: false,
        sizeBytes: 0,
        mimeType: '',
        driveModifiedAt: '',
        error: (eMeta && eMeta.message) || String(eMeta) || 'Drive meta failed'
      };
    }
  }
  return out;
}
/* OLD: longOpsFetchDriveMetaBatch_ — только Number(f.getSize())||0, без учёта stub 1 у Google MIME. */

/**
 * @deprecated UrlFetch — жрёт дневную квоту; оставлен для редких fallback.
 * @param {string[]} driveFileIds
 * @returns {Object.<string, { ok: boolean, sizeBytes: number, mimeType: string, driveModifiedAt: string, error: string }>}
 */
function longOpsFetchDriveMetaBatchUrlFetch_(driveFileIds) {
  var ids = driveFileIds || [];
  var out = {};
  if (!ids.length) {
    return out;
  }
  var token = ScriptApp.getOAuthToken();
  var reqs = ids.map(function (id) {
    return {
      url:
        'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(String(id)) +
        '?supportsAllDrives=true&fields=' +
        encodeURIComponent('id,size,quotaBytesUsed,modifiedTime,mimeType'),
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };
  });
  var resps = UrlFetchApp.fetchAll(reqs);
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i] || '');
    var resp = resps[i];
    var code = resp ? resp.getResponseCode() : 0;
    var body = {};
    try {
      body = JSON.parse((resp && resp.getContentText()) || '{}');
    } catch (eParse) {
      body = {};
    }
    if (code < 200 || code >= 300 || !body.id) {
      out[id] = {
        ok: false,
        sizeBytes: 0,
        mimeType: '',
        driveModifiedAt: '',
        error: (body.error && body.error.message) || 'Drive meta failed'
      };
      continue;
    }
    out[id] = {
      ok: true,
      sizeBytes: parseDriveSizeBytes_(
        body.size != null && body.size !== '' ? body.size : body.quotaBytesUsed
      ),
      mimeType: String(body.mimeType || ''),
      driveModifiedAt: body.modifiedTime || '',
      error: ''
    };
  }
  return out;
}

/**
 * @deprecated UrlFetch permissions.list — жрёт квоту; ACL импорта → DriveApp.
 * @param {string[]} driveFileIds
 * @returns {Object.<string, Array<{ email: string, displayName: string, level: string }>>}
 */
function longOpsFetchDrivePermsBatch_(driveFileIds) {
  var ids = driveFileIds || [];
  var out = {};
  if (!ids.length) {
    return out;
  }
  var token = ScriptApp.getOAuthToken();
  var reqs = ids.map(function (id) {
    return {
      url:
        'https://www.googleapis.com/drive/v3/files/' +
        encodeURIComponent(String(id)) +
        '/permissions?supportsAllDrives=true&fields=' +
        encodeURIComponent('permissions(type,role,emailAddress,displayName)'),
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };
  });
  var resps = UrlFetchApp.fetchAll(reqs);
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i] || '');
    var slim = [];
    var resp = resps[i];
    var code = resp ? resp.getResponseCode() : 0;
    if (code >= 200 && code < 300) {
      var body = {};
      try {
        body = JSON.parse((resp && resp.getContentText()) || '{}');
      } catch (eP) {
        body = {};
      }
      var perms = body.permissions || [];
      for (var p = 0; p < perms.length; p++) {
        var perm = perms[p];
        if (!perm || String(perm.type || '') !== 'user') {
          continue;
        }
        var email = String(perm.emailAddress || '').trim();
        if (!email) {
          continue;
        }
        var level = drivePermissionRoleToCatalogLevel_(perm.role);
        if (!level) {
          continue;
        }
        slim.push({
          email: email,
          displayName: String(perm.displayName || '').trim() || email,
          level: level
        });
      }
    }
    out[id] = slim;
  }
  return out;
}

/**
 * Copy или place файла в плоскую папку каталога.
 *
 * @param {string} sourceDriveFileId
 * @param {'copy'|'move'} mode
 * @param {GoogleAppsScript.Drive.Folder} catalogRootFolder
 * @returns {GoogleAppsScript.Drive.File}
 */
function longOpsTransferIntoCatalogRoot_(sourceDriveFileId, mode, catalogRootFolder) {
  var sourceFile = DriveApp.getFileById(String(sourceDriveFileId));
  return placeFileInCatalogRoot_(sourceFile, catalogRootFolder, mode === 'move' ? 'move' : 'copy');
}

/**
 * Пометить файл каталога failed (ошибка одного файла не рвёт цепочку).
 *
 * @param {string} catalogId
 * @param {string} message
 * @param {Object=} filesBatch beginFilesUpdateBatch_
 */
function longOpsMarkFileFailed_(catalogId, message, filesBatch) {
  var msg = String(message || 'failed');
  if (filesBatch) {
    patchFilesBatchRow_(filesBatch, catalogId, {
      status: 'failed',
      lastError: msg
    });
    return;
  }
  markCatalogFileFailed_(catalogId, msg);
}
