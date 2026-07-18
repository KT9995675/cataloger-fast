/**
 * §7 — «Подписать»: утверждение файлов каталога.
 *
 * @param {{
 *   fileIds: string[]
 * }} input
 * @returns {{
 *   ok: true,
 *   approved: Array<{
 *     id: string,
 *     approvedBy: string,
 *     approvedByName: string
 *   }>,
 *   skippedAlready: number,
 *   skippedNoPermission: number,
 *   skippedNotFound: number
 * }}
 */
function approveCatalogFiles(input) {
  assertCatalogReady_();

  input = input || {};
  var fileIds = input.fileIds || [];
  if (!Array.isArray(fileIds) || !fileIds.length) {
    throw catalogError_('INVALID_INPUT', 'fileIds must be a non-empty array.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  var engine = createAclEngine_();
  var now = new Date();
  var approved = [];
  var skippedAlready = 0;
  var skippedNoPermission = 0;
  var skippedNotFound = 0;
  var seen = {};

  for (var i = 0; i < fileIds.length; i++) {
    var catalogId = String(fileIds[i] || '').trim();
    if (!catalogId || seen[catalogId]) {
      continue;
    }
    seen[catalogId] = true;

    var file = engine.filesByCatalogId[catalogId];
    if (!file) {
      skippedNotFound += 1;
      continue;
    }
    if (parseBoolean_(file.approved)) {
      skippedAlready += 1;
      continue;
    }

    if (!canApproveCatalogFile_(engine, userEmail, loginRole, catalogId)) {
      skippedNoPermission += 1;
      continue;
    }

    writeFileApprovedFields_(catalogId, true, userEmail, now);
    file.approved = true;
    file.approved_by = userEmail;
    file.approved_at = now;
    demoteExplicitEditorsOnFile_(engine, catalogId);

    approved.push({
      id: catalogId,
      approvedBy: userEmail,
      approvedByName: resolveUserLabelFromEngine_(engine, userEmail)
    });
  }

  return {
    ok: true,
    approved: approved,
    skippedAlready: skippedAlready,
    skippedNoPermission: skippedNoPermission,
    skippedNotFound: skippedNotFound
  };
}

/**
 * @param {Object} engine
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @param {string} catalogId
 * @returns {boolean}
 */
function canApproveCatalogFile_(engine, userEmail, loginRole, catalogId) {
  if (loginRole === 'controller') {
    return true;
  }
  var permission = getEffectivePermissionForUserFromEngine_(
    engine,
    'file',
    catalogId,
    userEmail
  );
  return permission === 'editor';
}

/**
 * @param {string} catalogId
 * @param {boolean} approved
 * @param {string} approvedBy
 * @param {Date} approvedAt
 */
function writeFileApprovedFields_(catalogId, approved, approvedBy, approvedAt) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Files');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Files');
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
  }

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var catalogCol = headers.indexOf('catalog_id');
  var approvedCol = headers.indexOf('approved');
  var approvedByCol = headers.indexOf('approved_by');
  var approvedAtCol = headers.indexOf('approved_at');
  if (catalogCol < 0 || approvedCol < 0 || approvedByCol < 0 || approvedAtCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Files sheet missing approved columns.');
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][catalogCol]) !== catalogId) {
      continue;
    }
    var row = i + 1;
    sheet.getRange(row, approvedCol + 1).setValue(approved === true);
    sheet.getRange(row, approvedByCol + 1).setValue(approvedBy || '');
    sheet.getRange(row, approvedAtCol + 1).setValue(approvedAt || '');
    return;
  }

  throw catalogError_('FILE_NOT_FOUND', 'File not found: ' + catalogId);
}

/**
 * Явные ACL файла: editor → commenter (§7).
 *
 * @param {Object} engine
 * @param {string} catalogId
 */
function demoteExplicitEditorsOnFile_(engine, catalogId) {
  var entries = readExplicitAclEntries_('file', catalogId, engine)
    .map(function (entry) {
      var level = entry.permissionLevel;
      if (level === 'editor') {
        level = 'commenter';
      }
      return {
        principalType: entry.principalType,
        principalId: entry.principalId,
        permissionLevel: level
      };
    })
    .filter(function (entry) {
      return entry.permissionLevel && entry.permissionLevel !== 'none';
    });

  replaceAclForObjects_(
    [{ objectType: 'file', objectId: catalogId }],
    entries,
    engine
  );
}
