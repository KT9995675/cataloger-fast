/**
 * §4.3 — ежедневный отзыв временных доступов на Drive (04:00 Москва).
 * Заглушка: логика — в следующих функциях.
 */
function revokeTemporaryDriveAccess() {
  // TODO: реализовать отзыв доступов (§4.3, Jobs.revoke_access)
}

/**
 * §15 — полная инициализация каталога (один раз).
 * Для шаблона / «Первый запуск» с очисткой хвостов — `runFirstLaunch`.
 *
 * @param {{ driveFolderUrl: string, virtualRootName: string }} input
 * @returns {{
 *   ok: true,
 *   schemaVersion: string,
 *   catalogRootFolderId: string,
 *   catalogVirtualRootFolderId: string,
 *   trashFolderId: string,
 *   controllerEmail: string,
 *   setupAt: string
 * }}
 */
function setupCatalog(input) {
  input = input || {};
  var driveFolderUrl = String(input.driveFolderUrl || '').trim();
  var virtualRootName = String(input.virtualRootName || '').trim();

  if (!driveFolderUrl) {
    throw catalogError_('INVALID_INPUT', 'driveFolderUrl is required.');
  }
  if (!virtualRootName) {
    throw catalogError_('INVALID_INPUT', 'virtualRootName is required.');
  }

  assertSetupAllowed_();
  return completeCatalogSetup_(driveFolderUrl, virtualRootName);
}

/**
 * §15.5 — «Первый запуск»: очистка хвостов шаблона + setup.
 * Только владелец таблицы; только если каталог ещё не инициализирован.
 *
 * @param {{ driveFolderUrl: string, virtualRootName: string }} input
 * @returns {{
 *   ok: true,
 *   schemaVersion: string,
 *   catalogRootFolderId: string,
 *   catalogVirtualRootFolderId: string,
 *   trashFolderId: string,
 *   controllerEmail: string,
 *   setupAt: string,
 *   cleared: boolean
 * }}
 */
function runFirstLaunch(input) {
  input = input || {};
  var driveFolderUrl = String(input.driveFolderUrl || '').trim();
  var virtualRootName = String(input.virtualRootName || '').trim();

  if (!driveFolderUrl) {
    throw catalogError_('INVALID_INPUT', 'Укажите ссылку на пустую папку Google Drive.');
  }
  if (!virtualRootName) {
    throw catalogError_('INVALID_INPUT', 'Укажите имя корневой папки каталога.');
  }

  var controllerEmail = assertSpreadsheetOwner_();
  var state = isCatalogInitialized();
  if (state.initialized) {
    throw catalogError_(
      'CATALOG_ALREADY_INITIALIZED',
      'Каталог уже инициализирован. «Первый запуск» недоступен.'
    );
  }

  // Проверяем папку до очистки хвостов шаблона.
  validateAndResolveDriveFolder_(driveFolderUrl, controllerEmail);

  clearCatalogDataForFirstLaunch_();
  var result = completeCatalogSetup_(driveFolderUrl, virtualRootName);
  result.cleared = true;
  result.controllerEmail = result.controllerEmail || controllerEmail;
  return result;
}

/**
 * @param {string} driveFolderUrl
 * @param {string} virtualRootName
 * @returns {Object}
 */
function completeCatalogSetup_(driveFolderUrl, virtualRootName) {
  var controllerEmail = assertSpreadsheetOwner_();

  var sheetsResult = setupCatalogSheets();
  var catalogRootFolderId = validateAndResolveDriveFolder_(driveFolderUrl, controllerEmail);

  var virtualRootFolderId = Utilities.getUuid();
  var trashFolderId = '__TRASH__';
  var now = new Date();

  writeVirtualTreeBootstrap_(virtualRootFolderId, virtualRootName, trashFolderId, now);
  writeControllerUser_(controllerEmail, now);
  writeCatalogDocumentProperties_({
    schemaVersion: SCHEMA_VERSION_,
    catalogRootFolderId: catalogRootFolderId,
    catalogVirtualRootFolderId: virtualRootFolderId,
    controllerEmail: controllerEmail,
    setupAt: now.toISOString()
  });
  ensureDailyRevokeTrigger_();
  ensureCatalogJobsTrigger_();
  bumpCatalogRev_();

  return {
    ok: true,
    schemaVersion: SCHEMA_VERSION_,
    catalogRootFolderId: catalogRootFolderId,
    catalogVirtualRootFolderId: virtualRootFolderId,
    trashFolderId: trashFolderId,
    controllerEmail: controllerEmail,
    setupAt: now.toISOString(),
    sheets: sheetsResult
  };
}

/**
 * Очищает строки данных листов каталога и DocumentProperties (для шаблона).
 * Не вызывать для уже инициализированного каталога.
 */
function clearCatalogDataForFirstLaunch_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schema = getCatalogSheetSchema_();
  Object.keys(schema).forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), schema[sheetName].length, 1);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
    // Ensure header row matches schema
    sheet.getRange(1, 1, 1, schema[sheetName].length).setValues([schema[sheetName]]);
    if (sheet.getFrozenRows() < 1) {
      sheet.setFrozenRows(1);
    }
  });

  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty(PROP_SCHEMA_VERSION_);
  props.deleteProperty(PROP_CATALOG_ROOT_FOLDER_ID_);
  props.deleteProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_);
  props.deleteProperty(PROP_CONTROLLER_EMAIL_);
  props.deleteProperty(PROP_SETUP_AT_);
  props.deleteProperty(PROP_CATALOG_REV_);
}

function assertSetupAllowed_() {
  var state = isCatalogInitialized();
  if (state.initialized) {
    throw catalogError_('CATALOG_ALREADY_INITIALIZED', 'Catalog is already initialized.');
  }
  var hasPartialProps = !!(
    state.schemaVersion ||
    state.catalogRootFolderId ||
    state.catalogVirtualRootFolderId
  );
  if (hasPartialProps) {
    throw catalogError_(
      'CATALOG_INCONSISTENT',
      'Catalog configuration is incomplete or corrupt. Contact the controller.'
    );
  }
}

/**
 * @returns {string} owner email
 */
function assertSpreadsheetOwner_() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required to run setup.');
  }
  if (!isSpreadsheetOwnerEmail_(email)) {
    throw catalogError_('NOT_OWNER', 'Only the spreadsheet owner can initialize the catalog.');
  }
  return email;
}

/**
 * @param {string} email
 * @returns {boolean}
 */
function isSpreadsheetOwnerEmail_(email) {
  if (!email) {
    return false;
  }

  var controllerEmail = PropertiesService.getDocumentProperties().getProperty(
    PROP_CONTROLLER_EMAIL_
  );
  if (controllerEmail && controllerEmail.toLowerCase() === email.toLowerCase()) {
    return true;
  }

  try {
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var file = DriveApp.getFileById(spreadsheetId);
    var owner = file.getOwner();
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
 * @param {string} url
 * @param {string} ownerEmail
 * @returns {string} folder id
 */
function validateAndResolveDriveFolder_(url, ownerEmail) {
  var folderId = parseDriveFolderId_(url);
  var folder;

  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw catalogError_('INVALID_FOLDER_URL', 'Drive folder not found or not accessible.');
  }

  var folderOwner = folder.getOwner();
  if (!folderOwner || folderOwner.getEmail().toLowerCase() !== ownerEmail.toLowerCase()) {
    throw catalogError_('FOLDER_NOT_OWNED', 'Drive folder must belong to the spreadsheet owner.');
  }

  if (folder.getFiles().hasNext() || folder.getFolders().hasNext()) {
    throw catalogError_('FOLDER_NOT_EMPTY', 'Drive folder must be empty.');
  }

  return folderId;
}

/**
 * @param {string} url
 * @returns {string}
 */
function parseDriveFolderId_(url) {
  var foldersMatch = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (foldersMatch) {
    return foldersMatch[1];
  }
  var idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch) {
    return idMatch[1];
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(url)) {
    return url;
  }
  throw catalogError_('INVALID_FOLDER_URL', 'Cannot parse Drive folder id from URL.');
}

/**
 * @param {string} virtualRootFolderId
 * @param {string} virtualRootName
 * @param {string} trashFolderId
 * @param {Date} now
 */
function writeVirtualTreeBootstrap_(virtualRootFolderId, virtualRootName, trashFolderId, now) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Tree sheet is missing.');
  }

  if (findTreeRowByFolderId_(sheet, virtualRootFolderId) || findTreeRowByFolderId_(sheet, trashFolderId)) {
    throw catalogError_('CATALOG_INCONSISTENT', 'Tree already contains bootstrap folders.');
  }

  var headers = getCatalogSheetSchema_().Tree;
  var rootRow = [];
  var trashRow = [];
  headers.forEach(function (h) {
    if (h === 'folder_id') {
      rootRow.push(virtualRootFolderId);
      trashRow.push(trashFolderId);
    } else if (h === 'parent_folder_id') {
      rootRow.push('');
      trashRow.push(virtualRootFolderId);
    } else if (h === 'name') {
      rootRow.push(virtualRootName);
      trashRow.push('## Корзина');
    } else if (h === 'folder_created_at') {
      rootRow.push(now);
      trashRow.push(now);
    } else if (h === 'is_system') {
      rootRow.push(false);
      trashRow.push(true);
    } else {
      rootRow.push('');
      trashRow.push('');
    }
  });
  sheet.appendRow(rootRow);
  sheet.appendRow(trashRow);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} folderId
 * @returns {boolean}
 */
function findTreeRowByFolderId_(sheet, folderId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === folderId) {
      return true;
    }
  }
  return false;
}

/**
 * @param {string} controllerEmail
 * @param {Date} now
 */
function writeControllerUser_(controllerEmail, now) {
  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Users sheet is missing.');
  }

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  var emailCol = headers.indexOf('email');
  var roleCol = headers.indexOf('login_role');
  var nameCol = headers.indexOf('display_name');
  if (emailCol < 0 || roleCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Users sheet headers are invalid.');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol] || '').toLowerCase() === controllerEmail.toLowerCase()) {
      if (String(data[i][roleCol]) !== 'controller') {
        sheet.getRange(i + 1, roleCol + 1).setValue('controller');
      }
      if (nameCol >= 0 && !String(data[i][nameCol] || '').trim()) {
        sheet.getRange(i + 1, nameCol + 1).setValue(controllerEmail);
      }
      return;
    }
  }

  appendOrEnsureUserRow_({
    email: controllerEmail,
    loginRole: 'controller',
    addedBy: '',
    displayName: controllerEmail
  });
}

/**
 * @param {{
 *   schemaVersion: string,
 *   catalogRootFolderId: string,
 *   catalogVirtualRootFolderId: string,
 *   controllerEmail: string,
 *   setupAt: string
 * }} config
 */
function writeCatalogDocumentProperties_(config) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty(PROP_SCHEMA_VERSION_)) {
    throw catalogError_('CATALOG_ALREADY_INITIALIZED', 'DocumentProperties already configured.');
  }

  props.setProperties({
    SCHEMA_VERSION: config.schemaVersion,
    CATALOG_ROOT_FOLDER_ID: config.catalogRootFolderId,
    CATALOG_VIRTUAL_ROOT_FOLDER_ID: config.catalogVirtualRootFolderId,
    CONTROLLER_EMAIL: config.controllerEmail,
    SETUP_AT: config.setupAt
  });
}

function ensureDailyRevokeTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'revokeTemporaryDriveAccess') {
      return;
    }
  }

  ScriptApp.newTrigger('revokeTemporaryDriveAccess')
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .inTimezone('Europe/Moscow')
    .create();
}
