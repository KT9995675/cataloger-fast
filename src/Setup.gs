/**
 * §4.3 — ежедневный отзыв временных доступов на Drive (04:00 Москва).
 * Заглушка: логика — в следующих функциях.
 */
function revokeTemporaryDriveAccess() {
  // TODO: реализовать отзыв доступов (§4.3, Jobs.revoke_access)
}

/**
 * §15 — полная инициализация каталога (один раз).
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

  sheet.appendRow([virtualRootFolderId, '', virtualRootName, now, false]);
  sheet.appendRow([trashFolderId, virtualRootFolderId, 'Корзина', now, true]);
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Users sheet is missing.');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === controllerEmail.toLowerCase()) {
      if (String(data[i][1]) !== 'controller') {
        sheet.getRange(i + 1, 2).setValue('controller');
      }
      return;
    }
  }

  sheet.appendRow([controllerEmail, 'controller', now, '']);
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
