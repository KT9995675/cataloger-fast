/** @const {string} Версия схемы листов — §3.11 */
var SCHEMA_VERSION_ = '0.3';

/**
 * §15.2 п.1 — создаёт листы каталога и строку заголовков (row 1), если отсутствуют.
 * Идемпотентна: при совпадающих заголовках данные не меняет.
 * Допускает **аддитивный** апгрейд: actual — префикс expected → дописывает столбцы.
 *
 * @returns {{ ok: true, schemaVersion: string, created: string[], existing: string[], upgraded: string[] }}
 */
function setupCatalogSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schema = getCatalogSheetSchema_();
  var created = [];
  var existing = [];
  var upgraded = [];
  var sheetNames = Object.keys(schema);

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var expectedHeaders = schema[sheetName];
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      sheet.setFrozenRows(1);
      created.push(sheetName);
      continue;
    }

    var upgradeResult = ensureSheetHeadersMatchOrUpgrade_(sheet, sheetName, expectedHeaders);
    if (upgradeResult === 'upgraded') {
      upgraded.push(sheetName);
    } else {
      existing.push(sheetName);
    }

    if (sheet.getFrozenRows() < 1) {
      sheet.setFrozenRows(1);
    }
  }

  return {
    ok: true,
    schemaVersion: SCHEMA_VERSION_,
    created: created,
    existing: existing,
    upgraded: upgraded
  };
}

/**
 * Для уже инициализированного каталога: дописывает новые столбцы схемы и обновляет SCHEMA_VERSION.
 * Безопасно вызывать при каждом snapshot / bootstrap.
 */
function ensureCatalogSchemaUpToDate_() {
  if (!areCatalogSheetsPresent_()) {
    return;
  }

  var result = setupCatalogSheets();
  var props = PropertiesService.getDocumentProperties();
  var current = props.getProperty(PROP_SCHEMA_VERSION_);
  if (current && current !== SCHEMA_VERSION_) {
    props.setProperty(PROP_SCHEMA_VERSION_, SCHEMA_VERSION_);
  }
  backfillEmptyUserDisplayNames_();
  return result;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} sheetName
 * @param {string[]} expectedHeaders
 * @returns {'ok'|'upgraded'}
 */
function ensureSheetHeadersMatchOrUpgrade_(sheet, sheetName, expectedHeaders) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var actualHeaders = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  while (actualHeaders.length && !actualHeaders[actualHeaders.length - 1]) {
    actualHeaders.pop();
  }

  if (headersEqualStrict_(actualHeaders, expectedHeaders)) {
    return 'ok';
  }

  if (actualHeaders.length < expectedHeaders.length) {
    var isPrefix = true;
    for (var i = 0; i < actualHeaders.length; i++) {
      if (actualHeaders[i] !== expectedHeaders[i]) {
        isPrefix = false;
        break;
      }
    }
    if (isPrefix) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      return 'upgraded';
    }
  }

  throw catalogError_(
    'SCHEMA_MISMATCH',
    'Sheet "' + sheetName + '": header row does not match catalog schema (SPEC §3.6–3.8).'
  );
}

/**
 * @returns {boolean}
 */
function areCatalogSheetsPresent_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = Object.keys(getCatalogSheetSchema_());
  for (var i = 0; i < sheetNames.length; i++) {
    if (!ss.getSheetByName(sheetNames[i])) {
      return false;
    }
  }
  return true;
}

/**
 * @returns {Object.<string, string[]>}
 */
function getCatalogSheetSchema_() {
  return {
    Tree: ['folder_id', 'parent_folder_id', 'name', 'folder_created_at', 'is_system'],
    Files: [
      'catalog_id',
      'folder_id',
      'file_id',
      'display_name',
      'size_bytes',
      'drive_modified_at',
      'approved',
      'approved_by',
      'approved_at',
      'status',
      'last_error',
      'source_file_id',
      'mime_type'
    ],
    Users: ['email', 'login_role', 'added_at', 'added_by', 'display_name'],
    ACL: [
      'acl_id',
      'object_type',
      'object_id',
      'principal_type',
      'principal_id',
      'permission_level'
    ],
    Groups: ['group_id', 'name', 'created_at', 'created_by'],
    GroupMembers: ['group_id', 'email', 'added_at'],
    Jobs: [
      'job_id',
      'job_type',
      'status',
      'catalog_id',
      'payload_json',
      'progress',
      'progress_message',
      'created_at',
      'started_at',
      'completed_at',
      'last_error',
      'created_by'
    ]
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} width
 * @returns {string[]}
 */
function readSheetHeaderRow_(sheet, width) {
  var lastCol = Math.max(sheet.getLastColumn(), width);
  if (lastCol < 1) {
    return [];
  }
  var row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return row.map(function (cell) {
    return String(cell == null ? '' : cell).trim();
  });
}

/**
 * Строгое совпадение заголовков 1-в-1 (SPEC: SCHEMA_MISMATCH при расхождении).
 * @param {string[]} actual
 * @param {string[]} expected
 * @returns {boolean}
 */
function headersEqualStrict_(actual, expected) {
  if (actual.length < expected.length) {
    return false;
  }
  for (var i = 0; i < expected.length; i++) {
    if (actual[i] !== expected[i]) {
      return false;
    }
  }
  for (var j = expected.length; j < actual.length; j++) {
    if (actual[j] !== '') {
      return false;
    }
  }
  return true;
}

/**
 * @param {string} code
 * @param {string} message
 * @returns {Error}
 */
function catalogError_(code, message) {
  var err = new Error(message);
  err.name = code;
  return err;
}
