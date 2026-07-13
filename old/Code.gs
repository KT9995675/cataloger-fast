/**
 * Каталогизатор Google Drive — Google Apps Script
 * Источник правды: локальный git → clasp push
 * Спецификация: SPEC.md
 *
 * LEGACY SNAPSHOT
 * --------------
 * Это сохранённая “старая версия” из репозитория `google-drive-cataloger`.
 * В проекте `Cataloger Fast` этот файл используется только как донор функций/идей.
 */

var SCHEMA_VERSION = '1.4.2';

var CATALOG_MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

var FOLDER_CODE_CATALOG = '0';
var FOLDER_CODE_TRASH = '99';
var IMPORT_DRIVE_FOLDER_NAME = '_Import';
var SETTING_IMPORT_FOLDER_ID = 'import_folder_id';
var FOLDER_CODE_ROOT_MIN = 10;
var FOLDER_CODE_ROOT_MAX = 98;
var CATALOG_TREE_DEPTH_MAX = 5;

var CATALOG_TREE_HEADERS = [
  'level_0', 'level_1', 'level_2', 'level_3', 'level_4', 'level_5'
];

var CATALOG_DATA_HEADERS = [
  'id', 'parent_id', 'folder_code', 'name', 'type', 'file_id', 'mime_type',
  'permissions_json', 'approved', 'editors', 'commenters', 'readers', 'trash_parent_id'
];

var CATALOG_VIEW_COLUMN_KEYS = CATALOG_TREE_HEADERS.concat([
  'type', 'approved', 'editors', 'commenters', 'readers'
]);

var CATALOG_VIEW_DISPLAY_HEADERS = CATALOG_TREE_HEADERS.concat([
  'Размер', 'Утверждение', 'Редактор', 'Комментатор', 'Чтение'
]);

var CATALOG_INDEX_HEADERS = [
  'catalog_id', 'type', 'file_id', 'folder_code', 'name', 'approved', 'data_row_index'
];

var CATALOG_VIEW_STRIPE_WHITE = '#ffffff';
var CATALOG_VIEW_STRIPE_TEAL = '#d9f2f0';

var CATALOG_VIEW_HEADERS = CATALOG_VIEW_COLUMN_KEYS;

var SHEET_HEADERS = {
  Settings: ['key', 'value'],
  Users: ['email', 'name', 'active'],
  Groups: ['group_id', 'group_name', 'member_email'],
  CatalogData: CATALOG_DATA_HEADERS,
  CatalogIndex: CATALOG_INDEX_HEADERS,
  Catalog: CATALOG_VIEW_DISPLAY_HEADERS
};

var CATALOG_VIEW_VISIBLE_COUNT = 11;

var DEFAULT_SETTINGS = [
  ['root_folder_id', ''],
  ['import_folder_id', ''],
  ['import_webapp_url', ''],
  ['import_webapp_secret', ''],
  ['schema_version', SCHEMA_VERSION]
];

/**
 * Создаёт структуру листов, заголовки, системные папки и защиту Catalog.
 * Идемпотентна: повторный запуск не дублирует данные и не затирает строки.
 * Запуск: из редактора Apps Script → setupSchema → Выполнить
 * (нужна привязка к Google Таблице).
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function setupSchema() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return {
        ok: false,
        error: 'NO_SPREADSHEET',
        message: 'Нет активной таблицы. Привяжите скрипт к Google Таблице.'
      };
    }

    var createdSheets = [];
    var ensuredHeaders = [];
    var seeded = { Settings: false, systemFolders: false };
    var hiddenSheets = [];
    var protectedSheets = [];

    var sheetOrder = ['Catalog', 'CatalogData', 'CatalogIndex', 'Users', 'Groups', 'Settings'];
    for (var i = 0; i < sheetOrder.length; i++) {
      var sheetName = sheetOrder[i];
      var result = ensureSheet_(ss, sheetName, SHEET_HEADERS[sheetName]);
      if (result.created) {
        createdSheets.push(sheetName);
      }
      if (result.headersEnsured) {
        ensuredHeaders.push(sheetName);
      }
    }

    var catalogViewSheet = ss.getSheetByName('Catalog');
    var catalogDataSheet = ss.getSheetByName('CatalogData');
    var catalogIndexSheet = ss.getSheetByName('CatalogIndex');
    configureCatalogViewSheet_(catalogViewSheet);
    configureCatalogDataSheet_(catalogDataSheet);
    configureCatalogIndexSheet_(catalogIndexSheet);
    protectedSheets.push(applyCatalogProtection_(catalogViewSheet));

    var settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet.isSheetHidden()) {
      settingsSheet.showSheet();
    }
    settingsSheet.hideSheet();
    hiddenSheets.push('Settings');
    catalogDataSheet.hideSheet();
    hiddenSheets.push('CatalogData');
    catalogIndexSheet.hideSheet();
    hiddenSheets.push('CatalogIndex');

    if (settingsSheet.getLastRow() < 2) {
      settingsSheet.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
      seeded.Settings = true;
    } else {
      upsertSetting_(settingsSheet, 'schema_version', SCHEMA_VERSION);
    }
    ensureImportWebAppSettings_(settingsSheet);
    var importWebApp = configureImportWebAppDuringSetup_(settingsSheet);
    var importFolderInfo = null;
    if (getSetting_(settingsSheet, 'root_folder_id')) {
      try {
        importFolderInfo = { folder_id: ensureCatalogImportDriveFolder_(ss) };
      } catch (importFolderErr) {
        importFolderInfo = { error: String(importFolderErr) };
      }
    }

    seeded.systemFolders = seedSystemFolders_(catalogDataSheet);
    var renderInfo = renderCatalogViewCore_(catalogViewSheet, catalogDataSheet);
    removeCatalogApprovalEditTriggers_();

    return {
      ok: true,
      data: {
        schema_version: SCHEMA_VERSION,
        created_sheets: createdSheets,
        ensured_headers: ensuredHeaders,
        hidden_sheets: hiddenSheets,
        protected_sheets: protectedSheets,
        seeded: seeded,
        tree_render: renderInfo,
        import_webapp: importWebApp,
        import_folder: importFolderInfo
      },
      message: buildSetupSchemaMessage_(importWebApp)
    };
  } catch (err) {
    return {
      ok: false,
      error: 'SETUP_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string[]} headers
 * @returns {{created: boolean, headersEnsured: boolean}}
 */
function ensureSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeaders = firstRow.join('') === '';
  if (!needsHeaders) {
    for (var h = 0; h < headers.length; h++) {
      if (firstRow[h] !== headers[h]) {
        needsHeaders = true;
        break;
      }
    }
  }
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return { created: created, headersEnsured: needsHeaders };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function configureCatalogViewSheet_(sheet) {
  sheet.getRange(1, 1, 1, CATALOG_VIEW_DISPLAY_HEADERS.length).setValues([CATALOG_VIEW_DISPLAY_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.showColumns(1, CATALOG_VIEW_COLUMN_KEYS.length);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function configureCatalogIndexSheet_(sheet) {
  sheet.getRange(1, 1, 1, CATALOG_INDEX_HEADERS.length).setValues([CATALOG_INDEX_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.showColumns(1, CATALOG_INDEX_HEADERS.length);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function configureCatalogDataSheet_(sheet) {
  sheet.showColumns(1, SHEET_HEADERS.CatalogData.length);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {{view: GoogleAppsScript.Spreadsheet.Sheet, data: GoogleAppsScript.Spreadsheet.Sheet}}
 */
function getCatalogSheets_(ss) {
  var view = ss.getSheetByName('Catalog');
  var data = ss.getSheetByName('CatalogData');
  var index = ss.getSheetByName('CatalogIndex');
  if (!view || !data) {
    throw new Error('Сначала выполните setupSchema.');
  }
  return { view: view, data: data, index: index };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getCatalogIndexSheet_(ss) {
  var sheet = ss.getSheetByName('CatalogIndex');
  if (!sheet) {
    throw new Error('Лист CatalogIndex не найден. Выполните setupSchema.');
  }
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrEnsureCatalogIndexSheet_(ss) {
  var sheet = ss.getSheetByName('CatalogIndex');
  if (!sheet) {
    ensureSheet_(ss, 'CatalogIndex', SHEET_HEADERS.CatalogIndex);
    sheet = ss.getSheetByName('CatalogIndex');
    configureCatalogIndexSheet_(sheet);
    sheet.hideSheet();
  }
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {string}
 */
function applyCatalogProtection_(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === 'CATALOG_DATA') {
      protections[i].remove();
    }
  }
  // Защита диапазона отключена: блокирует appendRow при запуске из меню.
  // Скрытые колонки + будущий onEdit вернут ручную блокировку.
  return sheet.getName();
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {boolean} true, если добавлены системные папки
 */
function seedSystemFolders_(sheet) {
  var col = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var existingCodes = readColumnValues_(sheet, col.folder_code);
  var added = false;

  var catalogId = null;
  if (!hasFolderCode_(existingCodes, FOLDER_CODE_CATALOG)) {
    catalogId = generateId_();
    appendCatalogDataRow_(sheet, col, {
      id: catalogId,
      name: FOLDER_CODE_CATALOG + '_Каталог',
      type: 'folder',
      folder_code: FOLDER_CODE_CATALOG,
      parent_id: '',
      permissions_json: '[]'
    });
    added = true;
  } else {
    catalogId = findIdByFolderCode_(sheet, col, FOLDER_CODE_CATALOG);
  }

  if (!hasFolderCode_(existingCodes, FOLDER_CODE_TRASH)) {
    appendCatalogDataRow_(sheet, col, {
      id: generateId_(),
      name: FOLDER_CODE_TRASH + '_Корзина',
      type: 'folder',
      folder_code: FOLDER_CODE_TRASH,
      parent_id: catalogId || '',
      permissions_json: '[]'
    });
    added = true;
  }

  return added;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} row
 */
function appendCatalogDataRow_(sheet, col, row) {
  appendCatalogDataRowsBatch_(sheet, col, [row]);
}

/**
 * @param {Object.<string, number>} col
 * @param {Object} row
 * @returns {Array}
 */
function catalogDataRowValuesFromObject_(col, row) {
  var values = new Array(SHEET_HEADERS.CatalogData.length);
  values[col.id - 1] = row.id || '';
  values[col.parent_id - 1] = readCell_(row.parent_id);
  values[col.folder_code - 1] = readCell_(row.folder_code);
  values[col.name - 1] = row.name || '';
  values[col.type - 1] = row.type || '';
  values[col.file_id - 1] = row.file_id || '';
  values[col.mime_type - 1] = row.mime_type || '';
  values[col.permissions_json - 1] = row.permissions_json || '[]';
  values[col.approved - 1] = row.approved || '';
  values[col.editors - 1] = row.editors || '';
  values[col.commenters - 1] = row.commenters || '';
  values[col.readers - 1] = row.readers || '';
  values[col.trash_parent_id - 1] = row.trash_parent_id || '';
  return sanitizeRowValues_(values);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array<Object>} rows
 */
function appendCatalogDataRowsBatch_(sheet, col, rows) {
  if (!rows || !rows.length) {
    return;
  }
  var values = [];
  for (var i = 0; i < rows.length; i++) {
    values.push(catalogDataRowValuesFromObject_(col, rows[i]));
  }
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, SHEET_HEADERS.CatalogData.length).setValues(values);
}

/**
 * @param {Array} values
 * @returns {Array}
 */
function sanitizeRowValues_(values) {
  for (var i = 0; i < values.length; i++) {
    if (values[i] === undefined || values[i] === null) {
      values[i] = '';
    }
  }
  return values;
}

/**
 * @param {Array} values
 * @param {Object.<string, number>} col
 * @returns {{name: string, type: string, folder_code: string}}
 */
function catalogRowFromValues_(values, col) {
  return {
    name: readCell_(values[col.name - 1]),
    type: readCell_(values[col.type - 1]),
    folder_code: readCell_(values[col.folder_code - 1])
  };
}

/**
 * Перерисовка листа Catalog из CatalogData (запись + группировки, виден корень).
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function renderCatalogView() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return {
        ok: false,
        error: 'NO_SPREADSHEET',
        message: 'Нет активной таблицы.'
      };
    }

    var sheets = getCatalogSheets_(ss);
    var dataHeaderCheck = validateCatalogDataHeaders_(sheets.data);
    if (!dataHeaderCheck.ok) {
      return dataHeaderCheck;
    }

    var data = renderCatalogViewCore_(sheets.view, sheets.data);
    return {
      ok: true,
      data: data,
      message: 'Дерево отрисовано: ' + data.rows + ' строк, ' + data.groups + ' групп.'
    };
  } catch (err) {
    return {
      ok: false,
      error: 'RENDER_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {{skipPermissionSync?: boolean, skipDriveSizeScan?: boolean, skipDriveLinks?: boolean}=} options
 * @returns {{rows: number, groups: number}}
 */
function renderCatalogViewCore_(viewSheet, dataSheet, options) {
  options = options || {};
  configureCatalogViewSheet_(viewSheet);
  var ss = dataSheet.getParent();
  var indexSheet = getOrEnsureCatalogIndexSheet_(ss);
  configureCatalogIndexSheet_(indexSheet);
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var viewCol = getCatalogViewColumnMap_();
  var tree = buildCatalogTreeState_(dataSheet, dataCol);
  if (!tree) {
    wipeCatalogViewRows_(viewSheet);
    wipeCatalogIndexRows_(indexSheet);
    return { rows: 0, groups: 0 };
  }

  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var allNodes = readCatalogDataNodes_(dataSheet, dataCol);
  if (!options.skipPermissionSync) {
    ensureCatalogPermissionDisplay_(dataSheet, dataCol, allNodes, usersMap);
    tree = buildCatalogTreeState_(dataSheet, dataCol);
    allNodes = readCatalogDataNodes_(dataSheet, dataCol);
  }

  var sorted = tree.sorted;
  var visible = prepareVisibleCatalogNodes_(sorted, tree.nodesById);
  var sizeContext = options.skipDriveSizeScan
    ? { sizeByFileId: {}, folderSizeById: {} }
    : buildCatalogViewSizeContext_(ss, tree, allNodes);
  var viewRows = [];
  var indexRows = [];
  for (var v = 0; v < visible.length; v++) {
    viewRows.push(buildCatalogViewRow_(visible[v], viewCol, usersMap, sizeContext));
    indexRows.push(buildCatalogIndexRow_(visible[v]));
  }

  writeCatalogViewRows_(viewSheet, viewRows);
  writeCatalogIndexRows_(indexSheet, indexRows);

  clearCatalogRowGroups_(viewSheet);
  var groups = createCatalogViewRowGroups_(viewSheet, visible);
  resetCatalogViewTreeFormatting_(viewSheet, visible.length);
  applyCatalogViewFolderBold_(viewSheet, visible);
  applyCatalogViewApprovalCells_(viewSheet, visible, viewCol);
  applyCatalogViewAlternatingRows_(viewSheet, visible.length);
  applyCatalogViewSizeColumnFormat_(viewSheet, viewCol, visible.length);
  if (options.skipDriveLinks) {
    applyCatalogViewTreePlainNameCells_(viewSheet, visible);
  } else {
    applyCatalogViewTreeNameCells_(viewSheet, visible);
  }

  return { rows: visible.length, groups: groups };
}

/**
 * Быстрая перерисовка после смены parent_id (корзина и т.п.) — без Drive и синхронизации прав.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @returns {{rows: number, groups: number}}
 */
function renderCatalogViewLight_(viewSheet, dataSheet) {
  return renderCatalogViewCore_(viewSheet, dataSheet, {
    skipPermissionSync: true,
    skipDriveSizeScan: true
  });
}

/**
 * Убирает «0_Каталог» из отрисовки и назначает индексы строк view.
 *
 * @param {Array} sorted
 * @param {Object.<string, Object>} nodesById
 * @returns {Array}
 */
function prepareVisibleCatalogNodes_(sorted, nodesById) {
  var visible = filterVisibleCatalogNodes_(sorted);
  assignVisibleViewIndices_(visible, nodesById);
  return visible;
}

/**
 * @param {Array} sorted
 * @returns {Array}
 */
function filterVisibleCatalogNodes_(sorted) {
  var visible = [];
  for (var i = 0; i < sorted.length; i++) {
    if (!isCatalogSystemNode_(sorted[i])) {
      visible.push(sorted[i]);
    }
  }
  return visible;
}

/**
 * @param {Object} node
 * @returns {boolean}
 */
function isCatalogSystemNode_(node) {
  return readCell_(node.folder_code) === FOLDER_CODE_CATALOG;
}

/**
 * @param {Array} visibleNodes
 * @param {Object.<string, Object>} nodesById
 */
function assignVisibleViewIndices_(visibleNodes, nodesById) {
  for (var i = 0; i < visibleNodes.length; i++) {
    visibleNodes[i]._viewStartIndex = i;
    visibleNodes[i]._viewEndIndex = i;
  }

  for (var j = 0; j < visibleNodes.length; j++) {
    var node = visibleNodes[j];
    if (node.type !== 'folder') {
      continue;
    }
    var end = j;
    for (var k = j + 1; k < visibleNodes.length; k++) {
      if (isDescendantOfFolderNode_(visibleNodes[k], node.id, nodesById)) {
        end = k;
      }
    }
    node._viewEndIndex = end;
  }
}

/**
 * @param {Object} node
 * @param {string} ancestorFolderId
 * @param {Object.<string, Object>} nodesById
 * @returns {boolean}
 */
function isDescendantOfFolderNode_(node, ancestorFolderId, nodesById) {
  var current = nodesById[node.id];
  while (current && current.parent_id) {
    current = nodesById[current.parent_id];
    if (!current) {
      break;
    }
    if (current.id === ancestorFolderId) {
      return true;
    }
  }
  return false;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @returns {{sorted: Array, nodesById: Object, childrenByParent: Object, catalogRootId: string}|null}
 */
function buildCatalogTreeState_(dataSheet, dataCol) {
  var nodes = readCatalogDataNodes_(dataSheet, dataCol);
  if (!nodes.length) {
    return null;
  }

  var nodesById = {};
  var childrenByParent = {};
  var catalogRootId = findCatalogRootId_(nodes, dataCol);

  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    if (!node.id) {
      continue;
    }
    nodesById[node.id] = node;
    var parentKey = node.parent_id || '';
    if (!childrenByParent[parentKey]) {
      childrenByParent[parentKey] = [];
    }
    childrenByParent[parentKey].push(node.id);
  }

  if (!catalogRootId) {
    seedSystemFolders_(dataSheet);
    nodes = readCatalogDataNodes_(dataSheet, dataCol);
    nodesById = {};
    childrenByParent = {};
    catalogRootId = findCatalogRootId_(nodes, dataCol);
    for (var j = 0; j < nodes.length; j++) {
      var reNode = nodes[j];
      if (!reNode.id) {
        continue;
      }
      nodesById[reNode.id] = reNode;
      var reParentKey = reNode.parent_id || '';
      if (!childrenByParent[reParentKey]) {
        childrenByParent[reParentKey] = [];
      }
      childrenByParent[reParentKey].push(reNode.id);
    }
  }

  if (!catalogRootId) {
    throw new Error('Не найдена системная папка «Каталог». Выполните setupSchema.');
  }

  return {
    sorted: sortCatalogNodesDepthFirst_(catalogRootId, nodesById, childrenByParent),
    nodesById: nodesById,
    childrenByParent: childrenByParent,
    catalogRootId: catalogRootId
  };
}

/**
 * Добавляет одну новую папку в Catalog без полной перерисовки; при несовпадении — полный render.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {string} newNodeId
 * @param {string} parentId
 * @returns {boolean}
 */
function insertCatalogViewNodeAfterCreate_(viewSheet, dataSheet, newNodeId, parentId) {
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var viewCol = getCatalogViewColumnMap_();
  var indexSheet = getOrEnsureCatalogIndexSheet_(dataSheet.getParent());
  var indexCol = getCatalogIndexColumnMap_();
  var tree = buildCatalogTreeState_(dataSheet, dataCol);
  if (!tree) {
    renderCatalogViewCore_(viewSheet, dataSheet);
    return false;
  }

  var visible = prepareVisibleCatalogNodes_(tree.sorted, tree.nodesById);
  var newIndex = -1;
  for (var i = 0; i < visible.length; i++) {
    if (visible[i].id === newNodeId) {
      newIndex = i;
      break;
    }
  }
  if (newIndex < 0) {
    renderCatalogViewCore_(viewSheet, dataSheet);
    return false;
  }

  var visibleIds = visible.map(function (node) {
    return node.id;
  });
  var viewIds = readCatalogViewIds_(indexSheet, indexCol);
  if (viewIds.length !== visibleIds.length - 1) {
    renderCatalogViewCore_(viewSheet, dataSheet);
    return false;
  }

  var expectedIds = visibleIds.slice();
  expectedIds.splice(newIndex, 1);
  if (!arrayIdsMatch_(viewIds, expectedIds)) {
    renderCatalogViewCore_(viewSheet, dataSheet);
    return false;
  }

  var newNode = visible[newIndex];
  var insertRow = newIndex + 2;
  viewSheet.insertRowBefore(insertRow);
  viewSheet.getRange(insertRow, 1, 1, CATALOG_VIEW_COLUMN_KEYS.length)
    .setValues([buildCatalogViewRow_(newNode, viewCol, readUsersMapFromSpreadsheet_(dataSheet.getParent()))]);
  indexSheet.insertRowBefore(insertRow);
  indexSheet.getRange(insertRow, 1, 1, CATALOG_INDEX_HEADERS.length)
    .setValues([buildCatalogIndexRow_(newNode)]);
  resetCatalogViewTreeFormatting_(viewSheet, 1, insertRow);
  applyCatalogViewFolderBoldForRow_(viewSheet, newNode, insertRow);
  applyCatalogViewApprovalCellForRow_(viewSheet, insertRow, newNode, viewCol);
  applyCatalogViewTreeNameCellForRow_(viewSheet, newNode, insertRow);

  var siblings = tree.childrenByParent[parentId] || [];
  if (siblings.length === 1 && siblings[0] === newNodeId) {
    var parentNode = tree.nodesById[parentId];
    if (parentNode && parentNode.type === 'folder' && !isCatalogSystemNode_(parentNode)) {
      if (parentNode._viewStartIndex === undefined) {
        assignVisibleViewIndices_(visible, tree.nodesById);
      }
      var parentViewRow = parentNode._viewStartIndex + 2;
      try {
        viewSheet.getRange(parentViewRow + 1, 1, 1, 1).shiftRowGroupDepth(1);
      } catch (groupErr) {
        Logger.log('insertCatalogViewNodeAfterCreate_ group: ' + groupErr);
      }
    }
  }

  return true;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet
 * @param {Object.<string, number>} indexCol
 * @returns {string[]}
 */
function readCatalogViewIds_(indexSheet, indexCol) {
  var lastRow = indexSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return indexSheet.getRange(2, indexCol.catalog_id, lastRow - 1, 1)
    .getValues()
    .map(function (row) {
      return readCell_(row[0]);
    });
}

/**
 * @param {string[]} left
 * @param {string[]} right
 * @returns {boolean}
 */
function arrayIdsMatch_(left, right) {
  if (left.length !== right.length) {
    return false;
  }
  for (var i = 0; i < left.length; i++) {
    if (left[i] !== right[i]) {
      return false;
    }
  }
  return true;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object} node
 * @param {number} viewRow
 */
function applyCatalogViewFolderBoldForRow_(viewSheet, node, viewRow) {
  if (node.type !== 'folder') {
    return;
  }
  var depth = getCatalogItemDepth_(node);
  var col = Math.min(depth, CATALOG_TREE_DEPTH_MAX) + 1;
  viewSheet.getRange(viewRow, col).setFontWeight('bold');
}

/**
 * Записывает строки Catalog без полного удаления листа (быстрее, чем deleteRows на весь диапазон).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array[]} viewRows
 */
function writeCatalogViewRows_(viewSheet, viewRows) {
  var colCount = CATALOG_VIEW_COLUMN_KEYS.length;
  var newCount = viewRows.length;
  var lastRow = viewSheet.getLastRow();

  if (!newCount) {
    if (lastRow > 1) {
      viewSheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
      viewSheet.deleteRows(2, lastRow - 1);
    }
    return;
  }

  viewSheet.getRange(2, 1, newCount, colCount).setValues(viewRows);

  if (lastRow > newCount + 1) {
    viewSheet.deleteRows(newCount + 2, lastRow - newCount - 1);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 */
function wipeCatalogViewRows_(viewSheet) {
  writeCatalogViewRows_(viewSheet, []);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet
 * @param {Array[]} indexRows
 */
function writeCatalogIndexRows_(indexSheet, indexRows) {
  var colCount = CATALOG_INDEX_HEADERS.length;
  var newCount = indexRows.length;
  var lastRow = indexSheet.getLastRow();

  if (!newCount) {
    if (lastRow > 1) {
      indexSheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
      indexSheet.deleteRows(2, lastRow - 1);
    }
    return;
  }

  indexSheet.getRange(2, 1, newCount, colCount).setValues(indexRows);

  if (lastRow > newCount + 1) {
    indexSheet.deleteRows(newCount + 2, lastRow - newCount - 1);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet
 */
function wipeCatalogIndexRows_(indexSheet) {
  writeCatalogIndexRows_(indexSheet, []);
}

/**
 * Сбрасывает форматирование колонок дерева (в т.ч. гиперссылки файлов прошлых отрисовок).
 * setValues не очищает RichText — без этого папки могут остаться синими с подчёркиванием.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {number} rowCount
 * @param {number=} startRow
 */
function resetCatalogViewTreeFormatting_(viewSheet, rowCount, startRow) {
  if (!rowCount) {
    return;
  }
  startRow = startRow || 2;
  viewSheet.getRange(startRow, 1, rowCount, CATALOG_TREE_HEADERS.length).clearFormat();
}

/**
 * Папки в колонках level_* — жирным шрифтом.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} sortedNodes
 */
function applyCatalogViewFolderBold_(viewSheet, sortedNodes) {
  if (!sortedNodes.length) {
    return;
  }

  var treeColCount = CATALOG_TREE_HEADERS.length;
  viewSheet.getRange(2, 1, sortedNodes.length, treeColCount).setFontWeight('normal');

  for (var i = 0; i < sortedNodes.length; i++) {
    if (sortedNodes[i].type !== 'folder') {
      continue;
    }
    var depth = getCatalogItemDepth_(sortedNodes[i]);
    var col = Math.min(depth, CATALOG_TREE_DEPTH_MAX) + 1;
    viewSheet.getRange(i + 2, col).setFontWeight('bold');
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {number} rowCount
 */
function applyCatalogViewAlternatingRows_(viewSheet, rowCount) {
  if (!rowCount) {
    return;
  }
  var colCount = CATALOG_VIEW_VISIBLE_COUNT;
  for (var i = 0; i < rowCount; i++) {
    var bg = i % 2 === 0 ? CATALOG_VIEW_STRIPE_WHITE : CATALOG_VIEW_STRIPE_TEAL;
    viewSheet.getRange(i + 2, 1, 1, colCount).setBackground(bg);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object.<string, number>} viewCol
 * @param {number} rowCount
 */
function applyCatalogViewSizeColumnFormat_(viewSheet, viewCol, rowCount) {
  if (!rowCount || !viewCol.type) {
    return;
  }
  var range = viewSheet.getRange(2, viewCol.type, rowCount, 1);
  range.setHorizontalAlignment('right');
  range.setNumberFormat('0.0');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {object} tree
 * @param {Array} allNodes
 * @returns {{sizeByFileId: Object.<string, number>, folderSizeById: Object.<string, number>}}
 */
function buildCatalogViewSizeContext_(ss, tree, allNodes) {
  var sizeByFileId = {};
  var folderSizeById = {};
  if (!ss || !tree || !allNodes) {
    return { sizeByFileId: sizeByFileId, folderSizeById: folderSizeById };
  }

  try {
    var settingsSheet = ss.getSheetByName('Settings');
    var rootFolderId = settingsSheet ? getSetting_(settingsSheet, 'root_folder_id') : '';
    if (rootFolderId) {
      var skipFolders = {};
      var importFolderId = settingsSheet ? getSetting_(settingsSheet, 'import_folder_id') : '';
      if (importFolderId) {
        skipFolders[importFolderId] = true;
      }
      var driveFiles = listAllFilesInFolder_(rootFolderId, skipFolders);
      for (var f = 0; f < driveFiles.length; f++) {
        sizeByFileId[driveFiles[f].id] = driveFiles[f].size || 0;
      }
    }
  } catch (sizeErr) {
    Logger.log('buildCatalogViewSizeContext_ drive: ' + sizeErr);
  }

  var nodesById = tree.nodesById;
  for (var n = 0; n < allNodes.length; n++) {
    var folder = allNodes[n];
    if (folder.type !== 'folder' || isCatalogSystemNode_(folder)) {
      continue;
    }
    var total = 0;
    for (var s = 0; s < allNodes.length; s++) {
      var fileNode = allNodes[s];
      if (fileNode.type !== 'file' || !fileNode.file_id) {
        continue;
      }
      if (isDescendantOfFolderNode_(fileNode, folder.id, nodesById)) {
        total += sizeByFileId[fileNode.file_id] || 0;
      }
    }
    folderSizeById[folder.id] = total;
  }

  return { sizeByFileId: sizeByFileId, folderSizeById: folderSizeById };
}

/**
 * @param {Object} node
 * @param {{sizeByFileId: Object.<string, number>, folderSizeById: Object.<string, number>}=} sizeContext
 * @returns {string}
 */
function resolveCatalogViewSizeValue_(node, sizeContext) {
  if (!node) {
    return '';
  }
  if (node.type === 'file') {
    if (!node.file_id) {
      return '';
    }
    if (sizeContext && sizeContext.sizeByFileId[node.file_id] !== undefined) {
      return bytesToMegabytes_(sizeContext.sizeByFileId[node.file_id]);
    }
    try {
      var meta = Drive.Files.get(node.file_id, {
        supportsAllDrives: true,
        fields: 'size'
      });
      return bytesToMegabytes_(parseInt(meta.size, 10) || 0);
    } catch (fileErr) {
      return '';
    }
  }
  if (node.type === 'folder') {
    if (sizeContext && sizeContext.folderSizeById[node.id] !== undefined) {
      return bytesToMegabytes_(sizeContext.folderSizeById[node.id]);
    }
    return '';
  }
  return '';
}

/**
 * @param {Object} node
 * @param {Object.<string, number>} viewCol
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 * @param {{sizeByFileId: Object.<string, number>, folderSizeById: Object.<string, number>}=} sizeContext
 * @returns {Array}
 */
function buildCatalogViewRow_(node, viewCol, usersByEmail, sizeContext) {
  var depth = getCatalogItemDepth_(node);
  var treeValues = buildTreeLevelValues_(node.name || '', depth);
  var permissionDisplay = resolveNodePermissionDisplay_(node, usersByEmail);
  var values = new Array(CATALOG_VIEW_COLUMN_KEYS.length);
  values[viewCol.level_0 - 1] = treeValues[0];
  values[viewCol.level_1 - 1] = treeValues[1];
  values[viewCol.level_2 - 1] = treeValues[2];
  values[viewCol.level_3 - 1] = treeValues[3];
  values[viewCol.level_4 - 1] = treeValues[4];
  values[viewCol.level_5 - 1] = treeValues[5];
  values[viewCol.type - 1] = resolveCatalogViewSizeValue_(node, sizeContext);
  values[viewCol.approved - 1] = resolveCatalogApprovalViewValue_(node);
  values[viewCol.editors - 1] = formatCatalogPermissionCell_(permissionDisplay.editors);
  values[viewCol.commenters - 1] = formatCatalogPermissionCell_(permissionDisplay.commenters);
  values[viewCol.readers - 1] = formatCatalogPermissionCell_(permissionDisplay.readers);
  return sanitizeRowValues_(values);
}

/**
 * @param {Object} node
 * @returns {Array}
 */
function buildCatalogIndexRow_(node) {
  return sanitizeRowValues_([
    node.id || '',
    node.type || '',
    node.file_id || '',
    readCell_(node.folder_code),
    node.name || '',
    readCell_(node.approved),
    node.rowIndex || ''
  ]);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Object.<string, {email: string, name: string}>}
 */
function readUsersMapFromSpreadsheet_(ss) {
  if (!ss) {
    return {};
  }
  var usersSheet = ss.getSheetByName('Users');
  return usersSheet ? readUsersMap_(usersSheet) : {};
}

/**
 * @param {Object} node
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 * @returns {{editors: string, commenters: string, readers: string}}
 */
function resolveNodePermissionDisplay_(node, usersByEmail) {
  var editors = readCell_(node.editors);
  var commenters = readCell_(node.commenters);
  var readers = readCell_(node.readers);
  if (hasPermissionDisplayValue_(editors) ||
      hasPermissionDisplayValue_(commenters) ||
      hasPermissionDisplayValue_(readers)) {
    return {
      editors: editors,
      commenters: commenters,
      readers: readers
    };
  }
  return permissionsJsonToDisplay_(node.permissions_json, usersByEmail);
}

/**
 * @param {*} value
 * @returns {boolean}
 */
function hasPermissionDisplayValue_(value) {
  var cell = readCell_(value);
  if (!cell) {
    return false;
  }
  var text = String(cell).trim();
  return text !== '' && text !== '—' && text !== '-';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {Array} nodes
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 */
function ensureCatalogPermissionDisplay_(dataSheet, dataCol, nodes, usersByEmail) {
  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    var current = resolveNodePermissionDisplay_(node, usersByEmail);
    if (hasAnyPermissionDisplay_(current)) {
      writeCatalogPermissionDisplay_(dataSheet, dataCol, node, current);
      continue;
    }

    if (node.type !== 'file' || !node.file_id) {
      continue;
    }

    try {
      var permissions = readDrivePermissions_(node.file_id);
      var display = buildPermissionDisplay_(permissions, usersByEmail);
      var permissionsJson = JSON.stringify(permissions);
      dataSheet.getRange(node.rowIndex, dataCol.permissions_json).setValue(permissionsJson);
      writeCatalogPermissionDisplay_(dataSheet, dataCol, node, display);
      node.permissions_json = permissionsJson;
      node.editors = display.editors;
      node.commenters = display.commenters;
      node.readers = display.readers;
    } catch (err) {
      Logger.log('ensureCatalogPermissionDisplay_ ' + node.file_id + ': ' + err);
    }
  }
}

/**
 * @param {{editors: string, commenters: string, readers: string}} display
 * @returns {boolean}
 */
function hasAnyPermissionDisplay_(display) {
  return hasPermissionDisplayValue_(display.editors) ||
    hasPermissionDisplayValue_(display.commenters) ||
    hasPermissionDisplayValue_(display.readers);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {{editors: string, commenters: string, readers: string}} display
 */
function writeCatalogPermissionDisplay_(sheet, col, node, display) {
  if (!hasPermissionDisplayValue_(node.editors) && hasPermissionDisplayValue_(display.editors)) {
    sheet.getRange(node.rowIndex, col.editors).setValue(display.editors);
    node.editors = display.editors;
  }
  if (!hasPermissionDisplayValue_(node.commenters) && hasPermissionDisplayValue_(display.commenters)) {
    sheet.getRange(node.rowIndex, col.commenters).setValue(display.commenters);
    node.commenters = display.commenters;
  }
  if (!hasPermissionDisplayValue_(node.readers) && hasPermissionDisplayValue_(display.readers)) {
    sheet.getRange(node.rowIndex, col.readers).setValue(display.readers);
    node.readers = display.readers;
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} visibleNodes
 */
function applyCatalogViewTreeNameCells_(viewSheet, visibleNodes) {
  var fileIds = collectVisibleCatalogFileIds_(visibleNodes);
  var urlByFileId = fetchDriveFileOpenUrls_(fileIds);
  for (var i = 0; i < visibleNodes.length; i++) {
    applyCatalogViewTreeNameCellForRow_(viewSheet, visibleNodes[i], i + 2, urlByFileId);
  }
}

/**
 * Имена в дереве без ссылок на Drive (быстрая перерисовка).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} visibleNodes
 */
function applyCatalogViewTreePlainNameCells_(viewSheet, visibleNodes) {
  for (var i = 0; i < visibleNodes.length; i++) {
    var node = visibleNodes[i];
    if (!node) {
      continue;
    }
    var depth = getCatalogItemDepth_(node);
    var col = Math.min(depth, CATALOG_TREE_DEPTH_MAX) + 1;
    var label = String(node.name || '');
    viewSheet.getRange(i + 2, col).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(label).build()
    );
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} visibleNodes
 */
function applyCatalogViewFileLinks_(viewSheet, visibleNodes) {
  applyCatalogViewTreeNameCells_(viewSheet, visibleNodes);
}

/**
 * @param {Array} visibleNodes
 * @returns {Array<string>}
 */
function collectVisibleCatalogFileIds_(visibleNodes) {
  var ids = [];
  var seen = {};
  for (var i = 0; i < visibleNodes.length; i++) {
    var node = visibleNodes[i];
    if (!node || node.type !== 'file' || !node.file_id || seen[node.file_id]) {
      continue;
    }
    seen[node.file_id] = true;
    ids.push(node.file_id);
  }
  return ids;
}

/**
 * Имя в колонке level_*: ссылка для файлов, обычный текст для папок.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object} node
 * @param {number} viewRow
 * @param {Object.<string, string>=} urlByFileId
 */
function applyCatalogViewTreeNameCellForRow_(viewSheet, node, viewRow, urlByFileId) {
  if (!node) {
    return;
  }
  var depth = getCatalogItemDepth_(node);
  var col = Math.min(depth, CATALOG_TREE_DEPTH_MAX) + 1;
  var label = String(node.name || '');
  var cell = viewSheet.getRange(viewRow, col);
  if (node.type === 'file' && node.file_id) {
    var url = (urlByFileId && urlByFileId[node.file_id]) ||
      resolveDriveFileOpenUrl_(node.file_id, node.mime_type);
    cell.setRichTextValue(SpreadsheetApp.newRichTextValue()
      .setText(label)
      .setLinkUrl(url)
      .build());
    return;
  }
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText(label)
    .build());
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object} node
 * @param {number} viewRow
 * @param {Object.<string, string>=} urlByFileId
 */
function applyCatalogViewFileLinkForRow_(viewSheet, node, viewRow, urlByFileId) {
  applyCatalogViewTreeNameCellForRow_(viewSheet, node, viewRow, urlByFileId);
}

/**
 * @param {Array<string>} fileIds
 * @returns {Object.<string, string>}
 */
function fetchDriveFileOpenUrls_(fileIds) {
  var map = {};
  for (var i = 0; i < fileIds.length; i++) {
    var fileId = fileIds[i];
    map[fileId] = resolveDriveFileOpenUrl_(fileId, '');
  }
  return map;
}

/**
 * Канонический URL открытия из Drive API (webViewLink), с нормализацией /view → /edit.
 * @param {string} fileId
 * @param {string} mimeType
 * @param {number=} depth
 * @returns {string}
 */
function resolveDriveFileOpenUrl_(fileId, mimeType, depth) {
  depth = depth || 0;
  if (depth > 2) {
    return getDriveFileOpenUrlFallback_(fileId, mimeType);
  }
  mimeType = String(mimeType || '');
  if (mimeType && mimeType !== DRIVE_SHORTCUT_MIME) {
    return getDriveFileOpenUrlFallback_(fileId, mimeType);
  }
  try {
    var file = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,mimeType,webViewLink,shortcutDetails(targetId,targetMimeType)'
    });
    var targetMime = file.mimeType || mimeType;
    if (file.mimeType === DRIVE_SHORTCUT_MIME && file.shortcutDetails && file.shortcutDetails.targetId) {
      return resolveDriveFileOpenUrl_(
        file.shortcutDetails.targetId,
        file.shortcutDetails.targetMimeType || targetMime,
        depth + 1
      );
    }
    if (file.webViewLink) {
      return normalizeDriveOpenUrl_(file.webViewLink, targetMime);
    }
    mimeType = targetMime;
  } catch (err) {
    if (!isDriveFileNotFoundOrForbiddenError_(err)) {
      Logger.log('resolveDriveFileOpenUrl_ ' + fileId + ': ' + err);
    }
  }
  return getDriveFileOpenUrlFallback_(fileId, mimeType);
}

/**
 * @param {*} err
 * @returns {boolean}
 */
function isDriveFileNotFoundOrForbiddenError_(err) {
  var text = String(err || '').toLowerCase();
  return text.indexOf('not found') !== -1 ||
    text.indexOf('404') !== -1 ||
    text.indexOf('403') !== -1 ||
    text.indexOf('forbidden') !== -1 ||
    text.indexOf('insufficient') !== -1;
}

/**
 * @param {string} url
 * @param {string} mimeType
 * @returns {string}
 */
function normalizeDriveOpenUrl_(url, mimeType) {
  var str = String(url || '').trim();
  if (!str) {
    return str;
  }
  mimeType = String(mimeType || '');

  if (mimeType.indexOf('application/vnd.google-apps.') === 0 &&
      mimeType !== DRIVE_FOLDER_MIME &&
      mimeType !== DRIVE_SHORTCUT_MIME) {
    if (str.indexOf('docs.google.com') >= 0) {
      return str.replace(/\/view(\?|#|$)/, '/edit$1');
    }
    return str;
  }

  if (isDriveOfficeOrBinaryMime_(mimeType) || str.indexOf('drive.google.com/file/d/') >= 0) {
    return str.replace(/\/view(\?|#|$)/, '/edit$1');
  }

  return str;
}

/**
 * @param {string} mimeType
 * @returns {boolean}
 */
function isDriveOfficeOrBinaryMime_(mimeType) {
  return mimeType === 'application/pdf' ||
    mimeType === 'application/msword' ||
    mimeType === 'application/vnd.ms-excel' ||
    mimeType === 'application/vnd.ms-powerpoint' ||
    mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
}

/**
 * Запасной URL, если Drive API недоступен.
 * @param {string} fileId
 * @param {string} mimeType
 * @returns {string}
 */
function getDriveFileOpenUrlFallback_(fileId, mimeType) {
  mimeType = String(mimeType || '');
  if (mimeType === 'application/vnd.google-apps.document') {
    return 'https://docs.google.com/document/d/' + fileId + '/edit';
  }
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    return 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit';
  }
  if (mimeType === 'application/vnd.google-apps.presentation') {
    return 'https://docs.google.com/presentation/d/' + fileId + '/edit';
  }
  if (mimeType === 'application/vnd.google-apps.form') {
    return 'https://docs.google.com/forms/d/' + fileId + '/edit';
  }
  if (mimeType === 'application/vnd.google-apps.drawing') {
    return 'https://docs.google.com/drawings/d/' + fileId + '/edit';
  }
  return 'https://drive.google.com/file/d/' + fileId + '/edit';
}

/**
 * @param {string} permissionsJson
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 * @returns {{editors: string, commenters: string, readers: string}}
 */
function permissionsJsonToDisplay_(permissionsJson, usersByEmail) {
  var permissions = parseCatalogPermissionsJson_(permissionsJson);
  if (!permissions.length) {
    return { editors: '', commenters: '', readers: '' };
  }
  return buildPermissionDisplay_(permissions, usersByEmail || {});
}

/**
 * @param {string} permissionsJson
 * @returns {Array}
 */
function parseCatalogPermissionsJson_(permissionsJson) {
  var raw = readCell_(permissionsJson);
  if (!raw) {
    return [];
  }
  var parsed = [];
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    return [];
  }
  if (!parsed || !parsed.length) {
    return [];
  }
  return normalizeCatalogPermissions_(parsed);
}

/**
 * @param {Array} permissions
 * @returns {Array}
 */
function normalizeCatalogPermissions_(permissions) {
  var result = [];
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (!perm) {
      continue;
    }
    var level = perm.level || mapDriveRoleToLevel_(perm.role);
    if (!level) {
      continue;
    }
    var type = perm.type || 'user';
    var subject = perm.subject || perm.emailAddress || perm.email || '';
    if (type === 'domain' && !subject) {
      subject = perm.domain || '';
    }
    if (type === 'anyone' && !subject) {
      subject = 'anyone';
    }
    if (!subject) {
      continue;
    }
    result.push({
      subject: subject,
      type: type,
      level: level,
      display_name: perm.display_name || perm.displayName || subject
    });
  }
  return result;
}

/**
 * @param {*} value
 * @returns {string}
 */
function formatCatalogPermissionCell_(value) {
  var cell = readCell_(value);
  return cell ? String(cell) : '—';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} sortedNodes
 * @returns {number}
 */
function createCatalogViewRowGroups_(viewSheet, sortedNodes) {
  var candidates = [];
  for (var i = 0; i < sortedNodes.length; i++) {
    var node = sortedNodes[i];
    var startIndex = node._viewStartIndex;
    var endIndex = node._viewEndIndex;
    if (startIndex === undefined || endIndex === undefined) {
      continue;
    }
    if (node.type !== 'folder' || endIndex <= startIndex) {
      continue;
    }
    candidates.push({
      startRow: startIndex + 3,
      numRows: endIndex - startIndex,
      span: endIndex - startIndex
    });
  }

  candidates.sort(function (a, b) {
    return a.span - b.span;
  });

  var count = 0;
  for (var g = 0; g < candidates.length; g++) {
    var group = candidates[g];
    viewSheet.getRange(group.startRow, 1, group.numRows, 1).shiftRowGroupDepth(1);
    count++;
  }
  return count;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @returns {Array}
 */
function readCatalogDataNodes_(sheet, col) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.CatalogData.length).getValues();
  var nodes = [];
  for (var i = 0; i < rows.length; i++) {
    var values = rows[i];
    nodes.push(catalogDataNodeFromValues_(values, col, i + 2));
  }
  return nodes;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {number} rowIndex
 * @returns {Object|null}
 */
function readCatalogDataNodeAtRow_(sheet, col, rowIndex) {
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return null;
  }
  var values = sheet.getRange(rowIndex, 1, 1, SHEET_HEADERS.CatalogData.length).getValues()[0];
  var node = catalogDataNodeFromValues_(values, col, rowIndex);
  if (!node.id || !node.type) {
    return null;
  }
  return node;
}

/**
 * @param {Array} values
 * @param {Object.<string, number>} col
 * @param {number} rowIndex
 * @returns {Object}
 */
function catalogDataNodeFromValues_(values, col, rowIndex) {
  return {
    rowIndex: rowIndex,
    id: readCell_(values[col.id - 1]),
    parent_id: readCell_(values[col.parent_id - 1]),
    folder_code: readCell_(values[col.folder_code - 1]),
    name: readCell_(values[col.name - 1]),
    type: readCell_(values[col.type - 1]),
    file_id: readCell_(values[col.file_id - 1]),
    mime_type: readCell_(values[col.mime_type - 1]),
    permissions_json: readCell_(values[col.permissions_json - 1]) || '[]',
    approved: readCell_(values[col.approved - 1]),
    editors: readCell_(values[col.editors - 1]),
    commenters: readCell_(values[col.commenters - 1]),
    readers: readCell_(values[col.readers - 1]),
    trash_parent_id: readCell_(values[col.trash_parent_id - 1]),
    values: values
  };
}

/**
 * @param {Object} row
 * @returns {number}
 */
function getCatalogItemDepth_(row) {
  var folderCode = String(row.folder_code || '');

  if (row.type === 'folder' && folderCode === FOLDER_CODE_CATALOG) {
    return 0;
  }

  if (row.type === 'folder') {
    return getFolderDisplayDepth_(folderCode);
  }

  if (!folderCode) {
    return 0;
  }

  return getFolderDisplayDepth_(folderCode) + 1;
}

/**
 * Глубина папки для колонок level_*:
 * 10, 11, 99 → 0; 101 → 1; 1012 → 2
 *
 * @param {string} folderCode
 * @returns {number}
 */
function getFolderDisplayDepth_(folderCode) {
  folderCode = readCell_(folderCode);
  if (!folderCode || folderCode === FOLDER_CODE_CATALOG || folderCode === FOLDER_CODE_TRASH) {
    return 0;
  }
  if (isRootLevelFolderCode_(folderCode)) {
    return 0;
  }
  return folderCode.length - 2;
}

/**
 * Корневая папка под «Каталогом»: код 10–98 (две цифры, без 01–09 и 99).
 *
 * @param {string} folderCode
 * @returns {boolean}
 */
function isRootLevelFolderCode_(folderCode) {
  folderCode = readCell_(folderCode);
  if (!/^\d{2}$/.test(folderCode)) {
    return false;
  }
  var n = parseInt(folderCode, 10);
  return n >= FOLDER_CODE_ROOT_MIN && n <= FOLDER_CODE_ROOT_MAX;
}

/**
 * @param {string} name
 * @param {number} depth
 * @returns {string[]}
 */
function buildTreeLevelValues_(name, depth) {
  var levels = ['', '', '', '', '', ''];
  var safeDepth = Math.max(0, Math.min(depth, CATALOG_TREE_DEPTH_MAX));
  levels[safeDepth] = name;
  return levels;
}

function renderCatalogViewMenu_() {
  assertCatalogImportNotBusy_();
  var result = renderCatalogView();
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      var title = result.ok ? 'Обновление дерева' : 'Ошибка';
      ui.alert(title, result.message || result.error, ui.ButtonSet.OK);
    }
  } catch (err) {
    Logger.log('renderCatalogView: ' + JSON.stringify(result));
  }
}

/**
 * @param {string} rootId
 * @param {Object} nodesById
 * @param {Object.<string, string[]>} childrenByParent
 * @returns {Array}
 */
function sortCatalogNodesDepthFirst_(rootId, nodesById, childrenByParent) {
  var result = [];
  var visited = {};

  function walk(id) {
    if (!id || visited[id] || !nodesById[id]) {
      return;
    }
    visited[id] = true;
    var node = nodesById[id];
    node._startIndex = result.length;
    result.push(node);

    var childIds = (childrenByParent[id] || []).slice();
    childIds.sort(compareCatalogNodeIds_(nodesById));
    for (var i = 0; i < childIds.length; i++) {
      walk(childIds[i]);
    }
    node._endIndex = result.length - 1;
  }

  walk(rootId);

  var remainingIds = Object.keys(nodesById).filter(function (id) {
    return !visited[id];
  });
  remainingIds.sort(compareCatalogNodeIds_(nodesById));
  for (var r = 0; r < remainingIds.length; r++) {
    walk(remainingIds[r]);
  }

  return result;
}

/**
 * @param {Object} nodesById
 * @returns {function(string, string): number}
 */
function compareCatalogNodeIds_(nodesById) {
  return function (idA, idB) {
    return compareCatalogNodes_(nodesById[idA], nodesById[idB]);
  };
}

/**
 * @param {{folder_code: string, type: string, name: string}} a
 * @param {{folder_code: string, type: string, name: string}} b
 * @returns {number}
 */
function compareCatalogNodes_(a, b) {
  if (!a || !b) {
    return 0;
  }
  if (a.folder_code === FOLDER_CODE_TRASH) {
    return 1;
  }
  if (b.folder_code === FOLDER_CODE_TRASH) {
    return -1;
  }
  if (a.type === 'folder' && b.type !== 'folder') {
    return -1;
  }
  if (b.type === 'folder' && a.type !== 'folder') {
    return 1;
  }
  var codeA = a.folder_code || '';
  var codeB = b.folder_code || '';
  if (codeA !== codeB) {
    return codeA < codeB ? -1 : 1;
  }
  return String(a.name || '').localeCompare(String(b.name || ''));
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearCatalogRowGroups_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow < 3) {
    return;
  }
  var range = sheet.getRange(2, 1, lastRow - 1, 1);
  for (var i = 0; i < 8; i++) {
    try {
      range.shiftRowGroupDepth(-1);
    } catch (err) {
      break;
    }
  }
}

/**
 * @param {string[]} headers
 * @returns {Object.<string, number>}
 */
function columnIndexMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i + 1;
  }
  return map;
}

/**
 * @returns {Object.<string, number>}
 */
function getCatalogViewColumnMap_() {
  return columnIndexMap_(CATALOG_VIEW_COLUMN_KEYS);
}

/**
 * @returns {Object.<string, number>}
 */
function getCatalogIndexColumnMap_() {
  return columnIndexMap_(CATALOG_INDEX_HEADERS);
}

/**
 * @param {Array} values
 * @param {Object.<string, number>} col
 * @param {number} viewRowIndex
 * @returns {Object|null}
 */
function catalogIndexEntryFromValues_(values, col, viewRowIndex) {
  var catalogId = readCell_(values[col.catalog_id - 1]);
  if (!catalogId) {
    return null;
  }
  return {
    viewRowIndex: viewRowIndex,
    catalog_id: catalogId,
    type: readCell_(values[col.type - 1]),
    file_id: readCell_(values[col.file_id - 1]),
    folder_code: readCell_(values[col.folder_code - 1]),
    name: readCell_(values[col.name - 1]),
    approved: readCell_(values[col.approved - 1]),
    data_row_index: parseInt(readCell_(values[col.data_row_index - 1]), 10) || 0
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet
 * @param {Object.<string, number>} indexCol
 * @param {number} viewRowIndex
 * @returns {Object|null}
 */
function readCatalogIndexEntryAtViewRow_(indexSheet, indexCol, viewRowIndex) {
  if (!indexSheet || viewRowIndex < 2 || viewRowIndex > indexSheet.getLastRow()) {
    return null;
  }
  var values = indexSheet.getRange(viewRowIndex, 1, 1, CATALOG_INDEX_HEADERS.length).getValues()[0];
  return catalogIndexEntryFromValues_(values, indexCol, viewRowIndex);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} columnIndex
 * @returns {string[]}
 */
function readColumnValues_(sheet, columnIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(2, columnIndex, lastRow - 1, 1)
    .getValues()
    .map(function (row) { return readCell_(row[0]); });
}

/**
 * Читает значение ячейки; сохраняет 0 и false (в отличие от value || '').
 *
 * @param {*} value
 * @returns {string}
 */
function readCell_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

/**
 * @param {string[]} codes
 * @param {string} code
 * @returns {boolean}
 */
function hasFolderCode_(codes, code) {
  var target = readCell_(code);
  for (var i = 0; i < codes.length; i++) {
    if (readCell_(codes[i]) === target) {
      return true;
    }
  }
  return false;
}

/**
 * @param {Array} nodes
 * @param {Object.<string, number>} col
 * @returns {string}
 */
function findCatalogRootId_(nodes, col) {
  var catalogName = FOLDER_CODE_CATALOG + '_Каталог';
  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    if (!node.id) {
      continue;
    }
    if (node.folder_code === FOLDER_CODE_CATALOG) {
      return node.id;
    }
  }
  for (var j = 0; j < nodes.length; j++) {
    var fallback = nodes[j];
    if (!fallback.id || fallback.type !== 'folder') {
      continue;
    }
    if (fallback.name === catalogName) {
      fallback.folder_code = FOLDER_CODE_CATALOG;
      fallback.values[col.folder_code - 1] = FOLDER_CODE_CATALOG;
      return fallback.id;
    }
  }
  return '';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {string} folderCode
 * @returns {string}
 */
function findIdByFolderCode_(sheet, col, folderCode) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return '';
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.CatalogData.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (readCell_(rows[i][col.folder_code - 1]) === readCell_(folderCode)) {
      return readCell_(rows[i][col.id - 1]);
    }
  }
  return '';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} key
 * @param {string} value
 */
function upsertSetting_(sheet, key, value) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.appendRow([key, value]);
    return;
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * @returns {string}
 */
function generateId_() {
  return Utilities.getUuid();
}

var OWNER_TRANSFER_PENDING_KEY = 'OWNER_TRANSFER_PENDING';
var IMPORT_OWNER_BATCH_KEY = 'IMPORT_OWNER_BATCH_PENDING';
var PENDING_IMPORT_DRIVE_SYNC_KEY = 'pending_import_drive_sync_queue';
var IMPORT_DRIVE_SYNC_HANDLER = 'processPendingImportDriveSync_';
var IMPORT_DRIVE_BATCH_CHUNK_SIZE = 10;
var IMPORT_BATCH_PROGRESS_KEY = 'import_batch_progress';
var DRIVE_BATCH_URL = 'https://www.googleapis.com/batch/drive/v3';
var DRIVE_BATCH_MAX = 50;
var PENDING_DRIVE_IMPORT_JOB_KEY = 'pending_drive_import_job';
var IMPORT_IN_PROGRESS_KEY = 'import_in_progress';
var IMPORT_JOB_META_KEY = 'import_job_meta';
var IMPORT_JOB_HANDLER = 'processPendingDriveImportJob_';
var IMPORT_VIEW_DEFER_THRESHOLD = 10;
var IMPORT_WEBAPP_SECRET_KEY = 'IMPORT_WEBAPP_SECRET';
var DRIVE_FOLDER_MIME = 'application/vnd.google-apps.folder';
var DRIVE_SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';

/**
 * Первое сканирование: файлы из root_folder_id → дочерние «Каталога»,
 * права с Drive, пользователи → Users, предложение смены владельца.
 * Запуск: меню «Каталогизатор → Первое сканирование» или из редактора Apps Script.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function firstScan() {
  try {
    var result = firstScanCore_();
    showFirstScanResult_(result);
    if (result.ok && result.data && result.data.needs_owner_transfer.length > 0) {
      PropertiesService.getScriptProperties().setProperty(
        OWNER_TRANSFER_PENDING_KEY,
        JSON.stringify(result.data.needs_owner_transfer)
      );
      showOwnerTransferDialog_();
    }
    return result;
  } catch (err) {
    var fail = {
      ok: false,
      error: 'FIRST_SCAN_FAILED',
      message: String(err)
    };
    showFirstScanResult_(fail);
    return fail;
  }
}

/**
 * @returns {Array<{file_id: string, name: string, owner_email: string}>}
 */
function getOwnerTransferCandidates() {
  var raw = PropertiesService.getScriptProperties().getProperty(OWNER_TRANSFER_PENDING_KEY);
  if (!raw) {
    return [];
  }
  return JSON.parse(raw);
}

/**
 * @param {string[]} fileIds
 * @returns {{ok: boolean, data?: object, message?: string, error?: string}}
 */
function transferOwnershipFromFirstScan(fileIds) {
  try {
    if (!fileIds || !fileIds.length) {
      return { ok: false, error: 'NO_FILES', message: 'Не выбраны файлы.' };
    }
    var email = Session.getEffectiveUser().getEmail();
    if (!email) {
      return {
        ok: false,
        error: 'NO_EMAIL',
        message: 'Не удалось определить email. Запустите из таблицы под своим аккаунтом.'
      };
    }

    var transferred = [];
    var failed = [];
    for (var i = 0; i < fileIds.length; i++) {
      try {
        Drive.Permissions.create({
          type: 'user',
          role: 'owner',
          emailAddress: email
        }, fileIds[i], {
          transferOwnership: true,
          supportsAllDrives: true
        });
        transferred.push(fileIds[i]);
      } catch (err) {
        failed.push({ file_id: fileIds[i], message: String(err) });
      }
    }

    if (transferred.length) {
      var pending = getOwnerTransferCandidates().filter(function (item) {
        return transferred.indexOf(item.file_id) === -1;
      });
      PropertiesService.getScriptProperties().setProperty(
        OWNER_TRANSFER_PENDING_KEY,
        JSON.stringify(pending)
      );
    }

    return {
      ok: failed.length === 0,
      data: { transferred: transferred, failed: failed },
      message: 'Передано файлов: ' + transferred.length +
        (failed.length ? '. Ошибок: ' + failed.length : '.')
    };
  } catch (err) {
    return {
      ok: false,
      error: 'TRANSFER_FAILED',
      message: String(err)
    };
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var serviceMenu = ui.createMenu('Служебные функции')
    .addItem('🧹 Очистить корзину', 'emptyTrashCatalogMenu_')
    .addSeparator()
    .addItem('🔍 Первое сканирование', 'firstScan')
    .addItem('🔎 Проверка целостности', 'runCatalogIntegrityCheck')
    .addItem('🔧 Исправление по отчёту', 'runCatalogIntegrityRepairMenu_')
    .addItem('🔄 Обновить дерево', 'renderCatalogViewMenu_')
    .addItem('⏹ Сбросить зависший импорт', 'resetStuckCatalogImportMenu_');

  ui.createMenu('📁 Каталогизатор')
    .addItem('📂 Создать папку', 'createFolderDialog_')
    .addItem('📄 Создать файл', 'createFileDialog_')
    .addItem('📥 Импорт', 'importDialog_')
    .addSeparator()
    .addItem('↕️ Переместить / Копировать', 'moveCopyDialog_')
    .addItem('✏️ Переименовать', 'renameDialog_')
    .addItem('🗑️ В корзину', 'deleteCatalogItemsMenu_')
    .addItem('↩️ Восстановить из корзины', 'restoreCatalogItemsMenu_')
    .addSeparator()
    .addItem('✅ Утвердить файл', 'approveCatalogFileMenu_')
    .addItem('🔐 Назначить права', 'assignPermissionsDialog_')
    .addItem('👥 Группы', 'manageCatalogGroupsDialog_')
    .addSeparator()
    .addSubMenu(serviceMenu)
    .addToUi();
  removeCatalogApprovalEditTriggers_();
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {string}
 */
function getSpreadsheetOwnerEmail_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return '';
  }
  try {
    var owner = ss.getOwner();
    if (owner && owner.getEmail) {
      return normalizeEmail_(owner.getEmail());
    }
  } catch (err) {
    Logger.log('getSpreadsheetOwnerEmail_ getOwner: ' + err);
  }
  try {
    var fileOwner = DriveApp.getFileById(ss.getId()).getOwner();
    if (fileOwner && fileOwner.getEmail) {
      return normalizeEmail_(fileOwner.getEmail());
    }
  } catch (driveErr) {
    Logger.log('getSpreadsheetOwnerEmail_ DriveApp: ' + driveErr);
  }
  try {
    var driveFile = Drive.Files.get(ss.getId(), {
      supportsAllDrives: true,
      fields: 'owners(emailAddress)'
    });
    if (driveFile.owners && driveFile.owners.length) {
      return normalizeEmail_(driveFile.owners[0].emailAddress || '');
    }
  } catch (apiErr) {
    Logger.log('getSpreadsheetOwnerEmail_ Drive API: ' + apiErr);
  }
  return '';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {boolean}
 */
function isSpreadsheetOwner_(ss) {
  var ownerEmail = getSpreadsheetOwnerEmail_(ss);
  var userEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  return !!(ownerEmail && userEmail && ownerEmail === userEmail);
}

/**
 * Открывает диалог назначения прав (только владелец таблицы).
 */
function assignPermissionsDialog_() {
  var contextResult = getAssignPermissionsContext();
  var template = HtmlService.createTemplateFromFile('AssignPermissionsDialog');
  template.initialContext = JSON.stringify(contextResult);
  var html = template.evaluate()
    .setWidth(760)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Назначить права');
}

/**
 * Контекст для диалога назначения прав.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getAssignPermissionsContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { ok: false, error: 'NO_SPREADSHEET', message: 'Нет активной таблицы.' };
    }
    if (!isSpreadsheetOwner_(ss)) {
      return {
        ok: false,
        error: 'NOT_OWNER',
        message: 'Назначать права может только владелец таблицы.'
      };
    }

    var sheets = getCatalogSheets_(ss);
    var headerCheck = validateCatalogDataHeaders_(sheets.data);
    if (!headerCheck.ok) {
      return headerCheck;
    }

    if (!ss.getActiveSheet() || ss.getActiveSheet().getName() !== 'Catalog') {
      return {
        ok: false,
        error: 'WRONG_SHEET',
        message: 'Выберите строки на листе Catalog.'
      };
    }

    var selected = resolveSelectedCatalogEntries_(ss);
    if (!selected.length) {
      return {
        ok: false,
        error: 'NO_SELECTION',
        message: 'Выделите папки или файлы на листе Catalog.'
      };
    }

    var assignable = selected.filter(function (entry) {
      return !isProtectedCatalogNode_(entry);
    });
    if (!assignable.length) {
      return {
        ok: false,
        error: 'PROTECTED_ONLY',
        message: 'Нельзя назначать права для «Каталога» и «Корзины».'
      };
    }

    var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
    var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
    var plan = buildAssignPermissionsPlan_(assignable, allNodes, true);
    var currentState = resolveSharedPermissionsState_(plan.roots);
    var usersSheet = ss.getSheetByName('Users');
    var groupsSheet = ss.getSheetByName('Groups');

    var users = readActiveUsersList_(usersSheet);
    var groups = readCatalogGroupsList_(groupsSheet);

    return {
      ok: true,
      data: {
        summary_message: buildAssignPermissionsSummaryMessage_(plan.folderCount, plan.fileCount),
        preview: plan.preview,
        folder_count: plan.folderCount,
        file_count: plan.fileCount,
        mixed_permissions: currentState.mixed,
        subjects: buildAssignPermissionSubjects_(users, groups, currentState),
        has_approved_files: planHasApprovedFiles_(plan.affected),
        apply_to_subtree_default: plan.hasFolders,
        has_folders: plan.hasFolders
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'ASSIGN_PERMISSIONS_CONTEXT_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {Array} selectedEntries
 * @param {Array} allNodes
 * @param {boolean} applyToSubtree
 * @returns {{roots: Array, affected: Array, folderCount: number, fileCount: number, hasFolders: boolean, preview: string[]}}
 */
function buildAssignPermissionsPlan_(selectedEntries, allNodes, applyToSubtree) {
  var nodesById = mapCatalogNodesById_(allNodes);
  var childrenByParent = mapCatalogChildrenByParent_(allNodes);
  var selectedIds = {};
  for (var i = 0; i < selectedEntries.length; i++) {
    selectedIds[selectedEntries[i].id] = true;
  }

  var roots = [];
  for (var s = 0; s < selectedEntries.length; s++) {
    var entry = selectedEntries[s];
    if (!isDescendantOfSelectedNode_(entry.id, selectedIds, nodesById)) {
      roots.push(entry);
    }
  }

  var affectedIds = {};
  var hasFolders = false;
  for (var r = 0; r < roots.length; r++) {
    var root = roots[r];
    affectedIds[root.id] = true;
    if (root.type === 'folder') {
      hasFolders = true;
      if (applyToSubtree) {
        collectDescendantNodeIds_(root.id, childrenByParent, affectedIds);
      }
    }
  }

  var affected = [];
  for (var a = 0; a < allNodes.length; a++) {
    if (affectedIds[allNodes[a].id]) {
      affected.push(allNodes[a]);
    }
  }
  affected.sort(compareCatalogNodes_);

  var folderCount = 0;
  var fileCount = 0;
  for (var f = 0; f < affected.length; f++) {
    if (affected[f].type === 'folder') {
      folderCount++;
    } else {
      fileCount++;
    }
  }

  var preview = affected.slice(0, 12).map(function (node) {
    return '• ' + (node.name || node.type || node.id);
  });
  if (affected.length > 12) {
    preview.push('… и ещё ' + (affected.length - 12));
  }

  return {
    roots: roots,
    affected: affected,
    folderCount: folderCount,
    fileCount: fileCount,
    hasFolders: hasFolders,
    preview: preview
  };
}

/**
 * @param {number} folderCount
 * @param {number} fileCount
 * @returns {string}
 */
function buildAssignPermissionsSummaryMessage_(folderCount, fileCount) {
  return 'Будет обновлено: папок — ' + folderCount + ', файлов — ' + fileCount + '.';
}

/**
 * @param {Array} roots
 * @returns {{mixed: boolean, permissions: {editors: Array, commenters: Array, readers: Array}}}
 */
function resolveSharedPermissionsState_(roots) {
  var empty = { editors: [], commenters: [], readers: [] };
  if (!roots.length) {
    return { mixed: false, permissions: empty };
  }

  var canonical = canonicalizePermissionsJson_(roots[0].permissions_json);
  for (var i = 1; i < roots.length; i++) {
    if (canonicalizePermissionsJson_(roots[i].permissions_json) !== canonical) {
      return { mixed: true, permissions: empty };
    }
  }

  var parsed = normalizeCatalogPermissions_(parseCatalogPermissionsJson_(roots[0].permissions_json));
  return {
    mixed: false,
    permissions: splitPermissionsByLevel_(parsed)
  };
}

/**
 * @param {string} permissionsJson
 * @returns {string}
 */
function canonicalizePermissionsJson_(permissionsJson) {
  var permissions = normalizeCatalogPermissions_(parseCatalogPermissionsJson_(permissionsJson));
  permissions.sort(function (a, b) {
    var left = (a.type || '') + ':' + (a.subject || '') + ':' + (a.level || '');
    var right = (b.type || '') + ':' + (b.subject || '') + ':' + (b.level || '');
    return left.localeCompare(right);
  });
  return JSON.stringify(permissions);
}

/**
 * @param {Array} permissions
 * @returns {{editors: Array, commenters: Array, readers: Array}}
 */
function splitPermissionsByLevel_(permissions) {
  var buckets = { editors: [], commenters: [], readers: [] };
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.level === 'edit') {
      buckets.editors.push(perm);
    } else if (perm.level === 'comment') {
      buckets.commenters.push(perm);
    } else if (perm.level === 'read') {
      buckets.readers.push(perm);
    }
  }
  return buckets;
}

/**
 * Список пользователей и групп с текущей ролью для диалога назначения прав.
 *
 * @param {Array<{email: string, name: string}>} users
 * @param {Array<{group_id: string, group_name: string}>} groups
 * @param {{mixed: boolean, permissions: {editors: Array, commenters: Array, readers: Array}}} currentState
 * @returns {Array<{subject: string, type: string, label: string, level: string}>}
 */
function buildAssignPermissionSubjects_(users, groups, currentState) {
  var roleByKey = {};
  if (!currentState.mixed) {
    var buckets = currentState.permissions || { editors: [], commenters: [], readers: [] };
    var bucketLevels = [
      { key: 'editors', level: 'edit' },
      { key: 'commenters', level: 'comment' },
      { key: 'readers', level: 'read' }
    ];
    for (var b = 0; b < bucketLevels.length; b++) {
      var list = buckets[bucketLevels[b].key] || [];
      for (var i = 0; i < list.length; i++) {
        var perm = list[i];
        var subjectKey = perm.type === 'group'
          ? normalizeGroupId_(perm.subject)
          : normalizeEmail_(perm.subject);
        roleByKey[(perm.type || 'user') + ':' + subjectKey] = bucketLevels[b].level;
      }
    }
  }

  var subjects = [];
  for (var g = 0; g < groups.length; g++) {
    var group = groups[g];
    var groupKey = 'group:' + group.group_id;
    subjects.push({
      subject: group.group_id,
      type: 'group',
      label: formatCatalogGroupDisplayName_(group.group_name),
      level: roleByKey[groupKey] || 'none'
    });
  }
  for (var u = 0; u < users.length; u++) {
    var user = users[u];
    var userKey = 'user:' + user.email;
    subjects.push({
      subject: user.email,
      type: 'user',
      label: user.name,
      level: roleByKey[userKey] || 'none'
    });
  }
  return subjects;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{email: string, name: string}>}
 */
function readActiveUsersList_(sheet) {
  var list = [];
  if (!sheet) {
    return list;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return list;
  }
  var rows = sheet.getRange(2, 1, lastRow, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var active = String(rows[i][2] || 'true').toLowerCase();
    if (active === 'false' || active === '0' || active === 'нет') {
      continue;
    }
    var email = normalizeEmail_(rows[i][0]);
    if (!email) {
      continue;
    }
    list.push({
      email: email,
      name: String(rows[i][1] || email)
    });
  }
  list.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru');
  });
  return list;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{group_id: string, group_name: string}>}
 */
function readCatalogGroupsList_(sheet) {
  var map = {};
  if (!sheet) {
    return [];
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var rows = sheet.getRange(2, 1, lastRow, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var groupId = normalizeGroupId_(rows[i][0]);
    if (!groupId || map[groupId]) {
      continue;
    }
    map[groupId] = {
      group_id: groupId,
      group_name: String(rows[i][1] || groupId)
    };
  }
  return Object.keys(map).map(function (key) {
    return map[key];
  }).sort(function (a, b) {
    return a.group_name.localeCompare(b.group_name, 'ru');
  });
}

/**
 * @param {string} permissionsJson
 * @param {boolean} applyToSubtree
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function assignCatalogPermissions(permissionsJson, applyToSubtree) {
  try {
    return assignCatalogPermissionsCore_(permissionsJson, applyToSubtree === true);
  } catch (err) {
    return {
      ok: false,
      error: 'ASSIGN_PERMISSIONS_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} permissionsJson
 * @param {boolean} applyToSubtree
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function assignCatalogPermissionsCore_(permissionsJson, applyToSubtree) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Нет активной таблицы.');
  }
  if (!isSpreadsheetOwner_(ss)) {
    throw new Error('Назначать права может только владелец таблицы.');
  }

  var sheets = getCatalogSheets_(ss);
  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var selected = resolveSelectedCatalogEntries_(ss).filter(function (entry) {
    return !isProtectedCatalogNode_(entry);
  });
  if (!selected.length) {
    throw new Error('Выделите папки или файлы на листе Catalog.');
  }

  var permissions = normalizeCatalogPermissions_(parseCatalogPermissionsJson_(permissionsJson));
  validateAssignablePermissions_(permissions, ss);
  validateGroupPermissionsHaveMembers_(permissions, ss);

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var viewCol = getCatalogViewColumnMap_();
  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var plan = buildAssignPermissionsPlan_(selected, allNodes, applyToSubtree);
  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var groupsMap = readCatalogGroupsMap_(ss.getSheetByName('Groups'));
  enrichAssignablePermissionLabels_(permissions, usersMap, groupsMap);

  var driveSynced = 0;
  var driveFailed = [];
  var approvedFilesSkippedEditors = 0;
  for (var i = 0; i < plan.affected.length; i++) {
    var node = plan.affected[i];
    var permsToApply = permissions;
    if (node.type === 'file' && isCatalogNodeApproved_(node)) {
      permsToApply = filterPermissionsWithoutEditors_(permissions);
      approvedFilesSkippedEditors++;
    }
    applyPermissionsToCatalogNode_(sheets.data, dataCol, node, permsToApply, usersMap);
    if (node.type === 'file' && node.file_id) {
      try {
        syncDrivePermissionsFromJson_(node.file_id, permsToApply, ss);
        driveSynced++;
      } catch (driveErr) {
        driveFailed.push({
          name: node.name,
          message: String(driveErr)
        });
        Logger.log('assignCatalogPermissions Drive ' + node.file_id + ': ' + driveErr);
      }
    }
  }

  var affectedIds = plan.affected.map(function (node) {
    return node.id;
  });
  updateCatalogViewPermissionCells_(sheets.view, viewCol, sheets.data, dataCol, affectedIds, usersMap);

  return {
    ok: driveFailed.length === 0,
    data: {
      updated_count: plan.affected.length,
      drive_synced: driveSynced,
      drive_failed: driveFailed,
      approved_files_count: approvedFilesSkippedEditors
    },
    message: buildAssignPermissionsResultMessage_(
      plan.affected.length,
      driveSynced,
      driveFailed.length,
      approvedFilesSkippedEditors,
      driveFailed
    )
  };
}

/**
 * @param {Array} permissions
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function validateAssignablePermissions_(permissions, ss) {
  var usersMap = readActiveUsersMap_(ss.getSheetByName('Users'));
  var groupsMap = readCatalogGroupsMap_(ss.getSheetByName('Groups'));

  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'domain' || perm.type === 'anyone') {
      throw new Error('Допустимы только пользователи из Users или группы из Groups.');
    }
    if (perm.type === 'group') {
      if (!groupsMap[normalizeGroupId_(perm.subject)]) {
        throw new Error('Группа не найдена в листе Groups: ' + perm.subject);
      }
      continue;
    }
    if (!usersMap[normalizeEmail_(perm.subject)]) {
      throw new Error('Пользователь не найден среди активных Users: ' + perm.subject);
    }
  }
}

/**
 * @param {Array} permissions
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function validateGroupPermissionsHaveMembers_(permissions, ss) {
  var groupsSheet = ss.getSheetByName('Groups');
  var groupsMap = readCatalogGroupsMap_(groupsSheet);
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type !== 'group') {
      continue;
    }
    var groupId = normalizeGroupId_(perm.subject);
    var members = readCatalogGroupMembers_(groupsSheet, groupId);
    if (!members.length) {
      var group = groupsMap[groupId];
      var label = group ? group.group_name : perm.subject;
      throw new Error('В группе «' + label + '» нет участников. Добавьте участников в меню «Группы».');
    }
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, {email: string, name: string}>}
 */
function readActiveUsersMap_(sheet) {
  var map = {};
  var list = readActiveUsersList_(sheet);
  for (var i = 0; i < list.length; i++) {
    map[list[i].email] = list[i];
  }
  return map;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, {group_id: string, group_name: string}>}
 */
function readCatalogGroupsMap_(sheet) {
  var map = {};
  var list = readCatalogGroupsList_(sheet);
  for (var i = 0; i < list.length; i++) {
    map[list[i].group_id] = list[i];
  }
  return map;
}

/**
 * @param {string} groupId
 * @returns {string}
 */
function normalizeGroupId_(groupId) {
  return String(groupId || '').trim().toLowerCase();
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} groupId
 * @returns {Array<string>}
 */
function readCatalogGroupMembers_(sheet, groupId) {
  var members = [];
  if (!sheet || !groupId) {
    return members;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return members;
  }
  groupId = normalizeGroupId_(groupId);
  var rows = sheet.getRange(2, 1, lastRow, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (normalizeGroupId_(rows[i][0]) !== groupId) {
      continue;
    }
    var email = normalizeEmail_(rows[i][2]);
    if (email) {
      members.push(email);
    }
  }
  return members;
}

/**
 * Разворачивает виртуальные группы каталога в права пользователей для Drive.
 *
 * @param {Array} permissions
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Array}
 */
function expandCatalogGroupsInPermissions_(permissions, ss) {
  if (!permissions || !permissions.length) {
    return [];
  }
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss ? ss.getSheetByName('Groups') : null;
  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var groupsMap = readCatalogGroupsMap_(groupsSheet);
  var byUser = {};

  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (!perm || !perm.subject) {
      continue;
    }
    if (perm.type === 'group') {
      var groupId = normalizeGroupId_(perm.subject);
      var members = readCatalogGroupMembers_(groupsSheet, groupId);
      for (var m = 0; m < members.length; m++) {
        mergeExpandedUserPermission_(byUser, members[m], perm.level, usersMap);
      }
      continue;
    }
    if (perm.type === 'anyone' || perm.type === 'domain') {
      continue;
    }
    mergeExpandedUserPermission_(byUser, perm.subject, perm.level, usersMap);
  }

  var result = [];
  var emails = Object.keys(byUser);
  for (var e = 0; e < emails.length; e++) {
    result.push(byUser[emails[e]]);
  }
  return normalizeCatalogPermissions_(result);
}

/**
 * @param {Object.<string, {subject: string, type: string, level: string, display_name: string}>} byUser
 * @param {string} email
 * @param {string} level
 * @param {Object.<string, {email: string, name: string}>} usersMap
 */
function mergeExpandedUserPermission_(byUser, email, level, usersMap) {
  email = normalizeEmail_(email);
  if (!email) {
    return;
  }
  var existing = byUser[email];
  var finalLevel = existing ? pickHigherPermissionLevel_(existing.level, level) : level;
  var user = usersMap[email];
  byUser[email] = {
    subject: email,
    type: 'user',
    level: finalLevel,
    display_name: user ? user.name : email
  };
}

/**
 * @param {string} left
 * @param {string} right
 * @returns {number}
 */
function comparePermissionLevel_(left, right) {
  var order = { edit: 3, comment: 2, read: 1 };
  return (order[left] || 0) - (order[right] || 0);
}

/**
 * @param {string} left
 * @param {string} right
 * @returns {string}
 */
function pickHigherPermissionLevel_(left, right) {
  return comparePermissionLevel_(left, right) >= 0 ? left : right;
}

/**
 * @param {Array} permissions
 * @param {string} groupId
 * @returns {string|null}
 */
function findGroupPermissionLevel_(permissions, groupId) {
  groupId = normalizeGroupId_(groupId);
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'group' && normalizeGroupId_(perm.subject) === groupId) {
      return perm.level;
    }
  }
  return null;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {string} groupId
 * @returns {Array}
 */
function findCatalogDataNodesWithGroup_(dataSheet, dataCol, groupId) {
  var nodes = readCatalogDataNodes_(dataSheet, dataCol);
  var result = [];
  for (var i = 0; i < nodes.length; i++) {
    var permissions = parseCatalogPermissionsJson_(nodes[i].permissions_json);
    if (findGroupPermissionLevel_(permissions, groupId)) {
      result.push(nodes[i]);
    }
  }
  return result;
}

/**
 * @param {Array} permissions
 * @param {string} groupId
 * @param {Array<string>} memberEmails
 * @param {Object.<string, {email: string, name: string}>} usersMap
 * @param {Object.<string, {group_id: string, group_name: string}>} groupsMap
 * @returns {Array}
 */
function replaceGroupWithMemberPermissions_(permissions, groupId, memberEmails, usersMap, groupsMap) {
  var level = findGroupPermissionLevel_(permissions, groupId);
  if (!level) {
    return permissions;
  }
  var result = [];
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'group' && normalizeGroupId_(perm.subject) === normalizeGroupId_(groupId)) {
      continue;
    }
    result.push(perm);
  }
  for (var m = 0; m < memberEmails.length; m++) {
    result = upsertUserPermissionInList_(result, memberEmails[m], level, usersMap);
  }
  result = normalizeCatalogPermissions_(result);
  enrichAssignablePermissionLabels_(result, usersMap, groupsMap);
  return result;
}

/**
 * @param {Array} permissions
 * @param {string} email
 * @param {string} level
 * @param {Object.<string, {email: string, name: string}>} usersMap
 * @returns {Array}
 */
function upsertUserPermissionInList_(permissions, email, level, usersMap) {
  email = normalizeEmail_(email);
  if (!email) {
    return permissions;
  }
  var result = [];
  var found = false;
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'user' && normalizeEmail_(perm.subject) === email) {
      found = true;
      var user = usersMap[email];
      result.push({
        subject: email,
        type: 'user',
        level: pickHigherPermissionLevel_(perm.level, level),
        display_name: user ? user.name : email
      });
      continue;
    }
    result.push(perm);
  }
  if (!found) {
    var userInfo = usersMap[email];
    result.push({
      subject: email,
      type: 'user',
      level: level,
      display_name: userInfo ? userInfo.name : email
    });
  }
  return result;
}

/**
 * @param {Array} permissions
 * @param {string} email
 * @returns {Array}
 */
function removeUserPermissionFromList_(permissions, email) {
  email = normalizeEmail_(email);
  var result = [];
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'user' && normalizeEmail_(perm.subject) === email) {
      continue;
    }
    result.push(perm);
  }
  return result;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} groupId
 * @param {string} groupName
 * @param {Array<string>} memberEmails
 */
function writeCatalogGroupRows_(sheet, groupId, groupName, memberEmails) {
  groupId = normalizeGroupId_(groupId);
  groupName = String(groupName || groupId).trim();
  var kept = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rows = sheet.getRange(2, 1, lastRow, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (normalizeGroupId_(rows[i][0]) !== groupId) {
        kept.push(rows[i]);
      }
    }
  }
  for (var m = 0; m < memberEmails.length; m++) {
    kept.push([groupId, groupName, normalizeEmail_(memberEmails[m])]);
  }
  if (!memberEmails.length) {
    kept.push([groupId, groupName, '']);
  }
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, 3).setValues(kept);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} groupId
 */
function deleteCatalogGroupRows_(sheet, groupId) {
  writeCatalogGroupRows_(sheet, groupId, '', []);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} groupId
 */
function resyncDriveForCatalogGroup_(ss, groupId) {
  var sheets = getCatalogSheets_(ss);
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var nodes = findCatalogDataNodesWithGroup_(sheets.data, dataCol, groupId);
  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    if (node.type !== 'file' || !node.file_id) {
      continue;
    }
    var perms = parseCatalogPermissionsJson_(node.permissions_json);
    if (node.approved && isCatalogNodeApproved_(node)) {
      perms = filterPermissionsWithoutEditors_(perms);
    }
    try {
      syncDrivePermissionsFromJson_(node.file_id, perms, ss);
    } catch (err) {
      Logger.log('resyncDriveForCatalogGroup_ ' + node.file_id + ': ' + err);
    }
  }
}

/**
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getCatalogGroupsManageContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { ok: false, message: 'Нет активной таблицы.' };
    }
    if (!canUserManageCatalogGroups_(ss)) {
      return { ok: false, message: 'Управлять группами могут только редакторы таблицы.' };
    }

    var groupsSheet = ss.getSheetByName('Groups');
    var usersSheet = ss.getSheetByName('Users');
    if (!groupsSheet || !usersSheet) {
      return { ok: false, message: 'Сначала выполните setupSchema.' };
    }

    var groups = readCatalogGroupsList_(groupsSheet);
    for (var g = 0; g < groups.length; g++) {
      groups[g].member_emails = readCatalogGroupMembers_(groupsSheet, groups[g].group_id);
    }

    return {
      ok: true,
      data: {
        groups: groups,
        users: readActiveUsersList_(usersSheet)
      }
    };
  } catch (err) {
    return { ok: false, message: String(err) };
  }
}

/**
 * @param {string} groupId
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function getCatalogGroupDetails_(groupId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return { ok: false, message: 'Нет активной таблицы.' };
  }
  groupId = normalizeGroupId_(groupId);
  if (!groupId) {
    return { ok: false, message: 'Не указана группа.' };
  }
  var groupsSheet = ss.getSheetByName('Groups');
  var groupsMap = readCatalogGroupsMap_(groupsSheet);
  var group = groupsMap[groupId];
  if (!group) {
    return { ok: false, message: 'Группа не найдена.' };
  }
  return {
    ok: true,
    data: {
      group_id: group.group_id,
      group_name: group.group_name,
      member_emails: readCatalogGroupMembers_(groupsSheet, groupId)
    }
  };
}

function manageCatalogGroupsDialog_() {
  var context = getCatalogGroupsManageContext();
  if (!context.ok) {
    SpreadsheetApp.getUi().alert('Группы', context.message || context.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var template = HtmlService.createTemplateFromFile('GroupsManageDialog');
  template.initialContext = JSON.stringify(context);
  var html = template.evaluate()
    .setWidth(560)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Группы');
}

/**
 * Создание, изменение или удаление группы каталога.
 *
 * @param {string} payloadJson
 * @returns {{ok: boolean, data?: object, message?: string, error?: string}}
 */
function saveCatalogGroup(payloadJson) {
  try {
    return saveCatalogGroupCore_(payloadJson);
  } catch (err) {
    return { ok: false, error: 'SAVE_GROUP_FAILED', message: String(err) };
  }
}

/**
 * @param {string} payloadJson
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function saveCatalogGroupCore_(payloadJson) {
  var payload = {};
  try {
    payload = JSON.parse(payloadJson || '{}');
  } catch (err) {
    throw new Error('Некорректные данные группы.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Нет активной таблицы.');
  }
  if (!canUserManageCatalogGroups_(ss)) {
    throw new Error('Управлять группами могут только редакторы таблицы.');
  }

  var groupsSheet = ss.getSheetByName('Groups');
  var usersSheet = ss.getSheetByName('Users');
  if (!groupsSheet || !usersSheet) {
    throw new Error('Сначала выполните setupSchema.');
  }

  var deleteGroup = payload.delete === true;
  var groupId = normalizeGroupId_(payload.group_id);
  var groupName = String(payload.group_name || '').trim();
  var memberEmails = normalizeCatalogGroupMemberEmails_(payload.member_emails || [], usersSheet);

  if (deleteGroup) {
    if (!groupId) {
      throw new Error('Не указана группа для удаления.');
    }
    return deleteCatalogGroupCore_(ss, groupsSheet, groupId);
  }

  if (!groupName) {
    throw new Error('Введите название группы.');
  }

  var isNew = !groupId;
  if (isNew) {
    groupId = generateId_();
  } else if (!readCatalogGroupsMap_(groupsSheet)[groupId]) {
    throw new Error('Группа не найдена: ' + groupId);
  }

  var oldMembers = isNew ? [] : readCatalogGroupMembers_(groupsSheet, groupId);
  writeCatalogGroupRows_(groupsSheet, groupId, groupName, memberEmails);

  var membershipChanged = !arraysEqualAsSets_(oldMembers, memberEmails);
  if (!isNew && membershipChanged) {
    resyncDriveForCatalogGroup_(ss, groupId);
  }

  return {
    ok: true,
    data: {
      group_id: groupId,
      group_name: groupName,
      member_emails: memberEmails,
      created: isNew
    },
    message: isNew ? 'Группа создана.' : 'Группа сохранена.'
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} groupsSheet
 * @param {string} groupId
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function deleteCatalogGroupCore_(ss, groupsSheet, groupId) {
  groupId = normalizeGroupId_(groupId);
  var groupsMap = readCatalogGroupsMap_(groupsSheet);
  var group = groupsMap[groupId];
  if (!group) {
    throw new Error('Группа не найдена: ' + groupId);
  }

  var members = readCatalogGroupMembers_(groupsSheet, groupId);
  var sheets = getCatalogSheets_(ss);
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var viewCol = getCatalogViewColumnMap_();
  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var affectedNodes = findCatalogDataNodesWithGroup_(sheets.data, dataCol, groupId);
  var affectedIds = [];

  for (var i = 0; i < affectedNodes.length; i++) {
    var node = affectedNodes[i];
    var permissions = parseCatalogPermissionsJson_(node.permissions_json);
    var updated = replaceGroupWithMemberPermissions_(permissions, groupId, members, usersMap, groupsMap);
    applyPermissionsToCatalogNode_(sheets.data, dataCol, node, updated, usersMap);
    affectedIds.push(node.id);
    if (node.type === 'file' && node.file_id) {
      var permsForDrive = updated;
      if (isCatalogNodeApproved_(node)) {
        permsForDrive = filterPermissionsWithoutEditors_(updated);
      }
      try {
        syncDrivePermissionsFromJson_(node.file_id, permsForDrive, ss);
      } catch (driveErr) {
        Logger.log('deleteCatalogGroupCore_ drive ' + node.file_id + ': ' + driveErr);
      }
    }
  }

  deleteCatalogGroupRows_(groupsSheet, groupId);
  if (affectedIds.length) {
    updateCatalogViewPermissionCells_(sheets.view, viewCol, sheets.data, dataCol, affectedIds, usersMap);
  }

  return {
    ok: true,
    data: {
      group_id: groupId,
      affected_count: affectedIds.length
    },
    message: 'Группа удалена. В правах заменена на ' + members.length + ' участник(ов); обновлено элементов: ' + affectedIds.length + '.'
  };
}

/**
 * @param {Array} memberEmails
 * @param {GoogleAppsScript.Spreadsheet.Sheet} usersSheet
 * @returns {Array<string>}
 */
function normalizeCatalogGroupMemberEmails_(memberEmails, usersSheet) {
  var usersMap = readActiveUsersMap_(usersSheet);
  var result = [];
  var seen = {};
  for (var i = 0; i < memberEmails.length; i++) {
    var email = normalizeEmail_(memberEmails[i]);
    if (!email || seen[email]) {
      continue;
    }
    if (!usersMap[email]) {
      throw new Error('Пользователь не найден среди активных Users: ' + email);
    }
    seen[email] = true;
    result.push(email);
  }
  return result;
}

/**
 * @param {Array<string>} left
 * @param {Array<string>} right
 * @returns {boolean}
 */
function arraysEqualAsSets_(left, right) {
  var leftSet = {};
  var rightSet = {};
  for (var i = 0; i < left.length; i++) {
    leftSet[left[i]] = true;
  }
  for (var j = 0; j < right.length; j++) {
    rightSet[right[j]] = true;
  }
  var leftKeys = Object.keys(leftSet);
  var rightKeys = Object.keys(rightSet);
  if (leftKeys.length !== rightKeys.length) {
    return false;
  }
  for (var k = 0; k < leftKeys.length; k++) {
    if (!rightSet[leftKeys[k]]) {
      return false;
    }
  }
  return true;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {boolean}
 */
function canUserManageCatalogGroups_(ss) {
  return canUserApproveCatalogFiles_(ss);
}

/**
 * @param {Array} permissions
 * @param {Object.<string, {email: string, name: string}>} usersMap
 * @param {Object.<string, {group_id: string, group_name: string}>} groupsMap
 */
function enrichAssignablePermissionLabels_(permissions, usersMap, groupsMap) {
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (perm.type === 'group') {
      var group = groupsMap[normalizeGroupId_(perm.subject)];
      perm.display_name = group ? formatCatalogGroupDisplayName_(group.group_name) : perm.subject;
      continue;
    }
    var user = usersMap[normalizeEmail_(perm.subject)];
    perm.display_name = user ? user.name : perm.subject;
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {Array} permissions
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 */
function applyPermissionsToCatalogNode_(sheet, col, node, permissions, usersByEmail) {
  var display = buildPermissionDisplay_(permissions, usersByEmail);
  var json = JSON.stringify(permissions);
  sheet.getRange(node.rowIndex, col.permissions_json).setValue(json);
  sheet.getRange(node.rowIndex, col.editors).setValue(formatCatalogPermissionCell_(display.editors));
  sheet.getRange(node.rowIndex, col.commenters).setValue(formatCatalogPermissionCell_(display.commenters));
  sheet.getRange(node.rowIndex, col.readers).setValue(formatCatalogPermissionCell_(display.readers));
  node.permissions_json = json;
  node.editors = display.editors;
  node.commenters = display.commenters;
  node.readers = display.readers;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object.<string, number>} viewCol
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {Array<string>} nodeIds
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 */
function updateCatalogViewPermissionCells_(viewSheet, viewCol, dataSheet, dataCol, nodeIds, usersByEmail) {
  var indexSheet = getCatalogIndexSheet_(viewSheet.getParent());
  var indexCol = getCatalogIndexColumnMap_();
  var idToViewRow = buildCatalogViewIdToRowMap_(indexSheet, indexCol);
  for (var i = 0; i < nodeIds.length; i++) {
    var node = findCatalogDataEntryById_(dataSheet, dataCol, nodeIds[i]);
    var viewRow = idToViewRow[nodeIds[i]];
    if (!node || !viewRow) {
      continue;
    }
    var display = resolveNodePermissionDisplay_(node, usersByEmail);
    viewSheet.getRange(viewRow, viewCol.editors).setValue(formatCatalogPermissionCell_(display.editors));
    viewSheet.getRange(viewRow, viewCol.commenters).setValue(formatCatalogPermissionCell_(display.commenters));
    viewSheet.getRange(viewRow, viewCol.readers).setValue(formatCatalogPermissionCell_(display.readers));
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet
 * @param {Object.<string, number>} indexCol
 * @returns {Object.<string, number>}
 */
function buildCatalogViewIdToRowMap_(indexSheet, indexCol) {
  var map = {};
  var lastRow = indexSheet.getLastRow();
  if (lastRow < 2) {
    return map;
  }
  var values = indexSheet.getRange(2, indexCol.catalog_id, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var itemId = readCell_(values[i][0]);
    if (itemId) {
      map[itemId] = i + 2;
    }
  }
  return map;
}

/**
 * @param {Object} node
 * @returns {boolean}
 */
function isCatalogNodeApproved_(node) {
  var value = readCell_(node.approved);
  if (!value) {
    return false;
  }
  var text = String(value).trim();
  return text !== '' && text !== '—' && text !== '-';
}

/**
 * Значение ячейки «Утверждение» в Catalog.
 *
 * @param {Object} node
 * @returns {string}
 */
function resolveCatalogApprovalViewValue_(node) {
  if (node.type !== 'file') {
    return '—';
  }
  if (isCatalogNodeApproved_(node)) {
    return readCell_(node.approved);
  }
  return '—';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Array} visibleNodes
 * @param {Object.<string, number>} viewCol
 */
function applyCatalogViewApprovalCells_(viewSheet, visibleNodes, viewCol) {
  if (!visibleNodes.length) {
    return;
  }
  for (var i = 0; i < visibleNodes.length; i++) {
    applyCatalogViewApprovalCellForRow_(viewSheet, i + 2, visibleNodes[i], viewCol);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {number} row
 * @param {Object} node
 * @param {Object.<string, number>} viewCol
 */
function applyCatalogViewApprovalCellForRow_(viewSheet, row, node, viewCol) {
  var cell = viewSheet.getRange(row, viewCol.approved);
  cell.clearDataValidations();
  cell.clearNote();
  if (node.type !== 'file' || !isCatalogNodeApproved_(node)) {
    cell.setValue('—');
    return;
  }
  cell.setValue(readCell_(node.approved));
}

/**
 * Удаляет устаревший onEdit-триггер утверждения (чек-боксы больше не используются).
 */
function removeCatalogApprovalEditTriggers_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      if (trigger.getEventType() !== ScriptApp.EventType.ON_EDIT) {
        continue;
      }
      if (trigger.getHandlerFunction() === 'handleCatalogSheetEdit_') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  } catch (err) {
    Logger.log('removeCatalogApprovalEditTriggers_: ' + err);
  }
}

/**
 * @returns {{ok: boolean, data?: {item_id: string, file_label: string}, message?: string, error?: string}}
 */
function resolveApproveCatalogSelection_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return { ok: false, message: 'Нет активной таблицы.' };
  }

  try {
    getCatalogSheets_(ss);
  } catch (err) {
    return { ok: false, message: String(err) };
  }

  var selected;
  try {
    selected = resolveSelectedCatalogEntries_(ss);
  } catch (err) {
    return { ok: false, message: String(err) };
  }

  if (!selected.length) {
    return { ok: false, message: 'Выделите один файл на листе Catalog.' };
  }

  var files = selected.filter(function (entry) {
    return entry.type === 'file' && !isCatalogNodeApproved_(entry) && !isProtectedCatalogNode_(entry);
  });
  if (!files.length) {
    return { ok: false, message: 'Выделите неутверждённый файл (не папку).' };
  }
  if (files.length > 1) {
    return { ok: false, message: 'Утверждайте по одному файлу за раз.' };
  }

  if (!canUserApproveCatalogFiles_(ss)) {
    return { ok: false, message: 'Утвердить файл может только редактор таблицы.' };
  }

  var node = files[0];
  return {
    ok: true,
    data: {
      item_id: node.id,
      file_label: stripFolderPrefix_(node.name, node.folder_code) || node.name || 'файл'
    }
  };
}

function approveCatalogFileMenu_() {
  assertCatalogImportNotBusy_();
  var selection = resolveApproveCatalogSelection_();
  if (!selection.ok) {
    SpreadsheetApp.getUi().alert(
      'Утверждение',
      selection.message || selection.error || 'Ошибка',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var fileLabel = selection.data.file_label;
  var ui = SpreadsheetApp.getUi();
  var confirmed = ui.alert(
    'Утвердить файл?',
    'Файл «' + fileLabel + '»\n\nПосле утверждения все получат только просмотр. Отменить нельзя.',
    ui.ButtonSet.YES_NO
  ) === ui.Button.YES;
  if (!confirmed) {
    return;
  }

  var result = approveCatalogFile(selection.data.item_id);
  SpreadsheetApp.getUi().alert(
    result.ok ? 'Готово' : 'Ошибка',
    result.message || result.error || 'Неизвестная ошибка',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

var PENDING_APPROVE_DRIVE_SYNC_KEY = 'pending_approve_drive_sync_queue';
var APPROVE_DRIVE_SYNC_HANDLER = 'processPendingApproveDriveSync_';

/**
 * @returns {Array}
 */
function loadPendingApproveDriveSyncQueue_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(PENDING_APPROVE_DRIVE_SYNC_KEY);
    if (!raw) {
      return [];
    }
    var parsed = JSON.parse(raw);
    return parsed && parsed.length ? parsed : [];
  } catch (err) {
    Logger.log('loadPendingApproveDriveSyncQueue_: ' + err);
    return [];
  }
}

/**
 * @param {Array} queue
 */
function savePendingApproveDriveSyncQueue_(queue) {
  PropertiesService.getDocumentProperties().setProperty(
    PENDING_APPROVE_DRIVE_SYNC_KEY,
    JSON.stringify(queue || [])
  );
}

/**
 * @param {string} itemId
 * @param {string} fileId
 * @param {string} permissionsJson
 */
function scheduleApproveDriveSync_(itemId, fileId, permissionsJson) {
  if (!fileId) {
    return;
  }
  var queue = loadPendingApproveDriveSyncQueue_();
  queue.push({
    item_id: itemId,
    file_id: fileId,
    permissions_json: permissionsJson || '[]'
  });
  savePendingApproveDriveSyncQueue_(queue);
  ensureApproveDriveSyncTrigger_();
}

function ensureApproveDriveSyncTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVE_DRIVE_SYNC_HANDLER) {
      return;
    }
  }
  ScriptApp.newTrigger(APPROVE_DRIVE_SYNC_HANDLER)
    .timeBased()
    .after(2000)
    .create();
}

function deleteApproveDriveSyncTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVE_DRIVE_SYNC_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function processPendingApproveDriveSync_() {
  deleteApproveDriveSyncTriggers_();
  var queue = loadPendingApproveDriveSyncQueue_();
  if (!queue.length) {
    return;
  }
  savePendingApproveDriveSyncQueue_([]);
  for (var i = 0; i < queue.length; i++) {
    var job = queue[i];
    if (!job || !job.file_id) {
      continue;
    }
    try {
      var perms = buildApprovedViewOnlyPermissions_(parseCatalogPermissionsJson_(job.permissions_json));
      var driveResult = syncDrivePermissionsForApproval_(job.file_id, perms);
      if (!driveResult.ok) {
        Logger.log('processPendingApproveDriveSync_ ' + job.file_id + ': ' + driveResult.message);
      }
    } catch (err) {
      Logger.log('processPendingApproveDriveSync_ ' + job.file_id + ': ' + err);
    }
  }
}

/**
 * Утверждение на Drive: понижение редакторов/комментаторов до чтения, без жёстких ошибок.
 *
 * @param {string} fileId
 * @param {Array} viewOnlyPerms
 * @returns {{ok: boolean, message?: string}}
 */
function syncDrivePermissionsForApproval_(fileId, viewOnlyPerms) {
  if (!fileId) {
    return { ok: true };
  }

  var desiredMap = {};
  var permissions = normalizeCatalogPermissions_(viewOnlyPerms || []);
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (!perm.subject || perm.type === 'anyone' || perm.type === 'domain') {
      continue;
    }
    desiredMap[buildDrivePermissionKey_(perm.type, perm.subject)] = perm;
  }

  var errors = [];
  try {
    var current = listDrivePermissionsWithIds_(fileId);
    for (var c = 0; c < current.length; c++) {
      var currentPerm = current[c];
      if (currentPerm.inherited) {
        continue;
      }
      var key = buildDrivePermissionKey_(currentPerm.type, currentPerm.subject);
      var desired = desiredMap[key];
      var targetRole = 'reader';
      if (desired) {
        targetRole = mapLevelToDriveRole_(desired.level);
      } else if (currentPerm.role !== 'writer' && currentPerm.role !== 'commenter') {
        continue;
      }
      if (currentPerm.role === targetRole) {
        delete desiredMap[key];
        continue;
      }
      try {
        Drive.Permissions.update(
          { role: targetRole },
          fileId,
          currentPerm.id,
          { supportsAllDrives: true, sendNotificationEmail: false }
        );
      } catch (updateErr) {
        errors.push(String(updateErr));
        Logger.log('syncDrivePermissionsForApproval_ update ' + currentPerm.id + ': ' + updateErr);
      }
      delete desiredMap[key];
    }

    var remainingKeys = Object.keys(desiredMap);
    for (var r = 0; r < remainingKeys.length; r++) {
      var newPerm = desiredMap[remainingKeys[r]];
      try {
        Drive.Permissions.create({
          role: mapLevelToDriveRole_(newPerm.level),
          type: newPerm.type === 'group' ? 'group' : 'user',
          emailAddress: newPerm.subject
        }, fileId, {
          supportsAllDrives: true,
          sendNotificationEmail: false
        });
      } catch (createErr) {
        errors.push(String(createErr));
        Logger.log('syncDrivePermissionsForApproval_ create: ' + createErr);
      }
    }

    try {
      enforceWritersCanShare_(fileId);
    } catch (shareErr) {
      Logger.log('syncDrivePermissionsForApproval_ writersCanShare: ' + shareErr);
    }
  } catch (err) {
    return { ok: false, message: String(err) };
  }

  if (errors.length) {
    return { ok: false, message: 'частично (' + errors.length + ' ошибок)' };
  }
  return { ok: true };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {boolean}
 */
function canUserApproveCatalogFiles_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return false;
  }
  var userEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (!userEmail) {
    return false;
  }
  if (isSpreadsheetOwner_(ss)) {
    return true;
  }
  try {
    var editors = ss.getEditors();
    for (var i = 0; i < editors.length; i++) {
      if (normalizeEmail_(editors[i].getEmail()) === userEmail) {
        return true;
      }
    }
  } catch (err) {
    Logger.log('canUserApproveCatalogFiles_: ' + err);
  }
  return false;
}

/**
 * @param {Object.<string, {email: string, name: string}>} usersMap
 * @param {string} approverEmail
 * @returns {string}
 */
function resolveApproverLabel_(usersMap, approverEmail) {
  var user = usersMap[normalizeEmail_(approverEmail)];
  if (user && user.name) {
    return user.name;
  }
  return approverEmail || 'неизвестно';
}

/**
 * @param {Array} permissions
 * @returns {Array}
 */
function buildApprovedViewOnlyPermissions_(permissions) {
  var normalized = normalizeCatalogPermissions_(permissions || []);
  var seen = {};
  var result = [];
  for (var i = 0; i < normalized.length; i++) {
    var perm = normalized[i];
    var key = (perm.type || 'user') + ':' + normalizeEmail_(perm.subject);
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    result.push({
      subject: perm.subject,
      type: perm.type,
      level: 'read',
      display_name: perm.display_name
    });
  }
  return result;
}

/**
 * @param {string} itemId
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function approveCatalogFile(itemId) {
  try {
    return approveCatalogFileCore_(itemId);
  } catch (err) {
    return {
      ok: false,
      error: 'APPROVE_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} itemId
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function approveCatalogFileCore_(itemId) {
  if (!itemId) {
    throw new Error('Не найден элемент каталога.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Нет активной таблицы.');
  }
  if (!canUserApproveCatalogFiles_(ss)) {
    throw new Error('Утвердить файл может только редактор таблицы.');
  }

  var sheets = getCatalogSheets_(ss);
  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var viewCol = getCatalogViewColumnMap_();
  var node = findCatalogDataEntryById_(sheets.data, dataCol, itemId);
  if (!node || node.type !== 'file') {
    throw new Error('Утверждать можно только файлы.');
  }
  if (isCatalogNodeApproved_(node)) {
    throw new Error('Файл уже утверждён.');
  }

  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  if (trashId && isNodeInTrashSubtree_(node, trashId, nodesById)) {
    throw new Error('Нельзя утвердить файл в «Корзине».');
  }

  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var approverEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  var approverLabel = resolveApproverLabel_(usersMap, approverEmail);
  var viewOnlyPerms = buildApprovedViewOnlyPermissions_(parseCatalogPermissionsJson_(node.permissions_json));

  applyPermissionsToCatalogNode_(sheets.data, dataCol, node, viewOnlyPerms, usersMap);
  sheets.data.getRange(node.rowIndex, dataCol.approved).setValue(approverLabel);
  node.approved = approverLabel;
  updateCatalogViewAfterApproval_(sheets.view, viewCol, node, usersMap);
  SpreadsheetApp.flush();

  if (node.file_id) {
    scheduleApproveDriveSync_(itemId, node.file_id, node.permissions_json);
  }

  return {
    ok: true,
    data: {
      item_id: itemId,
      approved_by: approverLabel,
      permissions_count: viewOnlyPerms.length,
      file_id: node.file_id || ''
    },
    message: 'Файл «' + (stripFolderPrefix_(node.name, node.folder_code) || node.name || itemId) +
      '» утверждён в каталоге.\nПрава на Google Drive обновляются в фоне (обычно несколько секунд).'
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object.<string, number>} viewCol
 * @param {Object} node
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 */
function updateCatalogViewAfterApproval_(viewSheet, viewCol, node, usersByEmail) {
  var indexSheet = getCatalogIndexSheet_(viewSheet.getParent());
  var indexCol = getCatalogIndexColumnMap_();
  var idToViewRow = buildCatalogViewIdToRowMap_(indexSheet, indexCol);
  var viewRow = idToViewRow[node.id];
  if (!viewRow) {
    return;
  }
  var display = resolveNodePermissionDisplay_(node, usersByEmail);
  viewSheet.getRange(viewRow, viewCol.editors).setValue(formatCatalogPermissionCell_(display.editors));
  viewSheet.getRange(viewRow, viewCol.commenters).setValue(formatCatalogPermissionCell_(display.commenters));
  viewSheet.getRange(viewRow, viewCol.readers).setValue(formatCatalogPermissionCell_(display.readers));
  applyCatalogViewApprovalCellForRow_(viewSheet, viewRow, node, viewCol);
}

/**
 * @param {Array} affected
 * @returns {boolean}
 */
function planHasApprovedFiles_(affected) {
  for (var i = 0; i < affected.length; i++) {
    if (affected[i].type === 'file' && isCatalogNodeApproved_(affected[i])) {
      return true;
    }
  }
  return false;
}

/**
 * Утверждённым файлам нельзя назначать редакторов — только чтение и комментирование.
 *
 * @param {Array} permissions
 * @returns {Array}
 */
function filterPermissionsWithoutEditors_(permissions) {
  var result = [];
  for (var i = 0; i < permissions.length; i++) {
    if (permissions[i].level !== 'edit') {
      result.push(permissions[i]);
    }
  }
  return result;
}

/**
 * @param {number} updatedCount
 * @param {number} driveSynced
 * @param {number} driveFailedCount
 * @param {number=} approvedFilesCount
 * @returns {string}
 */
function buildAssignPermissionsResultMessage_(updatedCount, driveSynced, driveFailedCount, approvedFilesCount, driveFailed) {
  var parts = ['Обновлено элементов: ' + updatedCount + '.'];
  parts.push('Синхронизировано файлов на Drive: ' + driveSynced + '.');
  if (approvedFilesCount) {
    parts.push('Утверждённых файлов (без редакторов): ' + approvedFilesCount + '.');
  }
  if (driveFailedCount) {
    parts.push('Ошибок Drive: ' + driveFailedCount + '.');
    if (driveFailed && driveFailed.length) {
      parts.push(driveFailed[0].name + ': ' + driveFailed[0].message);
    }
  }
  return parts.join(' ');
}

/**
 * Открывает диалог создания Google-файла (Документ / Таблица / Презентация).
 */
function createFileDialog_() {
  var contextResult = getCreateFileContext();
  var template = HtmlService.createTemplateFromFile('CreateFileDialog');
  template.initialContext = JSON.stringify(contextResult);
  var html = template.evaluate()
    .setWidth(480)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Создать файл');
}

/**
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getCreateFileContext() {
  var context = getCreateFolderContext();
  if (!context.ok) {
    return context;
  }
  context.data.file_kinds = [
    { id: 'document', label: 'Google Документ' },
    { id: 'spreadsheet', label: 'Google Таблица' },
    { id: 'presentation', label: 'Google Презентация' }
  ];
  return context;
}

/**
 * @param {string} parentId
 * @param {string} label
 * @param {string} fileKind document|spreadsheet|presentation
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function createCatalogFile(parentId, label, fileKind) {
  try {
    return createCatalogFileCore_(parentId, label, fileKind);
  } catch (err) {
    return {
      ok: false,
      error: 'CREATE_FILE_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} parentId
 * @param {string} label
 * @param {string} fileKind
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function createCatalogFileCore_(parentId, label, fileKind) {
  var mimeType = mapCatalogCreateFileKindToMime_(fileKind);
  if (!mimeType) {
    throw new Error('Неизвестный тип файла.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var parent = findCatalogNodeById_(sheets.data, dataCol, parentId);
  validateCatalogImportParent_(parent);

  var cleanLabel = sanitizeFolderLabel_(label);
  if (!cleanLabel) {
    throw new Error('Введите имя файла.');
  }

  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var fileFolderCode = getChildFileFolderCode_(parent);
  var unique = allocateUniqueCatalogName_(
    parent.id,
    fileFolderCode,
    cleanLabel,
    allNodes,
    {},
    {}
  );

  var driveFile = ingestCatalogDriveFileViaImportStaging_({
    ss: ss,
    finalName: unique.name,
    mimeType: mimeType
  });

  var usersMap = readUsersMapFromSpreadsheet_(ss);
  var newNodeId;
  try {
    newNodeId = registerCatalogDriveFileRow_(
      sheets,
      dataCol,
      parent,
      driveFile.id,
      unique.name,
      driveFile.mimeType || mimeType,
      fileFolderCode,
      usersMap,
      true
    );
  } catch (permErr) {
    Drive.Files.remove(driveFile.id, { supportsAllDrives: true });
    throw new Error('Файл создан, но не удалось применить права: ' + permErr);
  }

  insertCatalogViewNodeAfterCreate_(sheets.view, sheets.data, newNodeId, parent.id);

  return {
    ok: true,
    data: {
      file_id: driveFile.id,
      name: unique.name,
      parent_id: parent.id
    },
    message: 'Файл создан: ' + unique.name
  };
}

/**
 * @param {string} fileKind
 * @returns {string}
 */
function mapCatalogCreateFileKindToMime_(fileKind) {
  if (fileKind === 'document') {
    return 'application/vnd.google-apps.document';
  }
  if (fileKind === 'spreadsheet') {
    return 'application/vnd.google-apps.spreadsheet';
  }
  if (fileKind === 'presentation') {
    return 'application/vnd.google-apps.presentation';
  }
  return '';
}

/**
 * Открывает диалог импорта из локального хранилища или Google Drive.
 */
function importDialog_() {
  var contextResult = getImportContext();
  var template = HtmlService.createTemplateFromFile('ImportDialog');
  template.initialContext = JSON.stringify(contextResult);
  var html = template.evaluate()
    .setWidth(600)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Импорт');
}

/**
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getImportContext() {
  var context = getCreateFolderContext();
  if (!context.ok) {
    return context;
  }
  context.data.max_upload_mb = Math.floor(CATALOG_MAX_UPLOAD_BYTES / (1024 * 1024));
  var importCheck = checkCatalogImportFolderAccess_();
  context.data.catalog_drive_access = importCheck.ok;
  if (importCheck.ok) {
    context.data.catalog_drive_folder_name = importCheck.data.name;
  } else {
    context.data.catalog_drive_access_message = importCheck.message || importCheck.error;
  }
  var webAppCheck = checkImportWebAppConfigured_();
  context.data.import_webapp_configured = webAppCheck.ok;
  if (!webAppCheck.ok) {
    context.data.import_webapp_message = webAppCheck.message || webAppCheck.error;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  context.data.catalog_owner_email = getSpreadsheetOwnerEmail_(ss);
  context.data.is_catalog_owner = isSpreadsheetOwner_(ss);
  context.data.import_batch_chunk_size = DRIVE_BATCH_MAX;
  return context;
}

/**
 * Проверяет доступ текущего пользователя к staging-папке _Import на Drive.
 * @returns {{ok: boolean, data?: {name: string, folder_id: string}, error?: string, message?: string}}
 */
function checkCatalogImportFolderAccess_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var importFolderId = ensureCatalogImportDriveFolder_(ss);
    var folderCheck = validateRootFolder_(importFolderId);
    if (!folderCheck.ok) {
      return folderCheck;
    }
    return {
      ok: true,
      data: {
        name: folderCheck.data.name,
        folder_id: importFolderId
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'IMPORT_FOLDER_ACCESS',
      message: String(err)
    };
  }
}

/**
 * Проверяет, видит ли текущий пользователь папку каталога на Drive (root_folder_id).
 * @returns {{ok: boolean, data?: {name: string, folder_id: string}, error?: string, message?: string}}
 */
function checkCatalogRootFolderAccess_() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!settingsSheet) {
    return {
      ok: false,
      error: 'NO_SETTINGS',
      message: 'Сначала выполните setupSchema.'
    };
  }
  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    return {
      ok: false,
      error: 'NO_ROOT_FOLDER',
      message: 'Укажите root_folder_id на листе Settings.'
    };
  }
  var folderCheck = validateRootFolder_(rootFolderId);
  if (!folderCheck.ok) {
    return folderCheck;
  }
  return {
    ok: true,
    data: {
      name: folderCheck.data.name,
      folder_id: rootFolderId
    }
  };
}

/**
 * Подготовка импорта: план, флаги can_move (без копирования).
 * @param {string} linksText
 * @param {string} extraIdsJson
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function prepareDriveImportFromSources(linksText, extraIdsJson) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rootCheck = checkCatalogImportFolderAccess_();
    if (!rootCheck.ok) {
      return rootCheck;
    }
    var webAppCheck = checkImportWebAppConfigured_();
    if (!webAppCheck.ok && !isSpreadsheetOwner_(ss)) {
      return webAppCheck;
    }

    var ids = collectDriveImportInputIds_(linksText, extraIdsJson);
    if (!ids.length) {
      return {
        ok: false,
        error: 'NO_SOURCES',
        message: 'Укажите ссылки, ID или выберите файлы из поиска.'
      };
    }

    var resolved = resolveDriveImportFromIds_(ids);
    if (!resolved.ok) {
      return resolved;
    }

    annotateDriveImportEntries_(resolved.data.entries);

    var movableCount = countMovableDriveImportEntries_(resolved.data.entries);
    var message = 'Файлов к импорту: ' + resolved.data.entries.length;
    if (movableCount) {
      message += ' (можно перенести без копии: ' + movableCount + ')';
    }
    if (resolved.data.resolve_errors.length) {
      message += '\nПропущено: ' + resolved.data.resolve_errors.join('; ');
    }

    return {
      ok: true,
      data: {
        entries: resolved.data.entries,
        count: resolved.data.entries.length,
        movable_count: movableCount,
        items: resolved.data.items
      },
      message: message
    };
  } catch (err) {
    return {
      ok: false,
      error: 'PREPARE_DRIVE_IMPORT_FAILED',
      message: String(err)
    };
  }
}

/**
 * Перед выполнением: доступ владельцу каталога для режима «копия».
 * @param {string} entriesJson
 * @returns {{ok: boolean, message?: string, error?: string}}
 */
function prepareDriveImportExecution(entriesJson) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (isSpreadsheetOwner_(ss)) {
      return { ok: true, message: 'Владелец каталога: подготовка не требуется.' };
    }
    resetImportOwnershipBatchPending_();
    var entries = JSON.parse(entriesJson || '[]');
    return grantCatalogOwnerReadAccessForImportEntries_(entries);
  } catch (err) {
    return {
      ok: false,
      error: 'PREPARE_EXECUTION_FAILED',
      message: String(err)
    };
  }
}

/**
 * @deprecated Используйте prepareDriveImportFromSources
 */
function beginDriveImportFromLinks(linksText) {
  var prepared = prepareDriveImportFromSources(linksText, '[]');
  if (!prepared.ok) {
    return prepared;
  }
  var grantResult = prepareDriveImportExecution(JSON.stringify(prepared.data.entries));
  if (!grantResult.ok) {
    return grantResult;
  }
  return {
    ok: true,
    data: {
      entries: prepared.data.entries,
      count: prepared.data.count
    },
    message: 'Импорт ' + prepared.data.count + ' файлов…'
  };
}

/**
 * Поиск файлов и папок на Drive по имени (среди доступных импортёру).
 * @param {string} query
 * @returns {{ok: boolean, data?: {results: Array}, error?: string, message?: string}}
 */
function searchDriveFilesForImport(query) {
  try {
    var trimmed = String(query || '').trim();
    if (trimmed.length < 2) {
      return {
        ok: false,
        error: 'QUERY_TOO_SHORT',
        message: 'Введите не менее 2 символов для поиска.'
      };
    }

    var escaped = escapeDriveQueryValue_(trimmed);
    var response = Drive.Files.list({
      q: "name contains '" + escaped + "' and trashed = false",
      pageSize: 40,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      orderBy: 'folder,name',
      fields: 'files(id,name,mimeType,owners(emailAddress))'
    });

    var batch = response.files || [];
    var results = [];
    for (var i = 0; i < batch.length; i++) {
      var item = batch[i];
      if (item.mimeType === DRIVE_SHORTCUT_MIME) {
        continue;
      }
      var ownerEmail = '';
      if (item.owners && item.owners.length) {
        ownerEmail = item.owners[0].emailAddress || '';
      }
      results.push({
        id: item.id,
        name: item.name || 'Без имени',
        mime_type: item.mimeType || '',
        is_folder: item.mimeType === DRIVE_FOLDER_MIME,
        can_move: isImporterOwnerOfDriveFileId_(item.id, ownerEmail)
      });
    }

    if (!results.length) {
      return {
        ok: true,
        data: { results: [] },
        message: 'Ничего не найдено. Проверьте имя и доступ к файлу.'
      };
    }

    return {
      ok: true,
      data: { results: results },
      message: 'Найдено: ' + results.length
    };
  } catch (err) {
    return {
      ok: false,
      error: 'SEARCH_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} value
 * @returns {string}
 */
function escapeDriveQueryValue_(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

/**
 * @param {string} linksText
 * @param {string} extraIdsJson
 * @returns {Array<string>}
 */
function collectDriveImportInputIds_(linksText, extraIdsJson) {
  var ids = parseDriveIdsFromText_(linksText);
  var seen = {};
  var merged = [];
  for (var i = 0; i < ids.length; i++) {
    if (!seen[ids[i]]) {
      seen[ids[i]] = true;
      merged.push(ids[i]);
    }
  }
  try {
    var extra = JSON.parse(extraIdsJson || '[]');
    for (var j = 0; j < extra.length; j++) {
      var id = String(extra[j] || '').trim();
      if (id && !seen[id]) {
        seen[id] = true;
        merged.push(id);
      }
    }
  } catch (parseErr) {
    throw new Error('Некорректный список выбранных файлов.');
  }
  return merged;
}

/**
 * @param {Array<string>} ids
 * @returns {{ok: boolean, data?: {items: Array, entries: Array, resolve_errors: Array}, error?: string, message?: string}}
 */
function resolveDriveImportFromIds_(ids) {
  var items = [];
  var resolveErrors = [];
  for (var i = 0; i < ids.length; i++) {
    var resolved = resolveDriveImportItem_(ids[i]);
    if (resolved.ok) {
      items.push(resolved.data);
    } else {
      resolveErrors.push(resolved.message || ids[i]);
    }
  }

  if (!items.length) {
    return {
      ok: false,
      error: 'RESOLVE_FAILED',
      message: 'Не удалось открыть ни один объект на Drive:\n' + resolveErrors.join('\n')
    };
  }

  var entries = [];
  for (var j = 0; j < items.length; j++) {
    entries = entries.concat(expandDriveImportItem_(items[j]));
  }
  if (!entries.length) {
    return {
      ok: false,
      error: 'EMPTY_PLAN',
      message: 'В указанных папках нет файлов для импорта.'
    };
  }

  return {
    ok: true,
    data: {
      items: items,
      entries: entries,
      resolve_errors: resolveErrors
    }
  };
}

/**
 * @param {Array} entries
 */
function annotateDriveImportEntries_(entries) {
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var canMove = isImporterOwnerOfDriveFileId_(entry.source_file_id);
    entry.can_move = canMove;
    entry.import_mode = 'copy';
  }
}

/**
 * @param {Array} entries
 * @returns {number}
 */
function countMovableDriveImportEntries_(entries) {
  var count = 0;
  for (var i = 0; i < entries.length; i++) {
    if (entries[i].can_move) {
      count++;
    }
  }
  return count;
}

/**
 * @param {string} fileId
 * @param {string=} ownerEmail
 * @returns {boolean}
 */
function isImporterOwnerOfDriveFileId_(fileId, ownerEmail) {
  if (!fileId) {
    return false;
  }
  var importerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (!importerEmail) {
    return false;
  }
  var owner = ownerEmail
    ? normalizeEmail_(ownerEmail)
    : getDriveFileOwnerEmail_(fileId);
  return !!(owner && owner === importerEmail);
}

/**
 * @param {Array} entries
 * @returns {{ok: boolean, message?: string, error?: string}}
 */
function grantCatalogOwnerReadAccessForImportEntries_(entries) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copySourceIds = {};
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    if (entry.import_mode === 'move') {
      continue;
    }
    if (entry.source_file_id) {
      copySourceIds[entry.source_file_id] = true;
    }
  }
  var ids = Object.keys(copySourceIds);
  if (!ids.length) {
    return { ok: true, message: 'Доступ для копирования не требуется.' };
  }

  var failed = [];
  for (var j = 0; j < ids.length; j++) {
    try {
      grantCatalogOwnerReadAccess_(ids[j], ss);
    } catch (grantErr) {
      failed.push(ids[j] + ': ' + (grantErr.message || String(grantErr)));
    }
  }
  if (failed.length) {
    return {
      ok: false,
      error: 'GRANT_FAILED',
      message: failed.join('\n')
    };
  }
  return { ok: true };
}

/**
 * Проверяет ссылки и строит план (без копирования).
 * @param {string} linksText
 * @returns {{ok: boolean, data?: {entries: Array, count: number, items: Array}, error?: string, message?: string}}
 */
function planDriveImportFromLinks(linksText) {
  return prepareDriveImportFromSources(linksText, '[]');
}

/**
 * Шаг 1 импорта с Drive: от имени импортёра открыть чтение владельцу каталога.
 * @param {string} linksText
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function grantCatalogOwnerReadAccessForImport(linksText) {
  try {
    var resolved = resolveDriveImportFromIds_(collectDriveImportInputIds_(linksText, '[]'));
    if (!resolved.ok) {
      return resolved;
    }
    annotateDriveImportEntries_(resolved.data.entries);
    return grantCatalogOwnerReadAccessForImportEntries_(resolved.data.entries);
  } catch (err) {
    return {
      ok: false,
      error: 'GRANT_ACCESS_FAILED',
      message: String(err)
    };
  }
}

/**
 * @deprecated Используйте resolveDriveImportFromIds_
 */
function resolveDriveImportItemsFromLinks_(linksText) {
  return resolveDriveImportFromIds_(collectDriveImportInputIds_(linksText, '[]'));
}

/**
 * @param {string} text
 * @returns {Array<string>}
 */
function parseDriveIdsFromText_(text) {
  var lines = String(text || '').split(/\r?\n/);
  var ids = [];
  var seen = {};
  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || '').trim();
    if (!line) {
      continue;
    }
    var id = parseDriveIdFromLink_(line);
    if (id && !seen[id]) {
      seen[id] = true;
      ids.push(id);
    }
  }
  return ids;
}

/**
 * @param {string} text
 * @returns {string}
 */
function parseDriveIdFromLink_(text) {
  var value = String(text || '').trim();
  if (!value) {
    return '';
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(value) && value.indexOf('http') !== 0 && value.indexOf('/') === -1) {
    return value;
  }

  var patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = value.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }
  return '';
}

/**
 * @param {string} fileId
 * @returns {{ok: boolean, data?: {id: string, name: string, mimeType: string}, message?: string}}
 */
function resolveDriveImportItem_(fileId) {
  try {
    var file = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,trashed,shortcutDetails(targetId,targetMimeType)'
    });
    if (!file || !file.id) {
      return { ok: false, message: fileId + ': объект не найден.' };
    }
    if (file.trashed) {
      return { ok: false, message: (file.name || fileId) + ': объект в корзине Drive.' };
    }
    if (file.mimeType === DRIVE_SHORTCUT_MIME && file.shortcutDetails && file.shortcutDetails.targetId) {
      return resolveDriveImportItem_(file.shortcutDetails.targetId);
    }
    return {
      ok: true,
      data: {
        id: file.id,
        name: file.name || 'Файл',
        mimeType: file.mimeType || ''
      }
    };
  } catch (err) {
    return {
      ok: false,
      message: fileId + ': ' + String(err)
    };
  }
}

/**
 * @param {Object} item
 * @returns {Array}
 */
function expandDriveImportItem_(item) {
  if (!item || !item.id) {
    return [];
  }
  if (item.mimeType === DRIVE_FOLDER_MIME) {
    var label = sanitizeFolderLabel_(item.name || item.title || 'Папка');
    return listDriveImportEntries_(item.id, [label]);
  }
  if (item.mimeType === DRIVE_SHORTCUT_MIME) {
    return [];
  }
  return [{
    path_segments: [],
    file_name: item.name || item.title || 'Файл',
    source_file_id: item.id,
    mime_type: item.mimeType || ''
  }];
}

/**
 * @param {string} folderId
 * @param {Array<string>} pathSegments
 * @returns {Array}
 */
function listDriveImportEntries_(folderId, pathSegments) {
  var entries = [];
  var pageToken = null;

  do {
    var response = Drive.Files.list({
      q: "'" + folderId + "' in parents and trashed = false",
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'nextPageToken, files(id,name,mimeType)'
    });

    var batch = response.files || [];
    for (var i = 0; i < batch.length; i++) {
      var item = batch[i];
      if (item.mimeType === DRIVE_FOLDER_MIME) {
        var childLabel = sanitizeFolderLabel_(item.name);
        entries = entries.concat(
          listDriveImportEntries_(item.id, pathSegments.concat([childLabel]))
        );
        continue;
      }
      if (item.mimeType === DRIVE_SHORTCUT_MIME) {
        continue;
      }
      entries.push({
        path_segments: pathSegments,
        file_name: item.name,
        source_file_id: item.id,
        mime_type: item.mimeType || ''
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return entries;
}

/**
 * @param {string} parentId
 * @param {string} pathSegmentsJson
 * @param {string} fileName
 * @param {string} base64Data
 * @param {string} mimeType
 * @param {boolean} isLast
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function importCatalogLocalFile(parentId, pathSegmentsJson, fileName, base64Data, mimeType, isLast, resetOwnershipBatch) {
  try {
    if (resetOwnershipBatch) {
      resetImportOwnershipBatchPending_();
    }
    var pathSegments = JSON.parse(pathSegmentsJson || '[]');
    var result = importCatalogLocalFileCore_(parentId, pathSegments, fileName, base64Data, mimeType);
    if (isLast) {
      finishCatalogImport_(parentId);
      result.message = appendImportFinishMessage_(result.message || '');
    }
    return result;
  } catch (err) {
    return {
      ok: false,
      error: 'IMPORT_LOCAL_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} parentId
 * @param {string} entryJson
 * @param {boolean} isLast
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function importCatalogDriveEntry(parentId, entryJson, isLast) {
  try {
    var entry = JSON.parse(entryJson || '{}');
    var result = importCatalogDriveEntryCore_(parentId, entry);
    if (isLast) {
      finishCatalogImport_(parentId);
      result.message = appendImportFinishMessage_(result.message || '');
    }
    return result;
  } catch (err) {
    return {
      ok: false,
      error: 'IMPORT_DRIVE_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} parentId
 * @param {Object=} options
 * @returns {string}
 */
function finishCatalogImport_(parentId, options) {
  options = options || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);
  if (options.lightRender) {
    renderCatalogViewLight_(sheets.view, sheets.data);
  } else {
    renderCatalogViewCore_(sheets.view, sheets.data);
  }
  return parentId;
}

/**
 * @param {string} baseMessage
 * @returns {string}
 */
function appendImportFinishMessage_(baseMessage) {
  var pending = consumeImportOwnershipBatchPending_();
  var message = baseMessage + ' Импорт завершён.';
  if (!pending.length) {
    return message;
  }
  message += '\n\nВладение не передано владельцу каталога (' + pending.length + '):';
  for (var i = 0; i < pending.length; i++) {
    message += '\n• ' + pending[i].name + ' — ' + pending[i].message;
  }
  message += '\n\nИмпортёр должен оставаться владельцем до принятия. ' +
    'Владелец каталога проверьте почту Google (приглашение на владение).';
  return message;
}

/**
 * Сбрасывает накопленный отчёт о непереданном владении перед новым импортом.
 */
function resetImportOwnershipBatchPending_() {
  PropertiesService.getScriptProperties().deleteProperty(IMPORT_OWNER_BATCH_KEY);
}

/**
 * @param {{file_id: string, name: string, message: string}} item
 */
function appendImportOwnershipBatchPending_(item) {
  var raw = PropertiesService.getScriptProperties().getProperty(IMPORT_OWNER_BATCH_KEY);
  var list = raw ? JSON.parse(raw) : [];
  list.push(item);
  PropertiesService.getScriptProperties().setProperty(
    IMPORT_OWNER_BATCH_KEY,
    JSON.stringify(list)
  );
}

/**
 * @returns {Array<{file_id: string, name: string, message: string}>}
 */
function consumeImportOwnershipBatchPending_() {
  var raw = PropertiesService.getScriptProperties().getProperty(IMPORT_OWNER_BATCH_KEY);
  if (!raw) {
    return [];
  }
  PropertiesService.getScriptProperties().deleteProperty(IMPORT_OWNER_BATCH_KEY);
  try {
    return JSON.parse(raw) || [];
  } catch (err) {
    return [];
  }
}

/**
 * @param {string} parentId
 * @param {string} entriesJson
 * @param {string=} batchOptionsJson JSON: { is_first?, is_last?, total_count?, done_offset? }
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function importCatalogDriveBatch(parentId, entriesJson, batchOptionsJson) {
  try {
    var batchOptions = {};
    if (batchOptionsJson) {
      try {
        batchOptions = JSON.parse(batchOptionsJson) || {};
      } catch (parseErr) {
        batchOptions = {};
      }
    }
    var isFirst = batchOptions.is_first !== false;
    var isLast = batchOptions.is_last !== false;
    if (isFirst) {
      resetImportOwnershipBatchPending_();
    }
    var entries = JSON.parse(entriesJson || '[]');
    var result = importCatalogDriveBatchCore_(parentId, entries, {
      isLast: isLast,
      totalCount: batchOptions.total_count || entries.length,
      doneOffset: batchOptions.done_offset || 0
    });
    if (isLast) {
      result.message = appendImportFinishMessage_(result.message || '');
    }
    return result;
  } catch (err) {
    return {
      ok: false,
      error: 'IMPORT_DRIVE_BATCH_FAILED',
      message: String(err)
    };
  }
}

/**
 * Отладка: один файл, замеры ms по шагам, инкрементальная вставка в Catalog (без полной перерисовки).
 *
 * @param {string} parentId
 * @param {string} entryJson
 * @returns {{ok: boolean, data?: object, message?: string, error?: string}}
 */
function debugImportSingleDriveFile(parentId, entryJson) {
  var started = Date.now();
  var last = started;
  var timings = [];
  var mark = function (step) {
    var now = Date.now();
    timings.push({ step: step, ms: now - last });
    last = now;
  };
  try {
    var entry = JSON.parse(entryJson || '{}');
    if (!entry || !entry.source_file_id) {
      throw new Error('Некорректная запись импорта.');
    }
    entry.import_mode = entry.import_mode || 'copy';
    mark('parse');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!isSpreadsheetOwner_(ss)) {
      throw new Error('Только владелец каталога.');
    }
    var batchState = buildCatalogImportBatchState_(ss);
    mark('load_catalog_cache');

    if (entry.import_mode === 'move') {
      importCatalogDriveMoveBatchEntry_(batchState, parentId, entry);
    } else {
      var plan = prepareCopyImportPlanItem_(batchState, parentId, entry);
      mark('prepare_plan');
      executeCopyImportPlanSequential_(batchState, plan);
    }
    mark('drive_copy');

    if (!batchState.pendingRows.length) {
      throw new Error('Строка каталога не создана.');
    }
    appendCatalogDataRowsBatch_(
      batchState.ctx.sheets.data,
      batchState.ctx.dataCol,
      batchState.pendingRows
    );
    mark('write_catalog_data');

    if (batchState.permissionSyncJobs.length) {
      scheduleImportDriveSyncJobs_(batchState.permissionSyncJobs);
    }
    mark('schedule_acl');

    var fileRow = batchState.pendingRows[batchState.pendingRows.length - 1];
    var inserted = insertCatalogViewNodeAfterCreate_(
      batchState.ctx.sheets.view,
      batchState.ctx.sheets.data,
      fileRow.id,
      fileRow.parent_id
    );
    mark(inserted ? 'insert_view_row' : 'full_light_render');
    if (!inserted) {
      finishCatalogImport_(parentId, { lightRender: true });
    }

    timings.push({ step: 'total', ms: Date.now() - started });
    var lines = timings.map(function (t) {
      return t.step + ': ' + t.ms + ' ms';
    });
    return {
      ok: true,
      data: {
        imported: batchState.imported,
        timings: timings,
        view_mode: inserted ? 'incremental' : 'light_render'
      },
      message: lines.join('\n')
    };
  } catch (err) {
    timings.push({ step: 'error', ms: Date.now() - started });
    return {
      ok: false,
      error: 'DEBUG_IMPORT_FAILED',
      data: { timings: timings },
      message: String(err) + '\n\n' + timings.map(function (t) {
        return t.step + ': ' + t.ms + ' ms';
      }).join('\n')
    };
  }
}

/**
 * v1.4.2: импорт copy для владельца — фазы 1–2 в RPC, copy на Drive в фоне.
 *
 * @param {string} parentId
 * @param {string} entriesJson
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function startCatalogDriveImport(parentId, entriesJson) {
  var started = Date.now();
  var profile = { last: started, steps: [] };
  var mark = function (step) {
    var now = Date.now();
    profile.steps.push({ step: step, ms: now - profile.last });
    profile.last = now;
  };
  try {
    assertCatalogImportNotBusy_();
    var entries = JSON.parse(entriesJson || '[]');
    if (!entries.length) {
      throw new Error('Нет файлов для импорта.');
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!isSpreadsheetOwner_(ss)) {
      throw new Error('Трёхфазный импорт доступен только владельцу каталога.');
    }
    mark('parse');

    var copyEntries = [];
    var moveEntries = [];
    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      if (!entry || !entry.source_file_id) {
        continue;
      }
      entry.import_mode = entry.import_mode || 'copy';
      if (entry.import_mode === 'move') {
        moveEntries.push(entry);
      } else {
        copyEntries.push(entry);
      }
    }
    if (!copyEntries.length && !moveEntries.length) {
      throw new Error('Нет корректных записей для импорта.');
    }

    resetImportOwnershipBatchPending_();
    var metadataById = scanDriveImportFileMetadata_(collectImportSourceFileIds_(entries));
    mark('scan_drive');

    for (var m = 0; m < copyEntries.length; m++) {
      var meta = metadataById[copyEntries[m].source_file_id];
      if (meta && meta.name) {
        copyEntries[m].file_name = meta.name;
      }
      if (meta && meta.mimeType) {
        copyEntries[m].mime_type = meta.mimeType;
      }
    }
    for (var mv = 0; mv < moveEntries.length; mv++) {
      var moveMeta = metadataById[moveEntries[m].source_file_id];
      if (moveMeta && moveMeta.name) {
        moveEntries[m].file_name = moveMeta.name;
      }
      if (moveMeta && moveMeta.mimeType) {
        moveEntries[m].mime_type = moveMeta.mimeType;
      }
    }

    var batchState = buildCatalogImportBatchState_(ss);
    mark('load_catalog');

    var backgroundJobs = [];
    var moved = [];
    var failed = [];

    for (var c = 0; c < copyEntries.length; c++) {
      try {
        var plan = prepareCopyImportPlanItem_(batchState, parentId, copyEntries[c]);
        var catalogId = registerCatalogPendingFileRowCached_(
          batchState,
          plan.targetParent,
          plan.uniqueName,
          copyEntries[c].mime_type || '',
          plan.fileFolderCode
        );
        backgroundJobs.push({
          catalog_id: catalogId,
          source_file_id: copyEntries[c].source_file_id,
          drive_name: plan.uniqueName,
          mime_type: copyEntries[c].mime_type || '',
          permissions_json: batchState.pendingImportMeta[catalogId] || '[]',
          status: 'pending',
          error: ''
        });
      } catch (copyPlanErr) {
        failed.push({
          name: copyEntries[c].file_name || copyEntries[c].source_file_id,
          message: String(copyPlanErr)
        });
      }
    }

    for (var mm = 0; mm < moveEntries.length; mm++) {
      try {
        importCatalogDriveMoveBatchEntry_(batchState, parentId, moveEntries[mm]);
        if (batchState.imported.length) {
          moved.push(batchState.imported[batchState.imported.length - 1]);
        }
      } catch (moveErr) {
        failed.push({
          name: moveEntries[mm].file_name || moveEntries[mm].source_file_id,
          message: String(moveErr)
        });
      }
    }

    var folderRows = [];
    var fileRows = [];
    for (var r = 0; r < batchState.pendingRows.length; r++) {
      if (batchState.pendingRows[r].type === 'folder') {
        folderRows.push(batchState.pendingRows[r]);
      } else {
        fileRows.push(batchState.pendingRows[r]);
      }
    }

    if (folderRows.length) {
      appendCatalogDataRowsBatch_(batchState.ctx.sheets.data, batchState.ctx.dataCol, folderRows);
    }
    mark('prepare');

    var dataStartRow = 0;
    if (fileRows.length) {
      dataStartRow = batchState.ctx.sheets.data.getLastRow() + 1;
      appendCatalogDataRowsBatch_(batchState.ctx.sheets.data, batchState.ctx.dataCol, fileRows);
      assignImportJobDataRowIndexes_(backgroundJobs, fileRows, dataStartRow);
    }
    mark('write_catalog');

    var viewMode = 'deferred';
    if (fileRows.length && fileRows.length <= IMPORT_VIEW_DEFER_THRESHOLD) {
      var insertedAll = true;
      for (var vi = 0; vi < fileRows.length; vi++) {
        if (!insertCatalogViewNodeAfterCreate_(
          batchState.ctx.sheets.view,
          batchState.ctx.sheets.data,
          fileRows[vi].id,
          fileRows[vi].parent_id
        )) {
          insertedAll = false;
          break;
        }
      }
      viewMode = insertedAll ? 'incremental' : 'deferred';
    } else if (moved.length && !backgroundJobs.length) {
      finishCatalogImport_(parentId, { lightRender: true });
      viewMode = 'light_render';
    }
    mark('view_' + viewMode);

    if (backgroundJobs.length) {
      saveDriveImportJobQueue_(backgroundJobs);
      setImportJobMeta_({
        total: backgroundJobs.length,
        done: 0,
        failed: 0,
        started_at: new Date().toISOString()
      });
      setImportInProgress_(true);
      updateImportBatchProgress_({
        done: 0,
        total: backgroundJobs.length,
        phase: 'background',
        message: 'Копирование на Drive…'
      });
      ensureDriveImportJobTrigger_();
    }
    mark('schedule_background');

    var totalMs = Date.now() - started;
    profile.steps.push({ step: 'total', ms: totalMs });
    var timingLines = profile.steps.map(function (item) {
      return item.step + ': ' + item.ms + ' ms';
    });
    var messageParts = timingLines.slice();
    if (fileRows.length) {
      messageParts.push('');
      messageParts.push('В каталог записано: ' + fileRows.length + ' файл(ов).');
      if (backgroundJobs.length) {
        messageParts.push('Копирование на Drive — в фоне (~' +
          estimateDriveImportSeconds_(backgroundJobs.length) + ' с).');
        messageParts.push('Другие операции каталога заблокированы до завершения.');
      }
    }
    if (moved.length) {
      messageParts.push('Перенесено: ' + moved.length + '.');
    }
    if (failed.length) {
      messageParts.push('Ошибок планирования: ' + failed.length + '.');
    }

    return {
      ok: failed.length === 0,
      data: {
        catalog_rows: fileRows.length,
        background_jobs: backgroundJobs.length,
        moved: moved,
        failed: failed,
        view_mode: viewMode,
        timings: profile.steps,
        background: backgroundJobs.length > 0
      },
      message: messageParts.join('\n')
    };
  } catch (err) {
    profile.steps.push({ step: 'error', ms: Date.now() - started });
    return {
      ok: false,
      error: 'START_DRIVE_IMPORT_FAILED',
      data: { timings: profile.steps },
      message: String(err)
    };
  }
}

/**
 * @param {Array} entries
 * @returns {string[]}
 */
function collectImportSourceFileIds_(entries) {
  var ids = [];
  var seen = {};
  for (var i = 0; i < entries.length; i++) {
    var id = entries[i] && entries[i].source_file_id;
    if (id && !seen[id]) {
      seen[id] = true;
      ids.push(id);
    }
  }
  return ids;
}

/**
 * @param {string[]} fileIds
 * @returns {Object.<string, {id: string, name: string, mimeType: string}>}
 */
function scanDriveImportFileMetadata_(fileIds) {
  var map = {};
  if (!fileIds || !fileIds.length) {
    return map;
  }
  for (var offset = 0; offset < fileIds.length; offset += DRIVE_BATCH_MAX) {
    var slice = fileIds.slice(offset, offset + DRIVE_BATCH_MAX);
    var requests = buildDriveMetadataBatchGetRequests_(slice);
    var results = executeDriveBatchRequests_(requests);
    for (var i = 0; i < slice.length; i++) {
      var result = findDriveBatchResultByContentId_(results, String(i + 1));
      if (result && result.ok && result.data && result.data.id) {
        map[slice[i]] = {
          id: result.data.id,
          name: result.data.name || '',
          mimeType: result.data.mimeType || ''
        };
      }
    }
  }
  return map;
}

/**
 * @param {string[]} fileIds
 * @returns {Array<{contentId: string, method: string, path: string}>}
 */
function buildDriveMetadataBatchGetRequests_(fileIds) {
  var requests = [];
  for (var i = 0; i < fileIds.length; i++) {
    requests.push({
      contentId: String(i + 1),
      method: 'GET',
      path: '/drive/v3/files/' + encodeURIComponent(fileIds[i]) +
        '?supportsAllDrives=true&fields=' + encodeURIComponent('id,name,mimeType')
    });
  }
  return requests;
}

/**
 * @param {Object} batchState
 * @param {Object} parent
 * @param {string} fullName
 * @param {string} mimeType
 * @param {string} fileFolderCode
 * @returns {string}
 */
function registerCatalogPendingFileRowCached_(batchState, parent, fullName, mimeType, fileFolderCode) {
  if (!batchState.pendingImportMeta) {
    batchState.pendingImportMeta = {};
  }
  var inherited = inheritPermissionsFromParent_(parent, batchState.ctx.usersMap);
  var newNodeId = generateId_();
  var row = {
    id: newNodeId,
    file_id: '',
    name: fullName,
    type: 'file',
    folder_code: fileFolderCode,
    parent_id: parent.id,
    mime_type: mimeType || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  };
  batchState.pendingRows.push(row);
  batchState.pendingImportMeta[newNodeId] = inherited.permissions_json;
  var node = {
    id: newNodeId,
    file_id: '',
    name: fullName,
    type: 'file',
    folder_code: fileFolderCode,
    parent_id: parent.id,
    mime_type: mimeType || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers,
    rowIndex: 0
  };
  batchState.allNodes.push(node);
  batchState.nodesById[newNodeId] = node;
  return newNodeId;
}

/**
 * @param {Array} jobs
 * @param {Array} fileRows
 * @param {number} dataStartRow
 */
function assignImportJobDataRowIndexes_(jobs, fileRows, dataStartRow) {
  for (var i = 0; i < fileRows.length; i++) {
    if (fileRows[i].file_id) {
      continue;
    }
    var rowIndex = dataStartRow + i;
    for (var j = 0; j < jobs.length; j++) {
      if (jobs[j].catalog_id === fileRows[i].id) {
        jobs[j].data_row_index = rowIndex;
        break;
      }
    }
  }
}

/**
 * @param {number} fileCount
 * @returns {number}
 */
function estimateDriveImportSeconds_(fileCount) {
  if (fileCount <= 1) {
    return 3;
  }
  if (fileCount <= DRIVE_BATCH_MAX) {
    return 30;
  }
  return Math.ceil(fileCount / DRIVE_BATCH_MAX) * 30;
}

function assertCatalogImportNotBusy_() {
  if (!isImportInProgress_()) {
    return;
  }
  var meta = getImportJobMeta_();
  var done = meta.done || 0;
  var total = meta.total || 0;
  throw new Error(
    'Идёт фоновый импорт (' + done + ' / ' + total + '). ' +
    'Дождитесь завершения или «Служебные → Сбросить зависший импорт».'
  );
}

function isImportInProgress_() {
  return PropertiesService.getDocumentProperties().getProperty(IMPORT_IN_PROGRESS_KEY) === '1';
}

/**
 * @param {boolean} active
 */
function setImportInProgress_(active) {
  var props = PropertiesService.getDocumentProperties();
  if (active) {
    props.setProperty(IMPORT_IN_PROGRESS_KEY, '1');
  } else {
    props.deleteProperty(IMPORT_IN_PROGRESS_KEY);
  }
}

/**
 * @param {Object} meta
 */
function setImportJobMeta_(meta) {
  PropertiesService.getDocumentProperties().setProperty(
    IMPORT_JOB_META_KEY,
    JSON.stringify(meta || {})
  );
}

/**
 * @returns {Object}
 */
function getImportJobMeta_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(IMPORT_JOB_META_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (err) {
    return {};
  }
}

/**
 * @param {Array} jobs
 */
function saveDriveImportJobQueue_(jobs) {
  PropertiesService.getDocumentProperties().setProperty(
    PENDING_DRIVE_IMPORT_JOB_KEY,
    JSON.stringify(jobs || [])
  );
}

/**
 * @returns {Array}
 */
function loadDriveImportJobQueue_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(PENDING_DRIVE_IMPORT_JOB_KEY);
    if (!raw) {
      return [];
    }
    var parsed = JSON.parse(raw);
    return parsed && parsed.length ? parsed : [];
  } catch (err) {
    Logger.log('loadDriveImportJobQueue_: ' + err);
    return [];
  }
}

function clearDriveImportJobState_() {
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty(PENDING_DRIVE_IMPORT_JOB_KEY);
  props.deleteProperty(IMPORT_JOB_META_KEY);
  setImportInProgress_(false);
  clearImportBatchProgress_();
}

function ensureDriveImportJobTrigger_() {
  deleteDriveImportJobTriggers_();
  ScriptApp.newTrigger(IMPORT_JOB_HANDLER)
    .timeBased()
    .after(1500)
    .create();
}

function deleteDriveImportJobTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === IMPORT_JOB_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function processPendingDriveImportJob_() {
  deleteDriveImportJobTriggers_();
  if (!isImportInProgress_()) {
    return;
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    ensureDriveImportJobTrigger_();
    return;
  }
  try {
    processPendingDriveImportJobCore_();
  } finally {
    lock.releaseLock();
  }
}

function processPendingDriveImportJobCore_() {
  var queue = loadDriveImportJobQueue_();
  var pending = [];
  for (var i = 0; i < queue.length; i++) {
    if (queue[i] && queue[i].status === 'pending') {
      pending.push(queue[i]);
    }
  }
  if (!pending.length) {
    finalizeDriveImportJob_();
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ctx = buildCatalogImportContextLight_(ss);
  var chunk = pending.slice(0, DRIVE_BATCH_MAX);
  var requests = buildDriveCopyBatchRequestsFromJobs_(chunk, ctx.rootFolderId);
  var results;
  try {
    results = executeDriveBatchRequests_(requests);
  } catch (batchErr) {
    Logger.log('processPendingDriveImportJobCore_ batch: ' + batchErr);
    for (var f = 0; f < chunk.length; f++) {
      executeDriveImportJobSequential_(ctx, chunk[f], queue);
    }
    saveDriveImportJobQueue_(queue);
    updateDriveImportJobProgress_(queue);
    ensureDriveImportJobTrigger_();
    return;
  }

  for (var j = 0; j < chunk.length; j++) {
    var job = chunk[j];
    var result = findDriveBatchResultByContentId_(results, String(j + 1));
    if (!result || !result.ok || !result.data || !result.data.id) {
      markDriveImportJobFailed_(queue, job.catalog_id, result && result.error
        ? result.error
        : 'Не удалось скопировать файл.');
      continue;
    }
    try {
      applyDriveImportJobCopyResult_(ctx, job, result.data, queue);
    } catch (applyErr) {
      try {
        Drive.Files.remove(result.data.id, { supportsAllDrives: true });
      } catch (removeErr) {
        Logger.log('processPendingDriveImportJobCore_ rollback: ' + removeErr);
      }
      markDriveImportJobFailed_(queue, job.catalog_id, String(applyErr));
    }
  }

  saveDriveImportJobQueue_(queue);
  updateDriveImportJobProgress_(queue);

  var stillPending = false;
  for (var p = 0; p < queue.length; p++) {
    if (queue[p] && queue[p].status === 'pending') {
      stillPending = true;
      break;
    }
  }
  if (stillPending) {
    ensureDriveImportJobTrigger_();
    return;
  }
  finalizeDriveImportJob_();
}

/**
 * @param {Array<{catalog_id: string, source_file_id: string, drive_name: string}>} jobs
 * @param {string} rootFolderId
 * @returns {Array}
 */
function buildDriveCopyBatchRequestsFromJobs_(jobs, rootFolderId) {
  var requests = [];
  for (var i = 0; i < jobs.length; i++) {
    var job = jobs[i];
    requests.push({
      contentId: String(i + 1),
      method: 'POST',
      path: '/drive/v3/files/' + encodeURIComponent(job.source_file_id) +
        '/copy?supportsAllDrives=true&fields=' + encodeURIComponent('id,name,mimeType'),
      body: {
        name: job.drive_name,
        parents: [rootFolderId]
      }
    });
  }
  return requests;
}

/**
 * @param {Object} ctx
 * @param {Object} job
 * @param {Array} queue
 */
function executeDriveImportJobSequential_(ctx, job, queue) {
  try {
    var copied = copyDriveFileDirectToCatalogRootForOwner_(
      job.source_file_id,
      ctx.rootFolderId,
      job.drive_name
    );
    applyDriveImportJobCopyResult_(ctx, job, copied, queue);
  } catch (err) {
    markDriveImportJobFailed_(queue, job.catalog_id, String(err));
  }
}

/**
 * @param {Object} ctx
 * @param {Object} job
 * @param {Object} driveFile
 * @param {Array} queue
 */
function applyDriveImportJobCopyResult_(ctx, job, driveFile, queue) {
  if (!job.data_row_index) {
    throw new Error('Не найдена строка CatalogData для ' + job.catalog_id);
  }
  ctx.sheets.data.getRange(job.data_row_index, ctx.dataCol.file_id).setValue(driveFile.id);
  if (driveFile.mimeType || job.mime_type) {
    ctx.sheets.data.getRange(job.data_row_index, ctx.dataCol.mime_type)
      .setValue(driveFile.mimeType || job.mime_type || '');
  }
  scheduleImportDriveSyncJobs_([{
    file_id: driveFile.id,
    permissions_json: job.permissions_json || '[]'
  }]);
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].catalog_id === job.catalog_id) {
      queue[i].status = 'done';
      queue[i].file_id = driveFile.id;
      queue[i].error = '';
      break;
    }
  }
}

/**
 * @param {Array} queue
 * @param {string} catalogId
 * @param {string} message
 */
function markDriveImportJobFailed_(queue, catalogId, message) {
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].catalog_id === catalogId) {
      queue[i].status = 'failed';
      queue[i].error = message;
      break;
    }
  }
}

/**
 * @param {Array} queue
 */
function updateDriveImportJobProgress_(queue) {
  var done = 0;
  var failed = 0;
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].status === 'done') {
      done++;
    } else if (queue[i].status === 'failed') {
      failed++;
    }
  }
  setImportJobMeta_({
    total: queue.length,
    done: done,
    failed: failed,
    started_at: getImportJobMeta_().started_at || ''
  });
  updateImportBatchProgress_({
    done: done,
    total: queue.length,
    phase: done + failed >= queue.length ? 'done' : 'background',
    message: 'Копирование на Drive…'
  });
}

function finalizeDriveImportJob_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);
  finishCatalogImport_(null, { lightRender: true });
  clearDriveImportJobState_();
}

function resetStuckCatalogImportMenu_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isSpreadsheetOwner_(ss)) {
    SpreadsheetApp.getUi().alert('Сброс импорта', 'Только владелец каталога.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Сброс импорта',
    'Снять блокировку и очистить очередь фонового импорта?\n' +
    'Строки каталога с пустым file_id останутся — удалите их вручную при необходимости.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    return;
  }
  deleteDriveImportJobTriggers_();
  clearDriveImportJobState_();
  ui.alert('Сброс импорта', 'Блокировка снята.', ui.ButtonSet.OK);
}

/**
 * @returns {{ok: boolean, data?: object, error?: string}}
 */
function getImportBatchProgress() {
  try {
    var raw = CacheService.getDocumentCache().get(IMPORT_BATCH_PROGRESS_KEY);
    if (!raw) {
      return { ok: true, data: { done: 0, total: 0, phase: 'idle' } };
    }
    return { ok: true, data: JSON.parse(raw) };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * @param {number} total
 */
function resetImportBatchProgress_(total) {
  CacheService.getDocumentCache().put(
    IMPORT_BATCH_PROGRESS_KEY,
    JSON.stringify({
      done: 0,
      total: total || 0,
      phase: 'prepare',
      message: 'Подготовка…'
    }),
    600
  );
}

function clearImportBatchProgress_() {
  CacheService.getDocumentCache().remove(IMPORT_BATCH_PROGRESS_KEY);
}

/**
 * @param {Object} patch
 */
function updateImportBatchProgress_(patch) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(IMPORT_BATCH_PROGRESS_KEY);
  var current = raw ? JSON.parse(raw) : { done: 0, total: 0, phase: 'prepare' };
  var keys = Object.keys(patch || {});
  for (var i = 0; i < keys.length; i++) {
    current[keys[i]] = patch[keys[i]];
  }
  cache.put(IMPORT_BATCH_PROGRESS_KEY, JSON.stringify(current), 600);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Object}
 */
function buildCatalogImportBatchState_(ss) {
  var ctx = buildCatalogImportContextLight_(ss);
  var allNodes = readCatalogDataNodes_(ctx.sheets.data, ctx.dataCol);
  var nodesById = {};
  for (var i = 0; i < allNodes.length; i++) {
    nodesById[allNodes[i].id] = allNodes[i];
  }
  return {
    ctx: ctx,
    allNodes: allNodes,
    nodesById: nodesById,
    reservedNames: {},
    pendingRows: [],
    permissionSyncJobs: [],
    imported: [],
    failed: []
  };
}

/**
 * @param {Object} parent
 * @param {Array} allNodes
 * @returns {string}
 */
function allocateFolderCodeFromNodes_(parent, allNodes) {
  var parentCode = readCell_(parent.folder_code);
  var siblingCodes = [];
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.parent_id === parent.id && node.type === 'folder' && node.folder_code) {
      siblingCodes.push(node.folder_code);
    }
  }
  var used = folderCodesToSet_(siblingCodes);

  if (parentCode === FOLDER_CODE_CATALOG || parentCode === '') {
    for (var n = FOLDER_CODE_ROOT_MIN; n <= FOLDER_CODE_ROOT_MAX; n++) {
      var rootCode = String(n);
      if (!used[rootCode]) {
        return rootCode;
      }
    }
    throw new Error(
      'Лимит папок в корне исчерпан (' + FOLDER_CODE_ROOT_MIN + '–' + FOLDER_CODE_ROOT_MAX + ').'
    );
  }

  for (var d = 1; d <= 9; d++) {
    var childCode = parentCode + String(d);
    if (!used[childCode]) {
      return childCode;
    }
  }
  throw new Error('Лимит подпапок исчерпан (9).');
}

/**
 * @param {Object} batchState
 * @param {string} nodeId
 * @returns {Object|null}
 */
function findCatalogNodeInBatchState_(batchState, nodeId) {
  return batchState.nodesById[nodeId] || null;
}

/**
 * @param {Object} batchState
 * @param {string} anchorParentId
 * @param {Array<string>} pathSegments
 * @returns {Object}
 */
function ensureImportFolderPathCached_(batchState, anchorParentId, pathSegments) {
  var parent = findCatalogNodeInBatchState_(batchState, anchorParentId);
  validateCatalogImportParent_(parent);
  var currentParent = parent;

  for (var i = 0; i < pathSegments.length; i++) {
    var segment = sanitizeFolderLabel_(pathSegments[i]);
    if (!segment) {
      continue;
    }
    var existing = findChildFolderByLabel_(currentParent.id, segment, batchState.allNodes);
    if (existing) {
      currentParent = existing;
      continue;
    }
    currentParent = createImportVirtualFolderCached_(batchState, currentParent, segment);
  }

  return currentParent;
}

/**
 * @param {Object} batchState
 * @param {Object} parent
 * @param {string} label
 * @returns {Object}
 */
function createImportVirtualFolderCached_(batchState, parent, label) {
  assertCanAddSubfolder_(readCell_(parent.folder_code));
  var folderCode = allocateFolderCodeFromNodes_(parent, batchState.allNodes);
  var fullName = buildFolderName_(folderCode, label);
  var inherited = inheritPermissionsFromParent_(parent, batchState.ctx.usersMap);
  var newNodeId = generateId_();
  var row = {
    id: newNodeId,
    name: fullName,
    type: 'folder',
    folder_code: folderCode,
    parent_id: parent.id,
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  };
  batchState.pendingRows.push(row);
  var node = {
    id: newNodeId,
    name: fullName,
    type: 'folder',
    folder_code: folderCode,
    parent_id: parent.id,
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers,
    rowIndex: 0
  };
  batchState.allNodes.push(node);
  batchState.nodesById[newNodeId] = node;
  return node;
}

/**
 * @param {Object} batchState
 * @param {Object} parent
 * @param {string} fileId
 * @param {string} fullName
 * @param {string} mimeType
 * @param {string} fileFolderCode
 * @returns {string}
 */
function registerCatalogDriveFileRowCached_(batchState, parent, fileId, fullName, mimeType, fileFolderCode) {
  var inherited = inheritPermissionsFromParent_(parent, batchState.ctx.usersMap);
  var newNodeId = generateId_();
  var row = {
    id: newNodeId,
    file_id: fileId,
    name: fullName,
    type: 'file',
    folder_code: fileFolderCode,
    parent_id: parent.id,
    mime_type: mimeType || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  };
  batchState.pendingRows.push(row);
  var node = {
    id: newNodeId,
    file_id: fileId,
    name: fullName,
    type: 'file',
    folder_code: fileFolderCode,
    parent_id: parent.id,
    mime_type: mimeType || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers,
    rowIndex: 0
  };
  batchState.allNodes.push(node);
  batchState.nodesById[newNodeId] = node;
  batchState.permissionSyncJobs.push({
    file_id: fileId,
    permissions_json: inherited.permissions_json
  });
  return newNodeId;
}

/**
 * @param {string} sourceFileId
 * @param {string} rootFolderId
 * @param {string} finalName
 * @returns {Object}
 */
function copyDriveFileDirectToCatalogRootForOwner_(sourceFileId, rootFolderId, finalName) {
  var copied = Drive.Files.copy({
    name: finalName,
    parents: [rootFolderId]
  }, sourceFileId, {
    supportsAllDrives: true,
    fields: 'id,name,mimeType'
  });
  if (!copied || !copied.id) {
    throw new Error('Drive не вернул id копии.');
  }
  return copied;
}

/**
 * @param {Object} headers
 * @param {string} name
 * @returns {string}
 */
function getHttpResponseHeader_(headers, name) {
  var wanted = String(name || '').toLowerCase();
  var keys = Object.keys(headers || {});
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i]).toLowerCase() === wanted) {
      var val = headers[keys[i]];
      return Array.isArray(val) ? String(val[0] || '') : String(val || '');
    }
  }
  return '';
}

/**
 * @param {string} contentTypeHeader
 * @param {string} responseText
 * @returns {string}
 */
function extractMultipartBoundary_(contentTypeHeader, responseText) {
  var match = /boundary=([^;\s]+)/i.exec(String(contentTypeHeader || ''));
  if (match && match[1]) {
    return match[1].replace(/"/g, '');
  }
  var bodyMatch = /--(batch[^\r\n]+)/.exec(String(responseText || ''));
  return bodyMatch ? bodyMatch[1] : '';
}

/**
 * @param {Array<{contentId: string, method: string, path: string, body?: Object}>} requests
 * @returns {Array<{contentId: string, ok: boolean, status: number, data?: Object, error?: string}>}
 */
function executeDriveBatchRequests_(requests) {
  if (!requests || !requests.length) {
    return [];
  }
  var boundary = 'batch_catalog_' + Utilities.getUuid().replace(/-/g, '');
  var body = buildDriveBatchHttpBody_(requests, boundary);
  var response = UrlFetchApp.fetch(DRIVE_BATCH_URL, {
    method: 'post',
    contentType: 'multipart/mixed; boundary=' + boundary,
    payload: body,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Drive batch HTTP ' + code + ': ' + response.getContentText().substring(0, 400));
  }
  return parseDriveBatchResponse_(
    response.getContentText(),
    getHttpResponseHeader_(response.getAllHeaders(), 'Content-Type')
  );
}

/**
 * @param {Array<{contentId: string, method: string, path: string, body?: Object}>} requests
 * @param {string} boundary
 * @returns {string}
 */
function buildDriveBatchHttpBody_(requests, boundary) {
  var chunks = [];
  for (var i = 0; i < requests.length; i++) {
    var req = requests[i];
    chunks.push('--' + boundary);
    chunks.push('Content-Type: application/http');
    chunks.push('Content-ID: <' + req.contentId + '>');
    chunks.push('');
    chunks.push(req.method + ' ' + req.path);
    chunks.push('Content-Type: application/json');
    if (req.body) {
      chunks.push('');
      chunks.push(JSON.stringify(req.body));
    }
    chunks.push('');
  }
  chunks.push('--' + boundary + '--');
  return chunks.join('\r\n');
}

/**
 * @param {string} responseText
 * @param {string} contentTypeHeader
 * @returns {Array<{contentId: string, ok: boolean, status: number, data?: Object, error?: string}>}
 */
function parseDriveBatchResponse_(responseText, contentTypeHeader) {
  var boundary = extractMultipartBoundary_(contentTypeHeader, responseText);
  if (!boundary) {
    throw new Error('Не удалось разобрать boundary ответа Drive batch.');
  }
  var results = [];
  var parts = String(responseText || '').split('--' + boundary);
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];
    if (!part || part === '--' || part.indexOf('HTTP/') === -1) {
      continue;
    }
    var contentIdMatch = /Content-ID:\s*<?([^>\r\n]+)>?/i.exec(part);
    var statusMatch = /HTTP\/[\d.]+\s+(\d+)/.exec(part);
    var status = statusMatch ? parseInt(statusMatch[1], 10) : 0;
    var contentId = contentIdMatch ? String(contentIdMatch[1]).trim() : String(results.length + 1);
    var bodyChunks = part.split('\r\n\r\n');
    var rawBody = bodyChunks.length ? bodyChunks[bodyChunks.length - 1].trim() : '';
    var jsonStart = rawBody.indexOf('{');
    if (jsonStart > 0) {
      rawBody = rawBody.substring(jsonStart);
    }
    if (status >= 200 && status < 300 && rawBody.indexOf('{') === 0) {
      try {
        results.push({
          contentId: contentId,
          ok: true,
          status: status,
          data: JSON.parse(rawBody)
        });
        continue;
      } catch (parseErr) {
        results.push({
          contentId: contentId,
          ok: false,
          status: status,
          error: 'JSON: ' + parseErr
        });
        continue;
      }
    }
    results.push({
      contentId: contentId,
      ok: false,
      status: status,
      error: rawBody.substring(0, 300) || ('HTTP ' + status)
    });
  }
  return results;
}

/**
 * @param {Array} results
 * @param {string} contentId
 * @returns {Object|null}
 */
function findDriveBatchResultByContentId_(results, contentId) {
  for (var i = 0; i < results.length; i++) {
    if (String(results[i].contentId) === String(contentId)) {
      return results[i];
    }
  }
  return null;
}

/**
 * @param {Array<{entry: Object, targetParent: Object, fileFolderCode: string, uniqueName: string}>} copyPlans
 * @param {string} rootFolderId
 * @returns {Array<{contentId: string, method: string, path: string, body: Object}>}
 */
function buildDriveCopyBatchRequests_(copyPlans, rootFolderId) {
  var requests = [];
  for (var i = 0; i < copyPlans.length; i++) {
    var plan = copyPlans[i];
    requests.push({
      contentId: String(i + 1),
      method: 'POST',
      path: '/drive/v3/files/' + encodeURIComponent(plan.entry.source_file_id) +
        '/copy?supportsAllDrives=true&fields=' + encodeURIComponent('id,name,mimeType'),
      body: {
        name: plan.uniqueName,
        parents: [rootFolderId]
      }
    });
  }
  return requests;
}

/**
 * @param {Object} batchState
 * @param {string} parentId
 * @param {Object} entry
 * @returns {{entry: Object, targetParent: Object, fileFolderCode: string, uniqueName: string}}
 */
function prepareCopyImportPlanItem_(batchState, parentId, entry) {
  var targetParent = ensureImportFolderPathCached_(
    batchState,
    parentId,
    entry.path_segments || []
  );
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    sanitizeFolderLabel_(entry.file_name),
    batchState.allNodes,
    {},
    batchState.reservedNames
  );
  return {
    entry: entry,
    targetParent: targetParent,
    fileFolderCode: fileFolderCode,
    uniqueName: unique.name
  };
}

/**
 * @param {Object} batchState
 * @param {{entry: Object, targetParent: Object, fileFolderCode: string, uniqueName: string}} plan
 * @param {Object} driveFile
 */
function applyDriveCopyPlanResult_(batchState, plan, driveFile) {
  registerCatalogDriveFileRowCached_(
    batchState,
    plan.targetParent,
    driveFile.id,
    plan.uniqueName,
    driveFile.mimeType || plan.entry.mime_type || '',
    plan.fileFolderCode
  );
  batchState.imported.push({
    name: plan.uniqueName,
    file_id: driveFile.id,
    mode: 'copy'
  });
}

/**
 * @param {Object} batchState
 * @param {{entry: Object, targetParent: Object, fileFolderCode: string, uniqueName: string}} plan
 * @param {string} message
 */
function failCopyImportPlan_(batchState, plan, message) {
  batchState.failed.push({
    name: plan.entry.file_name || plan.entry.source_file_id,
    message: message
  });
}

/**
 * @param {Object} batchState
 * @param {{entry: Object, targetParent: Object, fileFolderCode: string, uniqueName: string}} plan
 */
function executeCopyImportPlanSequential_(batchState, plan) {
  var driveFile = copyDriveFileDirectToCatalogRootForOwner_(
    plan.entry.source_file_id,
    batchState.ctx.rootFolderId,
    plan.uniqueName
  );
  applyDriveCopyPlanResult_(batchState, plan, driveFile);
}

/**
 * @param {Object} batchState
 * @param {Array} copyPlans
 * @param {number} totalCount
 */
function executeDriveCopyPlansInBatches_(batchState, copyPlans, totalCount) {
  if (!copyPlans.length) {
    return;
  }
  if (copyPlans.length === 1) {
    try {
      executeCopyImportPlanSequential_(batchState, copyPlans[0]);
    } catch (seqErr) {
      failCopyImportPlan_(batchState, copyPlans[0], String(seqErr));
    }
    return;
  }

  for (var offset = 0; offset < copyPlans.length; offset += DRIVE_BATCH_MAX) {
    var slice = copyPlans.slice(offset, offset + DRIVE_BATCH_MAX);
    var requests = buildDriveCopyBatchRequests_(slice, batchState.ctx.rootFolderId);
    var results;
    try {
      results = executeDriveBatchRequests_(requests);
    } catch (batchErr) {
      Logger.log('executeDriveCopyPlansInBatches_ batch fallback: ' + batchErr);
      for (var f = 0; f < slice.length; f++) {
        try {
          executeCopyImportPlanSequential_(batchState, slice[f]);
        } catch (seqErr) {
          failCopyImportPlan_(batchState, slice[f], String(seqErr));
        }
      }
      continue;
    }

    for (var i = 0; i < slice.length; i++) {
      var plan = slice[i];
      var result = findDriveBatchResultByContentId_(results, String(i + 1));
      if (!result || !result.ok || !result.data || !result.data.id) {
        var errText = result && result.error ? result.error : 'Не удалось скопировать файл.';
        try {
          executeCopyImportPlanSequential_(batchState, plan);
        } catch (seqErr) {
          failCopyImportPlan_(batchState, plan, errText + ' ' + seqErr);
        }
        continue;
      }
      try {
        applyDriveCopyPlanResult_(batchState, plan, result.data);
      } catch (registerErr) {
        try {
          Drive.Files.remove(result.data.id, { supportsAllDrives: true });
        } catch (removeErr) {
          Logger.log('executeDriveCopyPlansInBatches_ rollback: ' + removeErr);
        }
        failCopyImportPlan_(batchState, plan, String(registerErr));
      }
    }
  }
}

/**
 * @param {Object} batchState
 * @param {string} parentId
 * @param {Object} entry
 */
function importCatalogDriveCopyBatchEntry_(batchState, parentId, entry) {
  var targetParent = ensureImportFolderPathCached_(
    batchState,
    parentId,
    entry.path_segments || []
  );
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    sanitizeFolderLabel_(entry.file_name),
    batchState.allNodes,
    {},
    batchState.reservedNames
  );

  var driveFile;
  try {
    driveFile = copyDriveFileDirectToCatalogRootForOwner_(
      entry.source_file_id,
      batchState.ctx.rootFolderId,
      unique.name
    );
  } catch (copyErr) {
    throw new Error('Не удалось скопировать файл на Drive: ' + copyErr);
  }

  try {
    registerCatalogDriveFileRowCached_(
      batchState,
      targetParent,
      driveFile.id,
      unique.name,
      driveFile.mimeType || entry.mime_type || '',
      fileFolderCode
    );
  } catch (permErr) {
    try {
      Drive.Files.remove(driveFile.id, { supportsAllDrives: true });
    } catch (removeErr) {
      Logger.log('importCatalogDriveCopyBatchEntry_ rollback: ' + removeErr);
    }
    throw new Error('Не удалось записать файл в каталог: ' + permErr);
  }

  batchState.imported.push({
    name: unique.name,
    file_id: driveFile.id,
    mode: 'copy'
  });
}

/**
 * @param {Object} batchState
 * @param {string} parentId
 * @param {Object} entry
 */
function importCatalogDriveMoveBatchEntry_(batchState, parentId, entry) {
  if (!isImporterOwnerOfDriveFileId_(entry.source_file_id)) {
    throw new Error('Перенос возможен только для файлов, где вы владелец на Drive.');
  }

  var targetParent = ensureImportFolderPathCached_(
    batchState,
    parentId,
    entry.path_segments || []
  );
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    sanitizeFolderLabel_(entry.file_name),
    batchState.allNodes,
    {},
    batchState.reservedNames
  );

  var fileMeta = Drive.Files.get(entry.source_file_id, {
    supportsAllDrives: true,
    fields: 'id,name,parents,mimeType'
  });
  if (!fileMeta || !fileMeta.id) {
    throw new Error('Файл не найден на Drive.');
  }

  var driveFile = moveDriveFileToCatalogRoot_(
    entry.source_file_id,
    batchState.ctx.rootFolderId,
    unique.name,
    fileMeta.parents || []
  );

  registerCatalogDriveFileRowCached_(
    batchState,
    targetParent,
    entry.source_file_id,
    unique.name,
    driveFile.mimeType || entry.mime_type || '',
    fileFolderCode
  );

  batchState.imported.push({
    name: unique.name,
    file_id: entry.source_file_id,
    mode: 'move'
  });
}

/**
 * @param {Object} batchState
 * @param {string} parentId
 * @param {Array} entries
 * @param {{isLast?: boolean}=} options
 * @returns {string}
 */
function finalizeBatchImportView_(batchState, parentId, entries, options) {
  options = options || {};
  if (options.isLast === false) {
    return 'deferred';
  }
  if (batchState.imported.length > 0 && batchState.imported.length <= 10) {
    var insertedAll = true;
    var insertedAny = false;
    for (var i = 0; i < batchState.pendingRows.length; i++) {
      var fileRow = batchState.pendingRows[i];
      if (fileRow.type !== 'file') {
        continue;
      }
      if (insertCatalogViewNodeAfterCreate_(
        batchState.ctx.sheets.view,
        batchState.ctx.sheets.data,
        fileRow.id,
        fileRow.parent_id
      )) {
        insertedAny = true;
      } else {
        insertedAll = false;
        break;
      }
    }
    if (insertedAny && insertedAll) {
      return 'incremental';
    }
  }
  finishCatalogImport_(parentId, { lightRender: true });
  return 'light_render';
}

/**
 * @param {string} parentId
 * @param {Array} entries
 * @param {{isLast?: boolean, totalCount?: number, doneOffset?: number}=} options
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogDriveBatchCore_(parentId, entries, options) {
  options = options || {};
  var profile = entries.length === 1 ? { last: Date.now(), steps: [] } : null;
  var mark = function (step) {
    if (!profile) {
      return;
    }
    var now = Date.now();
    profile.steps.push({ step: step, ms: now - profile.last });
    profile.last = now;
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isSpreadsheetOwner_(ss)) {
    throw new Error('Пакетный импорт доступен только владельцу каталога.');
  }
  if (!entries || !entries.length) {
    throw new Error('Нет файлов для импорта.');
  }

  var totalCount = options.totalCount || entries.length;
  var doneOffset = options.doneOffset || 0;
  mark('start');
  var batchState = buildCatalogImportBatchState_(ss);
  mark('load_catalog');
  var copyPlans = [];
  var moveEntries = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    if (!entry || !entry.source_file_id) {
      batchState.failed.push({
        name: entry && entry.file_name ? entry.file_name : '—',
        message: 'Некорректная запись импорта.'
      });
      continue;
    }
    try {
      if (entry.import_mode === 'move') {
        moveEntries.push(entry);
      } else {
        copyPlans.push(prepareCopyImportPlanItem_(batchState, parentId, entry));
      }
    } catch (err) {
      batchState.failed.push({
        name: entry.file_name || entry.source_file_id,
        message: String(err)
      });
    }
  }
  mark('prepare');

  executeDriveCopyPlansInBatches_(batchState, copyPlans, totalCount);
  mark('drive_copy');

  for (var m = 0; m < moveEntries.length; m++) {
    try {
      importCatalogDriveMoveBatchEntry_(batchState, parentId, moveEntries[m]);
    } catch (moveErr) {
      batchState.failed.push({
        name: moveEntries[m].file_name || moveEntries[m].source_file_id,
        message: String(moveErr)
      });
    }
  }

  if (batchState.pendingRows.length) {
    appendCatalogDataRowsBatch_(
      batchState.ctx.sheets.data,
      batchState.ctx.dataCol,
      batchState.pendingRows
    );
  }
  mark('write_data');

  if (batchState.permissionSyncJobs.length) {
    scheduleImportDriveSyncJobs_(batchState.permissionSyncJobs);
  }

  var viewMode = finalizeBatchImportView_(batchState, parentId, entries, options);
  mark('view_' + viewMode);

  var progressDone = Math.min(totalCount, doneOffset + entries.length);
  var result = {
    ok: batchState.failed.length === 0,
    data: {
      imported: batchState.imported,
      failed: batchState.failed,
      count: batchState.imported.length,
      progress: {
        done: progressDone,
        total: totalCount,
        phase: options.isLast === false ? 'copy' : 'done'
      },
      view_mode: viewMode
    },
    message: buildImportBatchResultMessage_(batchState.imported, batchState.failed)
  };
  if (profile) {
    var totalMs = 0;
    for (var t = 0; t < profile.steps.length; t++) {
      totalMs += profile.steps[t].ms;
    }
    profile.steps.push({ step: 'total', ms: totalMs });
    result.data.timings = profile.steps;
    result.message = profile.steps.map(function (item) {
      return item.step + ': ' + item.ms + ' ms';
    }).join('\n') + '\n\n' + result.message;
  }
  return result;
}

/**
 * @param {Array} imported
 * @param {Array} failed
 * @returns {string}
 */
function buildImportBatchResultMessage_(imported, failed) {
  var parts = [];
  if (imported.length) {
    parts.push('Импортировано файлов: ' + imported.length + '.');
  }
  if (failed.length) {
    parts.push('Ошибок: ' + failed.length + '.');
    for (var i = 0; i < failed.length && i < 5; i++) {
      parts.push('• ' + failed[i].name + ' — ' + failed[i].message);
    }
    if (failed.length > 5) {
      parts.push('… и ещё ' + (failed.length - 5));
    }
  }
  if (!parts.length) {
    return 'Нет файлов для импорта.';
  }
  return parts.join('\n');
}

/**
 * @returns {Array}
 */
function loadPendingImportDriveSyncQueue_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(PENDING_IMPORT_DRIVE_SYNC_KEY);
    if (!raw) {
      return [];
    }
    var parsed = JSON.parse(raw);
    return parsed && parsed.length ? parsed : [];
  } catch (err) {
    Logger.log('loadPendingImportDriveSyncQueue_: ' + err);
    return [];
  }
}

/**
 * @param {Array} queue
 */
function savePendingImportDriveSyncQueue_(queue) {
  PropertiesService.getDocumentProperties().setProperty(
    PENDING_IMPORT_DRIVE_SYNC_KEY,
    JSON.stringify(queue || [])
  );
}

/**
 * @param {Array<{file_id: string, permissions_json: string}>} jobs
 */
function scheduleImportDriveSyncJobs_(jobs) {
  if (!jobs || !jobs.length) {
    return;
  }
  var queue = loadPendingImportDriveSyncQueue_();
  for (var i = 0; i < jobs.length; i++) {
    if (jobs[i] && jobs[i].file_id) {
      queue.push(jobs[i]);
    }
  }
  savePendingImportDriveSyncQueue_(queue);
  ensureImportDriveSyncTrigger_();
}

function ensureImportDriveSyncTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === IMPORT_DRIVE_SYNC_HANDLER) {
      return;
    }
  }
  ScriptApp.newTrigger(IMPORT_DRIVE_SYNC_HANDLER)
    .timeBased()
    .after(2000)
    .create();
}

function deleteImportDriveSyncTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === IMPORT_DRIVE_SYNC_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function processPendingImportDriveSync_() {
  deleteImportDriveSyncTriggers_();
  var queue = loadPendingImportDriveSyncQueue_();
  if (!queue.length) {
    return;
  }
  savePendingImportDriveSyncQueue_([]);
  for (var i = 0; i < queue.length; i++) {
    var job = queue[i];
    if (!job || !job.file_id) {
      continue;
    }
    try {
      syncDrivePermissionsFromJson_(job.file_id, job.permissions_json);
    } catch (err) {
      Logger.log('processPendingImportDriveSync_ ' + job.file_id + ': ' + err);
    }
  }
}

/**
 * Передаёт владение импортированным файлом владельцу каталога; при неудаче — в очередь отчёта.
 * @param {string} fileId
 * @param {string} fileName
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {string} примечание к сообщению импорта (может быть пустым)
 */
function finalizeImportedDriveFileOwnership_(fileId, fileName, ss) {
  var ownershipResult = transferDriveFileOwnershipToCatalogOwner_(fileId, ss);
  if (!ownershipResult) {
    return '';
  }
  if (ownershipResult.status === 'transferred' || ownershipResult.status === 'already_owner') {
    return '';
  }
  appendImportOwnershipBatchPending_({
    file_id: fileId,
    name: fileName,
    message: ownershipResult.message || 'не удалось передать владение'
  });
  return ' (владение: ' + (ownershipResult.message || 'ожидает принятия') + ')';
}

/**
 * @param {string} parentId
 * @param {Array<string>} pathSegments
 * @param {string} fileName
 * @param {string} base64Data
 * @param {string} mimeType
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogLocalFileCore_(parentId, pathSegments, fileName, base64Data, mimeType) {
  if (!base64Data) {
    throw new Error('Файл не выбран или не удалось прочитать данные.');
  }
  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Data);
  } catch (decodeErr) {
    throw new Error('Не удалось декодировать файл: ' + decodeErr);
  }
  if (!bytes || !bytes.length) {
    throw new Error('Файл пустой.');
  }
  if (bytes.length > CATALOG_MAX_UPLOAD_BYTES) {
    throw new Error('Файл больше ' + Math.floor(CATALOG_MAX_UPLOAD_BYTES / (1024 * 1024)) + ' МБ.');
  }

  var blob = Utilities.newBlob(
    bytes,
    mimeType || 'application/octet-stream',
    sanitizeFolderLabel_(fileName)
  );
  return importCatalogBlobFileCore_(parentId, pathSegments, sanitizeFolderLabel_(fileName), blob, false);
}

/**
 * @param {string} parentId
 * @param {Object} entry
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogDriveEntryCore_(parentId, entry) {
  if (!entry || !entry.source_file_id) {
    throw new Error('Некорректная запись импорта с Drive.');
  }
  if (entry.import_mode === 'move') {
    if (!entry.can_move) {
      throw new Error('Перенос недоступен для «' + (entry.file_name || entry.source_file_id) +
        '»: вы не владелец файла на Drive.');
    }
    return importCatalogDriveMoveCore_(
      parentId,
      entry.path_segments || [],
      entry.file_name,
      entry.source_file_id,
      entry.mime_type || '',
      false
    );
  }
  return importCatalogDriveCopyCore_(
    parentId,
    entry.path_segments || [],
    entry.file_name,
    entry.source_file_id,
    entry.mime_type || '',
    false
  );
}

/**
 * @param {string} parentId
 * @param {Array<string>} pathSegments
 * @param {string} fileName
 * @param {GoogleAppsScript.Base.Blob} blob
 * @param {boolean} refreshView
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogBlobFileCore_(parentId, pathSegments, fileName, blob, refreshView) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCanPromoteImportedFiles_(ss);
  var ctx = buildCatalogImportContext_(ss);
  var targetParent = ensureImportFolderPath_(
    ctx.sheets.data,
    ctx.dataCol,
    parentId,
    pathSegments,
    ctx.usersMap
  );

  var allNodes = readCatalogDataNodes_(ctx.sheets.data, ctx.dataCol);
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    fileName,
    allNodes,
    {},
    {}
  );

  blob.setName(unique.name);
  var driveFile = ingestCatalogDriveFileViaImportStaging_({
    ss: ss,
    finalName: unique.name,
    blob: blob
  });

  var newNodeId;
  try {
    newNodeId = registerCatalogDriveFileRow_(
      ctx.sheets,
      ctx.dataCol,
      targetParent,
      driveFile.id,
      unique.name,
      driveFile.mimeType || blob.getContentType() || '',
      fileFolderCode,
      ctx.usersMap,
      true
    );
  } catch (permErr) {
    Drive.Files.remove(driveFile.id, { supportsAllDrives: true });
    throw new Error('Файл загружен, но не удалось применить права: ' + permErr);
  }

  if (refreshView) {
    insertCatalogViewNodeAfterCreate_(ctx.sheets.view, ctx.sheets.data, newNodeId, targetParent.id);
  }

  var ownershipNote = finalizeImportedDriveFileOwnership_(driveFile.id, unique.name, ss);

  return {
    ok: true,
    data: { file_id: driveFile.id, name: unique.name, parent_id: targetParent.id },
    message: 'Импортирован: ' + unique.name + ownershipNote
  };
}

/**
 * @param {string} parentId
 * @param {Array<string>} pathSegments
 * @param {string} fileName
 * @param {string} sourceFileId
 * @param {string} mimeType
 * @param {boolean} refreshView
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogDriveCopyCore_(parentId, pathSegments, fileName, sourceFileId, mimeType, refreshView) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ctx = buildCatalogImportContext_(ss);
  var targetParent = ensureImportFolderPath_(
    ctx.sheets.data,
    ctx.dataCol,
    parentId,
    pathSegments,
    ctx.usersMap
  );

  var allNodes = readCatalogDataNodes_(ctx.sheets.data, ctx.dataCol);
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    sanitizeFolderLabel_(fileName),
    allNodes,
    {},
    {}
  );

  var inherited = inheritPermissionsFromParent_(targetParent, ctx.usersMap);
  var copyResult = invokeOwnerImportWebApp_(ss.getId(), 'copy', {
    source_file_id: sourceFileId,
    target_name: unique.name,
    permissions_json: inherited.permissions_json
  });
  var driveFile = {
    id: copyResult.data.file_id,
    name: copyResult.data.name,
    mimeType: copyResult.data.mime_type || mimeType || ''
  };

  var newNodeId;
  try {
    newNodeId = registerCatalogDriveFileRow_(
      ctx.sheets,
      ctx.dataCol,
      targetParent,
      driveFile.id,
      unique.name,
      driveFile.mimeType || mimeType || '',
      fileFolderCode,
      ctx.usersMap,
      false
    );
  } catch (permErr) {
    Drive.Files.remove(driveFile.id, { supportsAllDrives: true });
    throw new Error('Не удалось записать файл в каталог: ' + permErr);
  }

  if (refreshView) {
    insertCatalogViewNodeAfterCreate_(ctx.sheets.view, ctx.sheets.data, newNodeId, targetParent.id);
  }

  return {
    ok: true,
    data: { file_id: driveFile.id, name: unique.name, parent_id: targetParent.id },
    message: 'Скопирован: ' + unique.name
  };
}

/**
 * Переносит файл на Drive в root_folder_id (тот же file_id, без дублирования места).
 * @param {string} parentId
 * @param {Array<string>} pathSegments
 * @param {string} fileName
 * @param {string} sourceFileId
 * @param {string} mimeType
 * @param {boolean} refreshView
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function importCatalogDriveMoveCore_(parentId, pathSegments, fileName, sourceFileId, mimeType, refreshView) {
  if (!isImporterOwnerOfDriveFileId_(sourceFileId)) {
    throw new Error('Перенос возможен только для файлов, где вы владелец на Drive.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ctx = buildCatalogImportContext_(ss);
  var targetParent = ensureImportFolderPath_(
    ctx.sheets.data,
    ctx.dataCol,
    parentId,
    pathSegments,
    ctx.usersMap
  );

  var allNodes = readCatalogDataNodes_(ctx.sheets.data, ctx.dataCol);
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    sanitizeFolderLabel_(fileName),
    allNodes,
    {},
    {}
  );

  var fileMeta = Drive.Files.get(sourceFileId, {
    supportsAllDrives: true,
    fields: 'id,name,parents,mimeType'
  });
  if (!fileMeta || !fileMeta.id) {
    throw new Error('Файл не найден на Drive.');
  }

  var removeParents = fileMeta.parents || [];
  var driveFile;
  if (isSpreadsheetOwner_(ss)) {
    driveFile = moveDriveFileToCatalogRoot_(
      sourceFileId,
      ctx.rootFolderId,
      unique.name,
      removeParents
    );
  } else {
    driveFile = ingestCatalogDriveFileViaImportStaging_({
      ss: ss,
      finalName: unique.name,
      existingFileId: sourceFileId,
      removeParentIds: removeParents,
      mimeType: fileMeta.mimeType || mimeType || ''
    });
  }

  var newNodeId;
  try {
    newNodeId = registerCatalogDriveFileRow_(
      ctx.sheets,
      ctx.dataCol,
      targetParent,
      sourceFileId,
      unique.name,
      driveFile.mimeType || mimeType || '',
      fileFolderCode,
      ctx.usersMap,
      true
    );
  } catch (permErr) {
    throw new Error('Файл перенесён, но не удалось применить права: ' + permErr);
  }

  if (refreshView) {
    insertCatalogViewNodeAfterCreate_(ctx.sheets.view, ctx.sheets.data, newNodeId, targetParent.id);
  }

  return {
    ok: true,
    data: { file_id: sourceFileId, name: unique.name, parent_id: targetParent.id },
    message: 'Перенесён: ' + unique.name
  };
}

/**
 * Передаёт владение файлом владельцу таблицы-каталога.
 * @param {string} fileId
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {{ok: boolean, status?: string, skipped?: boolean, message?: string}}
 */
function transferDriveFileOwnershipToCatalogOwner_(fileId, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var targetEmail = getSpreadsheetOwnerEmail_(ss);
  if (!targetEmail) {
    return {
      ok: false,
      status: 'failed',
      message: 'Не удалось определить email владельца каталога.'
    };
  }

  var currentOwnerEmail = getDriveFileOwnerEmail_(fileId);
  if (currentOwnerEmail && currentOwnerEmail === targetEmail) {
    return { ok: true, status: 'already_owner', skipped: true };
  }

  var importerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (currentOwnerEmail && importerEmail && currentOwnerEmail !== importerEmail) {
    return {
      ok: false,
      status: 'not_importer_owner',
      message: 'текущий владелец ' + currentOwnerEmail + ', передача возможна только от импортёра'
    };
  }

  try {
    Drive.Permissions.create({
      type: 'user',
      role: 'owner',
      emailAddress: targetEmail
    }, fileId, {
      transferOwnership: true,
      supportsAllDrives: true,
      sendNotificationEmail: true,
      emailMessage: 'Файл импортирован в каталог Google Drive. Примите владение, чтобы завершить импорт.'
    });
  } catch (err) {
    return {
      ok: false,
      status: 'failed',
      message: formatDriveOwnershipTransferError_(String(err), targetEmail)
    };
  }

  var newOwnerEmail = getDriveFileOwnerEmail_(fileId);
  if (newOwnerEmail === targetEmail) {
    return { ok: true, status: 'transferred' };
  }

  return {
    ok: false,
    status: 'pending',
    message: 'отправлено приглашение ' + targetEmail + ' — нужно принять в почте Google Drive'
  };
}

/**
 * @param {string} details
 * @param {string} targetEmail
 * @returns {string}
 */
function formatDriveOwnershipTransferError_(details, targetEmail) {
  if (/ownershipChangeAcrossDomainNotPermitted|changeOwner|domain/i.test(details)) {
    return 'политика Google Workspace не позволяет передать владение на ' + targetEmail;
  }
  if (/recipientLimit|storage|quota/i.test(details)) {
    return 'у ' + targetEmail + ' недостаточно места на Drive';
  }
  return details;
}

/**
 * @param {string} fileId
 * @returns {string}
 */
function getDriveFileOwnerEmail_(fileId) {
  try {
    var file = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'owners(emailAddress)'
    });
    if (file.owners && file.owners.length) {
      return normalizeEmail_(file.owners[0].emailAddress || '');
    }
  } catch (err) {
    Logger.log('getDriveFileOwnerEmail_: ' + err);
  }
  return '';
}

/**
 * Минимальный контекст импорта владельца: без _Import, без validateRootFolder.
 *
 * @param {GoogleAppsScript.Spreadsheet=} ss
 * @returns {{sheets: object, dataCol: Object, rootFolderId: string, usersMap: Object}}
 */
function buildCatalogImportContextLight_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Сначала выполните setupSchema.');
  }
  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    throw new Error('Укажите root_folder_id на листе Settings.');
  }
  return {
    sheets: getCatalogSheets_(ss),
    dataCol: columnIndexMap_(SHEET_HEADERS.CatalogData),
    rootFolderId: rootFolderId,
    usersMap: readUsersMapFromSpreadsheet_(ss)
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {{sheets: object, dataCol: Object, rootFolderId: string, importFolderId: string, usersMap: Object}}
 */
function buildCatalogImportContext_(ss) {
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Сначала выполните setupSchema.');
  }
  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    throw new Error('Укажите root_folder_id на листе Settings.');
  }
  var folderCheck = validateRootFolder_(rootFolderId);
  if (!folderCheck.ok) {
    throw new Error(folderCheck.message || folderCheck.error);
  }
  var importFolderId = ensureCatalogImportDriveFolder_(ss);
  return {
    sheets: getCatalogSheets_(ss),
    dataCol: columnIndexMap_(SHEET_HEADERS.CatalogData),
    rootFolderId: rootFolderId,
    importFolderId: importFolderId,
    usersMap: readUsersMapFromSpreadsheet_(ss)
  };
}

/**
 * Создаёт или находит скрытую папку _Import на Drive и синхронизирует editor для active Users.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {string}
 */
function ensureCatalogImportDriveFolder_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var usersSheet = ss.getSheetByName('Users');
  if (!settingsSheet || !usersSheet) {
    throw new Error('Сначала выполните setupSchema.');
  }

  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    throw new Error('Укажите root_folder_id на листе Settings.');
  }
  var rootCheck = validateRootFolder_(rootFolderId);
  if (!rootCheck.ok) {
    throw new Error(rootCheck.message || rootCheck.error);
  }

  var importFolderId = getSetting_(settingsSheet, SETTING_IMPORT_FOLDER_ID);
  if (importFolderId) {
    var importCheck = validateRootFolder_(importFolderId);
    if (importCheck.ok) {
      syncImportFolderEditors_(importFolderId, usersSheet, ss);
      return importFolderId;
    }
  }

  importFolderId = findChildDriveFolderByName_(rootFolderId, IMPORT_DRIVE_FOLDER_NAME);
  if (!importFolderId) {
    var folder = Drive.Files.create({
      name: IMPORT_DRIVE_FOLDER_NAME,
      mimeType: DRIVE_FOLDER_MIME,
      parents: [rootFolderId]
    }, null, {
      supportsAllDrives: true,
      fields: 'id,name'
    });
    if (!folder || !folder.id) {
      throw new Error('Не удалось создать папку Import на Drive.');
    }
    importFolderId = folder.id;
    transferDriveFileOwnershipToCatalogOwner_(importFolderId, ss);
  }

  upsertSetting_(settingsSheet, SETTING_IMPORT_FOLDER_ID, importFolderId);
  syncImportFolderEditors_(importFolderId, usersSheet, ss);
  return importFolderId;
}

/**
 * @param {string} parentFolderId
 * @param {string} folderName
 * @returns {string}
 */
function findChildDriveFolderByName_(parentFolderId, folderName) {
  if (!parentFolderId || !folderName) {
    return '';
  }
  var escapedName = String(folderName).replace(/'/g, "\\'");
  var query = "'" + parentFolderId + "' in parents and mimeType = '" + DRIVE_FOLDER_MIME +
    "' and name = '" + escapedName + "' and trashed = false";
  var response = Drive.Files.list({
    q: query,
    pageSize: 1,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id,name)'
  });
  var files = response.files || [];
  return files.length ? files[0].id : '';
}

/**
 * Выдаёт editor на _Import всем active Users; снимает прямой доступ у неактивных.
 *
 * @param {string} importFolderId
 * @param {GoogleAppsScript.Spreadsheet.Sheet} usersSheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 */
function syncImportFolderEditors_(importFolderId, usersSheet, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var catalogOwner = normalizeEmail_(getSpreadsheetOwnerEmail_(ss));
  var activeUsers = readActiveUsersList_(usersSheet);
  var desiredWriters = {};
  for (var i = 0; i < activeUsers.length; i++) {
    var email = activeUsers[i].email;
    if (email && email !== catalogOwner) {
      desiredWriters[email] = true;
    }
  }

  var current = listDrivePermissionsWithIds_(importFolderId);
  var bestDirectByEmail = {};
  for (var c = 0; c < current.length; c++) {
    var perm = current[c];
    if (perm.inherited || perm.type !== 'user' || !perm.subject) {
      continue;
    }
    var key = normalizeEmail_(perm.subject);
    if (!bestDirectByEmail[key] || compareDriveRole_(perm.role, bestDirectByEmail[key].role) > 0) {
      bestDirectByEmail[key] = perm;
    }
  }

  for (var writerEmail in desiredWriters) {
    if (!desiredWriters.hasOwnProperty(writerEmail)) {
      continue;
    }
    var existing = bestDirectByEmail[writerEmail];
    if (existing && compareDriveRole_(existing.role, 'writer') >= 0) {
      continue;
    }
    try {
      Drive.Permissions.create({
        type: 'user',
        role: 'writer',
        emailAddress: writerEmail
      }, importFolderId, {
        supportsAllDrives: true,
        sendNotificationEmail: false
      });
    } catch (err) {
      if (!isDrivePermissionAlreadyExistsError_(String(err))) {
        Logger.log('syncImportFolderEditors_ add ' + writerEmail + ': ' + err);
      }
    }
  }

  for (var r = 0; r < current.length; r++) {
    var cur = current[r];
    if (cur.inherited || cur.type !== 'user' || !cur.subject) {
      continue;
    }
    var subject = normalizeEmail_(cur.subject);
    if (subject === catalogOwner || desiredWriters[subject]) {
      continue;
    }
    try {
      Drive.Permissions.remove(importFolderId, cur.id, { supportsAllDrives: true });
    } catch (removeErr) {
      Logger.log('syncImportFolderEditors_ remove ' + cur.id + ': ' + removeErr);
    }
  }
}

/**
 * @param {string} finalName
 * @returns {string}
 */
function buildImportStagingFileName_(finalName) {
  return 'staging_' + Utilities.getUuid().substring(0, 8) + '_' + String(finalName || 'file');
}

/**
 * Переносит файл из _Import в root_folder_id с финальным именем.
 *
 * @param {string} fileId
 * @param {string} importFolderId
 * @param {string} rootFolderId
 * @param {string} finalName
 * @returns {Object}
 */
function promoteDriveFileFromImportToCatalogRoot_(fileId, importFolderId, rootFolderId, finalName) {
  var updated = Drive.Files.update({
    name: finalName
  }, fileId, null, {
    addParents: rootFolderId,
    removeParents: importFolderId,
    supportsAllDrives: true,
    fields: 'id,name,mimeType,parents'
  });
  if (!updated || !updated.id) {
    throw new Error('Не удалось перенести файл из Import в каталог на Drive.');
  }
  if (driveFileParentsInclude_(updated.parents, rootFolderId)) {
    return updated;
  }
  if (driveFileParentsInclude_(updated.parents, importFolderId)) {
    throw new Error('Файл остался в папке Import на Drive. Проверьте import_folder_id и права.');
  }
  return updated;
}

/**
 * @param {string} fileId
 * @param {string} rootFolderId
 * @param {string} finalName
 * @param {Array<string>} removeParentIds
 * @returns {Object}
 */
function moveDriveFileToCatalogRoot_(fileId, rootFolderId, finalName, removeParentIds) {
  var removeParents = (removeParentIds || []).filter(function (parentId) {
    return parentId && parentId !== rootFolderId;
  });
  var optionalArgs = {
    addParents: rootFolderId,
    supportsAllDrives: true,
    fields: 'id,name,mimeType,parents'
  };
  if (removeParents.length) {
    optionalArgs.removeParents = removeParents.join(',');
  }
  var updated = Drive.Files.update({
    name: finalName
  }, fileId, null, optionalArgs);
  if (!updated || !updated.id) {
    throw new Error('Не удалось переместить файл в каталог на Drive.');
  }
  if (driveFileParentsInclude_(updated.parents, rootFolderId)) {
    return updated;
  }
  if (removeParents.length && driveFileParentsInclude_(updated.parents, removeParents[0])) {
    throw new Error('Файл не удалось отвязать от прежней папки на Drive.');
  }
  return updated;
}

/**
 * @param {Array<string>|undefined} parents
 * @param {string} parentId
 * @returns {boolean}
 */
function driveFileParentsInclude_(parents, parentId) {
  if (!parentId) {
    return false;
  }
  parents = parents || [];
  for (var i = 0; i < parents.length; i++) {
    if (parents[i] === parentId) {
      return true;
    }
  }
  return false;
}

/**
 * @param {string} fileId
 * @param {string} parentId
 */
function assertDriveFileHasParent_(fileId, parentId) {
  var meta = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'parents'
  });
  var parents = meta.parents || [];
  for (var i = 0; i < parents.length; i++) {
    if (parents[i] === parentId) {
      return;
    }
  }
  throw new Error('Файл не переместился в каталог на Drive (нет доступа к root_folder_id).');
}

/**
 * Перенос из _Import в root от имени владельца каталога (напрямую или через веб-приложение).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} fileId
 * @param {string} finalName
 * @returns {{mimeType: string}}
 */
function promoteImportedFileViaOwnerWebApp_(ss, fileId, finalName) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (isSpreadsheetOwner_(ss) || checkImportWebAppConfigured_().ok) {
    var result = invokeOwnerImportWebApp_(ss.getId(), 'promote', {
      file_id: fileId,
      target_name: finalName
    });
    return { mimeType: (result.data && result.data.mime_type) || '' };
  }

  throw new Error(
    'Для переноса файла из Import в каталог нужно веб-приложение импорта (развёрнуто от имени владельца). ' +
    formatImportWebAppNotConfiguredMessage_()
  );
}

/**
 * Загрузка/создание/перенос через staging-папку _Import → root + владелец каталога.
 *
 * @param {Object} params
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} params.ss
 * @param {string} params.finalName
 * @param {GoogleAppsScript.Base.Blob=} params.blob
 * @param {string=} params.mimeType
 * @param {string=} params.existingFileId
 * @param {Array<string>=} params.removeParentIds
 * @returns {{id: string, name: string, mimeType: string}}
 */
function ingestCatalogDriveFileViaImportStaging_(params) {
  var ss = params.ss || SpreadsheetApp.getActiveSpreadsheet();
  var ctx = buildCatalogImportContext_(ss);
  var fileId = params.existingFileId || '';
  var mimeType = params.mimeType || '';

  if (!fileId) {
    var stagingName = buildImportStagingFileName_(params.finalName);
    if (params.blob) {
      params.blob.setName(stagingName);
      var uploaded = Drive.Files.create({
        name: stagingName,
        parents: [ctx.importFolderId]
      }, params.blob, {
        supportsAllDrives: true,
        fields: 'id,name,mimeType'
      });
      if (!uploaded || !uploaded.id) {
        throw new Error('Drive не вернул id загруженного файла.');
      }
      fileId = uploaded.id;
      mimeType = uploaded.mimeType || params.blob.getContentType() || mimeType;
    } else {
      if (!params.mimeType) {
        throw new Error('Не указан mimeType для создания файла.');
      }
      var created = Drive.Files.create({
        name: stagingName,
        mimeType: params.mimeType,
        parents: [ctx.importFolderId]
      }, null, {
        supportsAllDrives: true,
        fields: 'id,name,mimeType'
      });
      if (!created || !created.id) {
        throw new Error('Drive не вернул id созданного файла.');
      }
      fileId = created.id;
      mimeType = created.mimeType || params.mimeType;
    }
  } else if (params.removeParentIds && params.removeParentIds.length) {
    var moveStagingName = buildImportStagingFileName_(params.finalName);
    var movedToImport = Drive.Files.update({
      name: moveStagingName
    }, fileId, null, {
      addParents: ctx.importFolderId,
      removeParents: params.removeParentIds.join(','),
      supportsAllDrives: true,
      fields: 'id,name,mimeType,parents'
    });
    if (!movedToImport || !movedToImport.id) {
      throw new Error('Не удалось переместить файл в Import на Drive.');
    }
    mimeType = movedToImport.mimeType || mimeType;
  }

  finalizeImportedDriveFileOwnership_(fileId, params.finalName, ss);
  var promoted = promoteImportedFileViaOwnerWebApp_(ss, fileId, params.finalName);

  return {
    id: fileId,
    name: params.finalName,
    mimeType: promoted.mimeType || mimeType || ''
  };
}

/**
 * Локальный импорт возможен только если владелец каталога сам загружает
 * или настроено веб-приложение (перенос Import → root от его имени).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function ensureCanPromoteImportedFiles_(ss) {
  var ownerEmail = normalizeEmail_(getSpreadsheetOwnerEmail_(ss));
  var runnerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (ownerEmail && runnerEmail && ownerEmail === runnerEmail) {
    return;
  }
  var webAppCheck = checkImportWebAppConfigured_();
  if (!webAppCheck.ok) {
    throw new Error(
      'Для локального импорта нужно веб-приложение (развёрнуто от имени владельца каталога). ' +
      (webAppCheck.message || webAppCheck.error || '')
    );
  }
}

/**
 * @param {Object} parent
 */
function validateCatalogImportParent_(parent) {
  if (!parent) {
    throw new Error('Родительская папка не найдена.');
  }
  if (parent.folder_code === FOLDER_CODE_TRASH) {
    throw new Error('Нельзя импортировать в «Корзину».');
  }
  if (parent.type !== 'folder') {
    throw new Error('Родитель должен быть папкой.');
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {string} anchorParentId
 * @param {Array<string>} pathSegments
 * @param {Object} usersMap
 * @returns {Object}
 */
function ensureImportFolderPath_(dataSheet, dataCol, anchorParentId, pathSegments, usersMap) {
  var parent = findCatalogNodeById_(dataSheet, dataCol, anchorParentId);
  validateCatalogImportParent_(parent);
  var currentParent = parent;

  for (var i = 0; i < pathSegments.length; i++) {
    var segment = sanitizeFolderLabel_(pathSegments[i]);
    if (!segment) {
      continue;
    }
    var allNodes = readCatalogDataNodes_(dataSheet, dataCol);
    var existing = findChildFolderByLabel_(currentParent.id, segment, allNodes);
    if (existing) {
      currentParent = existing;
      continue;
    }
    currentParent = createImportVirtualFolderCore_(dataSheet, dataCol, currentParent, segment, usersMap);
  }

  return currentParent;
}

/**
 * @param {string} parentId
 * @param {string} label
 * @param {Array} allNodes
 * @returns {Object|null}
 */
function findChildFolderByLabel_(parentId, label, allNodes) {
  var wanted = sanitizeFolderLabel_(label);
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.parent_id !== parentId || node.type !== 'folder') {
      continue;
    }
    var nodeLabel = stripFolderPrefix_(node.name, node.folder_code);
    if (nodeLabel === wanted) {
      return node;
    }
  }
  return null;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {Object} parent
 * @param {string} label
 * @param {Object} usersMap
 * @returns {Object}
 */
function createImportVirtualFolderCore_(dataSheet, dataCol, parent, label, usersMap) {
  assertCanAddSubfolder_(readCell_(parent.folder_code));
  var folderCode = allocateFolderCode_(dataSheet, dataCol, parent);
  var fullName = buildFolderName_(folderCode, label);
  var inherited = inheritPermissionsFromParent_(parent, usersMap);
  var newNodeId = generateId_();

  appendCatalogDataRow_(dataSheet, dataCol, {
    id: newNodeId,
    name: fullName,
    type: 'folder',
    folder_code: folderCode,
    parent_id: parent.id,
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  });

  return {
    id: newNodeId,
    name: fullName,
    type: 'folder',
    folder_code: folderCode,
    parent_id: parent.id
  };
}

/**
 * @param {Object} sheets
 * @param {Object.<string, number>} dataCol
 * @param {Object} parent
 * @param {string} fileId
 * @param {string} fullName
 * @param {string} mimeType
 * @param {string} fileFolderCode
 * @param {Object} usersMap
 * @param {boolean} syncPermissions
 * @returns {string} new node id
 */
function registerCatalogDriveFileRow_(
  sheets,
  dataCol,
  parent,
  fileId,
  fullName,
  mimeType,
  fileFolderCode,
  usersMap,
  syncPermissions
) {
  var inherited = inheritPermissionsFromParent_(parent, usersMap);
  var permissions = parseCatalogPermissionsJson_(inherited.permissions_json);
  if (syncPermissions) {
    syncDrivePermissionsFromJson_(fileId, permissions, sheets.data.getParent());
  }

  var newNodeId = generateId_();
  appendCatalogDataRow_(sheets.data, dataCol, {
    id: newNodeId,
    file_id: fileId,
    name: fullName,
    type: 'file',
    folder_code: fileFolderCode,
    parent_id: parent.id,
    mime_type: mimeType || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  });
  return newNodeId;
}

/**
 * @param {string} parentId
 * @param {string} label
 * @param {string} base64Data
 * @param {string} mimeType
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function uploadCatalogFile(parentId, label, base64Data, mimeType) {
  try {
    return uploadCatalogFileCore_(parentId, label, base64Data, mimeType);
  } catch (err) {
    return {
      ok: false,
      error: 'UPLOAD_FILE_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} parentId
 * @param {string} label
 * @param {string} base64Data
 * @param {string} mimeType
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function uploadCatalogFileCore_(parentId, label, base64Data, mimeType) {
  var result = importCatalogLocalFileCore_(parentId, [], label, base64Data, mimeType);
  finishCatalogImport_(parentId);
  if (result.data) {
    result.message = 'Файл загружен: ' + result.data.name;
  }
  return result;
}

/**
 * Открывает HTML-диалог создания виртуальной папки.
 */
function createFolderDialog_() {
  var contextResult = getCreateFolderContext();
  var template = HtmlService.createTemplateFromFile('CreateFolderDialog');
  template.initialContext = JSON.stringify(contextResult);
  var html = template.evaluate()
    .setWidth(460)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Создать папку');
}

/**
 * Контекст для диалога: родительская папка по выделенной строке Catalog.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getCreateFolderContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = getCatalogSheets_(ss);

    var headerCheck = validateCatalogDataHeaders_(sheets.data);
    if (!headerCheck.ok) {
      return headerCheck;
    }

    var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
    var viewCol = getCatalogViewColumnMap_();
    var parent = resolveCreateFolderParent_(ss, sheets.view, sheets.data, viewCol, dataCol);
    if (!parent.ok) {
      return parent;
    }

    return {
      ok: true,
      data: {
        parent_id: parent.data.id,
        parent_label: parent.data.label,
        parent_name: parent.data.name,
        location_message: parent.data.location_message,
        selection_mode: parent.data.selection_mode
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'CREATE_CONTEXT_FAILED',
      message: String(err)
    };
  }
}

/**
 * Создаёт виртуальную папку: только строка в CatalogData (без объекта на Drive).
 *
 * @param {string} parentId
 * @param {string} label
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function createVirtualFolder(parentId, label) {
  try {
    return createVirtualFolderCore_(parentId, label);
  } catch (err) {
    return {
      ok: false,
      error: 'CREATE_FOLDER_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} parentId
 * @param {string} label
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function createVirtualFolderCore_(parentId, label) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);
  if (!ss.getSheetByName('Settings')) {
    throw new Error('Сначала выполните setupSchema.');
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var parent = findCatalogNodeById_(sheets.data, dataCol, parentId);
  if (!parent) {
    throw new Error('Родительская папка не найдена.');
  }
  if (parent.folder_code === FOLDER_CODE_TRASH) {
    throw new Error('Нельзя создавать папки в «Корзине».');
  }
  if (parent.type !== 'folder') {
    throw new Error('Родитель должен быть папкой.');
  }

  var cleanLabel = sanitizeFolderLabel_(label);
  if (!cleanLabel) {
    throw new Error('Введите имя папки.');
  }

  assertCanAddSubfolder_(readCell_(parent.folder_code));

  var folderCode = allocateFolderCode_(sheets.data, dataCol, parent);
  var fullName = buildFolderName_(folderCode, cleanLabel);
  var inherited = inheritPermissionsFromParent_(parent, readUsersMapFromSpreadsheet_(sheets.data.getParent()));
  var newNodeId = generateId_();

  appendCatalogDataRow_(sheets.data, dataCol, {
    id: newNodeId,
    name: fullName,
    type: 'folder',
    folder_code: folderCode,
    parent_id: parent.id,
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  });

  insertCatalogViewNodeAfterCreate_(sheets.view, sheets.data, newNodeId, parent.id);

  return {
    ok: true,
    data: {
      folder_code: folderCode,
      name: fullName,
      parent_id: parent.id
    },
    message: 'Папка создана: ' + fullName
  };
}

/**
 * Открывает HTML-диалог переименования элемента Catalog.
 */
function renameDialog_() {
  var html = HtmlService.createHtmlOutputFromFile('RenameDialog')
    .setWidth(460)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Переименовать');
}

/**
 * Контекст для диалога: одна выделенная строка Catalog.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getRenameContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = getCatalogSheets_(ss);

    var headerCheck = validateCatalogDataHeaders_(sheets.data);
    if (!headerCheck.ok) {
      return headerCheck;
    }

    var viewCol = getCatalogViewColumnMap_();
    var selected = resolveSingleSelectedCatalogEntry_(ss, sheets.view, viewCol);
    if (!selected.ok) {
      return selected;
    }

    var entry = selected.data;
    if (isProtectedCatalogNode_(entry.node)) {
      return {
        ok: false,
        error: 'PROTECTED_NODE',
        message: 'Нельзя переименовать «Каталог» и «Корзину».'
      };
    }

    var label = stripFolderPrefix_(entry.node.name, entry.node.folder_code);
    return {
      ok: true,
      data: {
        item_id: entry.node.id,
        current_label: label,
        current_name: entry.node.name,
        type: entry.node.type,
        rename_message: formatRenameMessage_(entry.node, label)
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'RENAME_CONTEXT_FAILED',
      message: String(err)
    };
  }
}

/**
 * Переименовывает папку или файл в Catalog (на Drive — только файлы).
 *
 * @param {string} itemId
 * @param {string} newLabel
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function renameCatalogItem(itemId, newLabel) {
  try {
    return renameCatalogItemCore_(itemId, newLabel);
  } catch (err) {
    return {
      ok: false,
      error: 'RENAME_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} itemId
 * @param {string} newLabel
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function renameCatalogItemCore_(itemId, newLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var entry = findCatalogDataEntryById_(sheets.data, dataCol, itemId);
  if (!entry) {
    throw new Error('Элемент не найден в CatalogData.');
  }
  if (isProtectedCatalogNode_(entry)) {
    throw new Error('Нельзя переименовать «Каталог» и «Корзину».');
  }

  var cleanLabel = sanitizeFolderLabel_(newLabel);
  if (!cleanLabel) {
    throw new Error('Введите новое имя.');
  }

  var oldName = entry.name;
  var newName = buildCatalogFullName_(entry.folder_code, cleanLabel);
  if (newName === oldName) {
    throw new Error('Имя не изменилось.');
  }

  if (entry.file_id && entry.type !== 'folder') {
    Drive.Files.update({ name: newName }, entry.file_id, null, {
      supportsAllDrives: true,
      fields: 'id,name'
    });
  }

  updateCatalogDataRowName_(sheets.data, dataCol, entry.rowIndex, newName);
  renderCatalogViewCore_(sheets.view, sheets.data);

  return {
    ok: true,
    data: {
      id: entry.id,
      old_name: oldName,
      new_name: newName
    },
    message: 'Переименовано: ' + newName
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {Object.<string, number>} viewCol
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function resolveSingleSelectedCatalogEntry_(ss, viewSheet, viewCol) {
  if (!ss.getActiveSheet() || ss.getActiveSheet().getName() !== 'Catalog') {
    return {
      ok: false,
      error: 'WRONG_SHEET',
      message: 'Выберите одну строку на листе Catalog.'
    };
  }

  var rowIndexes = getSelectedCatalogRowIndexes_(ss);
  if (!rowIndexes.length) {
    return {
      ok: false,
      error: 'NO_SELECTION',
      message: 'Выберите одну строку с папкой или файлом на листе Catalog.'
    };
  }
  if (rowIndexes.length > 1) {
    return {
      ok: false,
      error: 'MULTIPLE_SELECTION',
      message: 'Выберите только одну строку.'
    };
  }

  var entry = resolveViewRowToDataEntry_(ss, rowIndexes[0]);
  if (!entry || !entry.id || !entry.type) {
    return {
      ok: false,
      error: 'INVALID_ROW',
      message: 'Выбранная строка пуста. Выберите папку или файл.'
    };
  }

  return {
    ok: true,
    data: {
      rowIndex: entry.rowIndex,
      file_id: entry.file_id,
      node: entry
    }
  };
}

/**
 * @param {Object} node
 * @param {string} label
 * @returns {string}
 */
function formatRenameMessage_(node, label) {
  var typeLabel = node.type === 'folder' ? 'Папка' : 'Файл';
  var hint = node.folder_code
    ? 'Префикс ' + node.folder_code + '_ сохранится.'
    : 'Файл без кода папки — префикс не добавляется.';
  return typeLabel + ': «' + label + '»\n' + hint;
}

/**
 * @param {string} folderCode
 * @param {string} label
 * @returns {string}
 */
function buildCatalogFullName_(folderCode, label) {
  if (readCell_(folderCode)) {
    return buildFolderName_(folderCode, label);
  }
  return sanitizeFolderLabel_(label);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {number} rowIndex
 * @param {string} newName
 */
function updateCatalogDataRowName_(sheet, col, rowIndex, newName) {
  sheet.getRange(rowIndex, col.name).setValue(newName);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} viewRowIndex
 * @returns {Object|null}
 */
function resolveViewRowToDataEntry_(ss, viewRowIndex) {
  var sheets = getCatalogSheets_(ss);
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var indexSheet = sheets.index || getCatalogIndexSheet_(ss);
  var indexCol = getCatalogIndexColumnMap_();
  var indexEntry = readCatalogIndexEntryAtViewRow_(indexSheet, indexCol, viewRowIndex);
  if (!indexEntry || !indexEntry.data_row_index) {
    return null;
  }
  return readCatalogDataNodeAtRow_(sheets.data, dataCol, indexEntry.data_row_index);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @param {string} itemId
 * @returns {Object|null}
 */
function findCatalogDataEntryById_(dataSheet, dataCol, itemId) {
  if (!itemId) {
    return null;
  }
  var nodes = readCatalogDataNodes_(dataSheet, dataCol);
  for (var i = 0; i < nodes.length; i++) {
    if (nodes[i].id === itemId) {
      return nodes[i];
    }
  }
  return null;
}

/**
 * Родитель для новой папки по активной ячейке Catalog (view → CatalogData).
 */
function resolveCreateFolderParent_(ss, viewSheet, dataSheet, viewCol, dataCol) {
  var rootParent = getCatalogRootForCreate_(dataSheet, dataCol);
  if (!rootParent.ok) {
    return rootParent;
  }

  var onCatalog = ss.getActiveSheet() && ss.getActiveSheet().getName() === 'Catalog';
  if (!onCatalog) {
    return buildCreateFolderParentResult_(rootParent.data, 'root_default');
  }

  var cell = ss.getActiveCell();
  if (!cell || cell.getRow() <= 1) {
    return buildCreateFolderParentResult_(rootParent.data, 'root_default');
  }

  var node = resolveViewRowToDataEntry_(ss, cell.getRow());
  if (!node || !node.id || !node.type) {
    return buildCreateFolderParentResult_(rootParent.data, 'root_default');
  }

  if (node.folder_code === FOLDER_CODE_TRASH) {
    return {
      ok: false,
      error: 'PARENT_IS_TRASH',
      message: 'Нельзя создавать папки в «Корзине». Выберите другую папку или файл.'
    };
  }

  var parentNode = resolveParentFolderForCreate_(dataSheet, dataCol, node);
  if (!parentNode) {
    return {
      ok: false,
      error: 'NO_PARENT',
      message: 'Не найдена родительская папка для выбранного элемента.'
    };
  }

  if (parentNode.folder_code === FOLDER_CODE_TRASH) {
    return {
      ok: false,
      error: 'PARENT_IS_TRASH',
      message: 'Нельзя создавать папки в «Корзине». Выберите другую папку или файл.'
    };
  }

  var mode = node.type === 'folder' ? 'inside_folder' : 'file_parent';
  return buildCreateFolderParentResult_(parentNode, mode);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} col
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getCatalogRootForCreate_(dataSheet, col) {
  var rootNode = getCatalogRootNode_(dataSheet, col);
  if (!rootNode) {
    return {
      ok: false,
      error: 'NO_CATALOG_ROOT',
      message: 'Не найдена системная папка «Каталог». Выполните setupSchema.'
    };
  }
  return {
    ok: true,
    data: {
      id: rootNode.id,
      name: rootNode.name,
      label: stripFolderPrefix_(rootNode.name, rootNode.folder_code) || 'Каталог',
      folder_code: rootNode.folder_code
    }
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} col
 * @returns {Object|null}
 */
function getCatalogRootNode_(dataSheet, col) {
  var catalogId = findIdByFolderCode_(dataSheet, col, FOLDER_CODE_CATALOG);
  if (!catalogId) {
    return null;
  }
  return findCatalogNodeById_(dataSheet, col, catalogId);
}

/**
 * @param {{id: string, name: string, label: string}} parentNode
 * @param {string} mode root_default | inside_folder | file_parent
 * @returns {{ok: boolean, data: object}}
 */
function buildCreateFolderParentResult_(parentNode, mode) {
  return {
    ok: true,
    data: {
      id: parentNode.id,
      name: parentNode.name,
      label: parentNode.label || stripFolderPrefix_(parentNode.name, parentNode.folder_code),
      selection_mode: mode,
      location_message: formatCreateFolderLocationMessage_(parentNode, mode)
    }
  };
}

/**
 * @param {{name: string, label: string, folder_code: string}} parentNode
 * @param {string} mode
 * @returns {string}
 */
function formatCreateFolderLocationMessage_(parentNode, mode) {
  var label = parentNode.label || stripFolderPrefix_(parentNode.name, parentNode.folder_code) || 'Каталог';
  if (mode === 'root_default') {
    return 'Новая папка будет создана в корне каталога («' + label + '»).';
  }
  if (mode === 'inside_folder') {
    return 'Новая папка будет создана внутри папки «' + label + '».';
  }
  if (mode === 'file_parent') {
    return 'Новая папка будет создана в папке «' + label + '» (родитель выбранного файла).';
  }
  return 'Родительская папка: «' + label + '».';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @returns {Object|null}
 */
function resolveParentFolderForCreate_(dataSheet, col, node) {
  if (node.type === 'folder') {
    return node;
  }
  if (!node.parent_id) {
    return null;
  }
  return findCatalogNodeById_(dataSheet, col, node.parent_id);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} col
 * @param {string} nodeId
 * @returns {Object|null}
 */
function findCatalogNodeById_(dataSheet, col, nodeId) {
  return findCatalogDataEntryById_(dataSheet, col, nodeId);
}

/**
 * @param {Object} parent
 * @returns {{permissions_json: string, editors: string, commenters: string, readers: string}}
 */
function inheritPermissionsFromParent_(parent, usersByEmail) {
  var editors = readCell_(parent.editors);
  var commenters = readCell_(parent.commenters);
  var readers = readCell_(parent.readers);
  if (!editors && !commenters && !readers) {
    var fromJson = permissionsJsonToDisplay_(parent.permissions_json, usersByEmail);
    editors = fromJson.editors;
    commenters = fromJson.commenters;
    readers = fromJson.readers;
  }
  return {
    permissions_json: parent.permissions_json || '[]',
    editors: editors || '',
    commenters: commenters || '',
    readers: readers || ''
  };
}

/**
 * Коды папок — прямые дети одного родителя (для автонумерации).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {string} parentId
 * @returns {string[]}
 */
function listDirectChildFolderCodes_(sheet, col, parentId) {
  var codes = [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2 || !parentId) {
    return codes;
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.CatalogData.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var values = rows[i];
    if (readCell_(values[col.parent_id - 1]) !== parentId) {
      continue;
    }
    if (readCell_(values[col.type - 1]) !== 'folder') {
      continue;
    }
    var code = readCell_(values[col.folder_code - 1]);
    if (code) {
      codes.push(code);
    }
  }
  return codes;
}

/**
 * @param {string[]} codes
 * @returns {Object.<string, boolean>}
 */
function folderCodesToSet_(codes) {
  var used = {};
  for (var i = 0; i < codes.length; i++) {
    var code = readCell_(codes[i]);
    if (code) {
      used[code] = true;
    }
  }
  return used;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} parent
 * @returns {string}
 */
function allocateFolderCode_(sheet, col, parent) {
  var parentCode = readCell_(parent.folder_code);
  var siblingCodes = listDirectChildFolderCodes_(sheet, col, parent.id);
  var used = folderCodesToSet_(siblingCodes);

  if (parentCode === FOLDER_CODE_CATALOG || parentCode === '') {
    for (var n = FOLDER_CODE_ROOT_MIN; n <= FOLDER_CODE_ROOT_MAX; n++) {
      var rootCode = String(n);
      if (!used[rootCode]) {
        return rootCode;
      }
    }
    throw new Error(
      'Лимит папок в корне исчерпан (' + FOLDER_CODE_ROOT_MIN + '–' + FOLDER_CODE_ROOT_MAX + ').'
    );
  }

  for (var d = 1; d <= 9; d++) {
    var childCode = parentCode + String(d);
    if (!used[childCode]) {
      return childCode;
    }
  }
  throw new Error('Лимит подпапок исчерпан (9).');
}

/**
 * @param {string} parentCode
 */
function assertCanAddSubfolder_(parentCode) {
  var depth = getFolderDisplayDepth_(parentCode);
  if (depth >= CATALOG_TREE_DEPTH_MAX) {
    throw new Error('Достигнута максимальная глубина вложенности (' + CATALOG_TREE_DEPTH_MAX + ').');
  }
}

/**
 * @param {string} label
 * @returns {string}
 */
function sanitizeFolderLabel_(label) {
  return String(label || '').trim().replace(/[/\\]/g, '-');
}

/**
 * @param {string} folderCode
 * @param {string} label
 * @returns {string}
 */
function buildFolderName_(folderCode, label) {
  return folderCode + '_' + sanitizeFolderLabel_(label);
}

/**
 * @param {string} fullName
 * @param {string} folderCode
 * @returns {string}
 */
function stripFolderPrefix_(fullName, folderCode) {
  if (!folderCode) {
    return fullName;
  }
  var prefix = folderCode + '_';
  if (fullName.indexOf(prefix) === 0) {
    return fullName.substring(prefix.length);
  }
  return fullName;
}

/**
 * Перемещает выделенные строки Catalog в виртуальную «Корзину» (файлы на Drive не удаляются).
 * Запуск: выделить строки на Catalog → «Каталогизатор → В корзину».
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function deleteCatalogItems(rootIdsJson) {
  try {
    var rootIds = null;
    if (rootIdsJson) {
      rootIds = JSON.parse(rootIdsJson);
    }
    return moveCatalogItemsToTrashCore_(rootIds);
  } catch (err) {
    return {
      ok: false,
      error: 'TRASH_MOVE_FAILED',
      message: String(err)
    };
  }
}

function deleteCatalogItemsMenu_() {
  assertCatalogImportNotBusy_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }

  try {
    getCatalogSheets_(ss);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Ошибка', String(err), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var selected;
  try {
    selected = resolveSelectedCatalogEntries_(ss);
  } catch (err) {
    SpreadsheetApp.getUi().alert('В корзину', String(err), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (!selected.length) {
    SpreadsheetApp.getUi().alert(
      'В корзину',
      'Выделите строки с папками или файлами (не заголовок).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var blocked = selected.filter(function (entry) {
    return isProtectedCatalogNode_(entry);
  });
  var movable = selected.filter(function (entry) {
    return !isProtectedCatalogNode_(entry);
  });

  if (!movable.length) {
    SpreadsheetApp.getUi().alert(
      'В корзину',
      'Нельзя переместить в корзину «Каталог» и «Корзину».',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var sheets = getCatalogSheets_(ss);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);

  var notInTrash = movable.filter(function (entry) {
    return !isNodeInTrashSubtree_(entry, trashId, nodesById);
  });
  var alreadyInTrash = movable.length - notInTrash.length;

  if (!notInTrash.length) {
    SpreadsheetApp.getUi().alert(
      'В корзину',
      'Выделенные элементы уже находятся в «Корзине».',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var plan = buildMoveCopyPlan_(notInTrash, allNodes);
  var preview = plan.preview.join('\n');
  var ui = SpreadsheetApp.getUi();
  var prompt = 'Переместить в корзину: папок — ' + plan.folderCount +
    ', файлов — ' + plan.fileCount + '.\n\nФайлы на Drive останутся до «Очистить корзину».\n\n' + preview;
  if (blocked.length) {
    prompt += '\n\nПропущено системных папок: ' + blocked.length + '.';
  }
  if (alreadyInTrash) {
    prompt += '\n\nУже в корзине (пропущено): ' + alreadyInTrash + '.';
  }

  if (ui.alert('Переместить в корзину?', prompt, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  var rootIds = plan.roots.map(function (node) {
    return node.id;
  });
  var result = moveCatalogItemsToTrashCore_(rootIds);
  ui.alert(result.ok ? 'Готово' : 'Ошибка', result.message || result.error, ui.ButtonSet.OK);
}

/**
 * @param {string[]=} rootIds
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function moveCatalogItemsToTrashCore_(rootIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  if (!trashId) {
    throw new Error('Не найдена системная папка «Корзина». Выполните setupSchema.');
  }

  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  var selected;
  if (rootIds && rootIds.length) {
    selected = resolveCatalogEntriesByIds_(rootIds, nodesById);
  } else {
    selected = resolveSelectedCatalogEntries_(ss).filter(function (entry) {
      return !isProtectedCatalogNode_(entry);
    });
  }
  if (!selected.length) {
    throw new Error('Выделите строки с папками или файлами на листе Catalog.');
  }

  var movable = selected.filter(function (entry) {
    return !isNodeInTrashSubtree_(entry, trashId, nodesById);
  });
  if (!movable.length) {
    throw new Error('Выделенные элементы уже находятся в «Корзине».');
  }

  var plan = buildMoveCopyPlan_(movable, allNodes);
  if (!plan.roots.length) {
    throw new Error('Не удалось определить элементы для перемещения в корзину.');
  }

  var moved = [];
  var failed = [];
  for (var i = 0; i < plan.roots.length; i++) {
    var root = plan.roots[i];
    try {
      moveCatalogItemToTrash_(sheets.data, dataCol, root, trashId);
      moved.push({
        name: root.name,
        type: root.type
      });
    } catch (err) {
      failed.push({
        name: root.name,
        type: root.type,
        message: String(err)
      });
    }
  }

  if (moved.length) {
    SpreadsheetApp.flush();
    renderCatalogViewLight_(sheets.view, sheets.data);
  }

  return {
    ok: failed.length === 0,
    data: {
      moved: moved,
      failed: failed,
      folder_count: plan.folderCount,
      file_count: plan.fileCount
    },
    message: buildTrashMoveMessage_(moved.length, failed.length, plan.folderCount, plan.fileCount)
  };
}

/**
 * @param {number} movedCount
 * @param {number} failedCount
 * @param {number} folderCount
 * @param {number} fileCount
 * @returns {string}
 */
function buildTrashMoveMessage_(movedCount, failedCount, folderCount, fileCount) {
  var parts = ['В корзину: корней ' + movedCount + ' (папок — ' + folderCount + ', файлов — ' + fileCount + ').'];
  if (failedCount) {
    parts.push('Ошибок: ' + failedCount + '.');
  }
  return parts.join(' ');
}

/**
 * Восстанавливает выделенные элементы из «Корзины».
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function restoreCatalogItems(rootIdsJson) {
  try {
    var rootIds = null;
    if (rootIdsJson) {
      rootIds = JSON.parse(rootIdsJson);
    }
    return restoreCatalogItemsCore_(rootIds);
  } catch (err) {
    return {
      ok: false,
      error: 'RESTORE_FAILED',
      message: String(err)
    };
  }
}

function restoreCatalogItemsMenu_() {
  assertCatalogImportNotBusy_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }

  try {
    getCatalogSheets_(ss);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Ошибка', String(err), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var selected;
  try {
    selected = resolveSelectedCatalogEntries_(ss);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Восстановление', String(err), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (!selected.length) {
    SpreadsheetApp.getUi().alert(
      'Восстановление',
      'Выделите элементы в «Корзине» (не заголовок).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var sheets = getCatalogSheets_(ss);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);

  var restorable = selected.filter(function (entry) {
    return !isProtectedCatalogNode_(entry) && isNodeInTrashSubtree_(entry, trashId, nodesById);
  });

  if (!restorable.length) {
    SpreadsheetApp.getUi().alert(
      'Восстановление',
      'Выделите папки или файлы из «Корзины» (не «Каталог» и не саму «Корзину»).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var plan = buildMoveCopyPlan_(restorable, allNodes);
  var preview = plan.preview.join('\n');
  var ui = SpreadsheetApp.getUi();
  var prompt = 'Восстановить: папок — ' + plan.folderCount + ', файлов — ' + plan.fileCount +
    '.\n\nЭлементы вернутся в прежнее место или в «Каталог», если родитель недоступен.\n\n' + preview;

  if (ui.alert('Восстановить из корзины?', prompt, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  var rootIds = plan.roots.map(function (node) {
    return node.id;
  });
  var result = restoreCatalogItemsCore_(rootIds);
  ui.alert(result.ok ? 'Готово' : 'Ошибка', result.message || result.error, ui.ButtonSet.OK);
}

/**
 * @param {string[]=} rootIds
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function restoreCatalogItemsCore_(rootIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  if (!trashId) {
    throw new Error('Не найдена системная папка «Корзина». Выполните setupSchema.');
  }

  var catalogRootId = findIdByFolderCode_(sheets.data, dataCol, FOLDER_CODE_CATALOG);
  if (!catalogRootId) {
    throw new Error('Не найдена системная папка «Каталог». Выполните setupSchema.');
  }

  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  var selected;
  if (rootIds && rootIds.length) {
    selected = resolveCatalogEntriesByIds_(rootIds, nodesById).filter(function (entry) {
      return !isProtectedCatalogNode_(entry);
    });
  } else {
    selected = resolveSelectedCatalogEntries_(ss).filter(function (entry) {
      return !isProtectedCatalogNode_(entry);
    });
  }
  if (!selected.length) {
    throw new Error('Выделите элементы в «Корзине» на листе Catalog.');
  }

  var restorable = selected.filter(function (entry) {
    return isNodeInTrashSubtree_(entry, trashId, nodesById);
  });
  if (!restorable.length) {
    throw new Error('Выделите папки или файлы из «Корзины».');
  }

  var plan = buildMoveCopyPlan_(restorable, allNodes);
  if (!plan.roots.length) {
    throw new Error('Не удалось определить элементы для восстановления.');
  }

  var restored = [];
  var failed = [];
  var reservedNames = {};

  plan.roots.sort(compareCatalogNodes_);
  for (var i = 0; i < plan.roots.length; i++) {
    var root = plan.roots[i];
    try {
      var restoreParentId = restoreCatalogItemFromTrash_(
        sheets.data,
        dataCol,
        allNodes,
        nodesById,
        root,
        trashId,
        catalogRootId,
        reservedNames
      );

      restored.push({
        name: root.name,
        type: root.type,
        restore_parent_id: restoreParentId
      });
    } catch (err) {
      failed.push({
        name: root.name,
        type: root.type,
        message: String(err)
      });
    }
  }

  if (restored.length) {
    SpreadsheetApp.flush();
    renderCatalogViewLight_(sheets.view, sheets.data);
  }

  return {
    ok: failed.length === 0,
    data: {
      restored: restored,
      failed: failed,
      folder_count: plan.folderCount,
      file_count: plan.fileCount
    },
    message: buildRestoreCatalogItemsMessage_(restored.length, failed.length, plan.folderCount, plan.fileCount)
  };
}

/**
 * @param {number} restoredCount
 * @param {number} failedCount
 * @param {number} folderCount
 * @param {number} fileCount
 * @returns {string}
 */
function buildRestoreCatalogItemsMessage_(restoredCount, failedCount, folderCount, fileCount) {
  var parts = ['Восстановлено: корней ' + restoredCount + ' (папок — ' + folderCount + ', файлов — ' + fileCount + ').'];
  if (failedCount) {
    parts.push('Ошибок: ' + failedCount + '.');
  }
  return parts.join(' ');
}

/**
 * Физически удаляет все элементы из «Корзины» (Drive + CatalogData). Только владелец таблицы.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function emptyTrashCatalog() {
  try {
    return emptyTrashCatalogCore_();
  } catch (err) {
    return {
      ok: false,
      error: 'EMPTY_TRASH_FAILED',
      message: String(err)
    };
  }
}

function emptyTrashCatalogMenu_() {
  assertCatalogImportNotBusy_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }

  if (!isSpreadsheetOwner_(ss)) {
    SpreadsheetApp.getUi().alert(
      'Очистить корзину',
      'Очистить корзину может только владелец таблицы.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  try {
    getCatalogSheets_(ss);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Ошибка', String(err), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var sheets = getCatalogSheets_(ss);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  var inTrash = collectTrashSubtreeNodes_(allNodes, trashId, nodesById);

  if (!inTrash.length) {
    SpreadsheetApp.getUi().alert('Очистить корзину', '«Корзина» пуста.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var fileCount = 0;
  var folderCount = 0;
  for (var i = 0; i < inTrash.length; i++) {
    if (inTrash[i].type === 'folder') {
      folderCount++;
    } else {
      fileCount++;
    }
  }

  var preview = inTrash.slice(0, 12).map(function (node) {
    return '• ' + (node.name || node.type || node.id);
  }).join('\n');
  if (inTrash.length > 12) {
    preview += '\n… и ещё ' + (inTrash.length - 12);
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = 'Безвозвратно удалить из каталога и с Drive:\nпапок — ' + folderCount +
    ', файлов — ' + fileCount + '.\n\n' + preview;
  if (ui.alert('Очистить корзину?', prompt, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  var result = emptyTrashCatalog();
  ui.alert(result.ok ? 'Готово' : 'Ошибка', result.message || result.error, ui.ButtonSet.OK);
}

/**
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function emptyTrashCatalogCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isSpreadsheetOwner_(ss)) {
    throw new Error('Очистить корзину может только владелец таблицы.');
  }

  var sheets = getCatalogSheets_(ss);
  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var trashId = getTrashFolderId_(sheets.data, dataCol);
  if (!trashId) {
    throw new Error('Не найдена системная папка «Корзина». Выполните setupSchema.');
  }

  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  var toDelete = collectTrashSubtreeNodes_(allNodes, trashId, nodesById);
  if (!toDelete.length) {
    return {
      ok: true,
      data: { deleted: [], failed: [], folder_count: 0, file_count: 0 },
      message: '«Корзина» пуста.'
    };
  }

  toDelete.sort(function (a, b) {
    return b.rowIndex - a.rowIndex;
  });

  var deleted = [];
  var failed = [];
  for (var i = 0; i < toDelete.length; i++) {
    var node = toDelete[i];
    try {
      if (node.file_id && node.type !== 'folder') {
        deleteDriveFile_(node.file_id);
      }
      sheets.data.deleteRow(node.rowIndex);
      deleted.push({
        name: node.name,
        type: node.type,
        file_id: node.file_id
      });
    } catch (err) {
      failed.push({
        name: node.name,
        type: node.type,
        message: String(err)
      });
    }
  }

  if (deleted.length) {
    renderCatalogViewCore_(sheets.view, sheets.data);
  }

  var fileCount = 0;
  var folderCount = 0;
  for (var d = 0; d < deleted.length; d++) {
    if (deleted[d].type === 'folder') {
      folderCount++;
    } else {
      fileCount++;
    }
  }

  return {
    ok: failed.length === 0,
    data: {
      deleted: deleted,
      failed: failed,
      folder_count: folderCount,
      file_count: fileCount
    },
    message: buildEmptyTrashMessage_(deleted.length, failed.length, folderCount, fileCount)
  };
}

/**
 * @param {number} deletedCount
 * @param {number} failedCount
 * @param {number} folderCount
 * @param {number} fileCount
 * @returns {string}
 */
function buildEmptyTrashMessage_(deletedCount, failedCount, folderCount, fileCount) {
  var parts = ['Удалено безвозвратно: ' + deletedCount + ' (папок — ' + folderCount + ', файлов — ' + fileCount + ').'];
  if (failedCount) {
    parts.push('Ошибок: ' + failedCount + '.');
  }
  return parts.join(' ');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} col
 * @returns {string}
 */
function getTrashFolderId_(dataSheet, col) {
  return findIdByFolderCode_(dataSheet, col, FOLDER_CODE_TRASH);
}

/**
 * @param {Object} node
 * @param {string} trashId
 * @param {Object.<string, Object>} nodesById
 * @returns {boolean}
 */
function isNodeInTrashSubtree_(node, trashId, nodesById) {
  if (!node || !node.id) {
    return false;
  }
  if (node.id === trashId || node.folder_code === FOLDER_CODE_TRASH) {
    return true;
  }
  var current = node;
  var guard = 0;
  while (current && current.parent_id && guard < 200) {
    if (current.parent_id === trashId) {
      return true;
    }
    current = nodesById[current.parent_id];
    if (!current) {
      return false;
    }
    if (current.folder_code === FOLDER_CODE_TRASH) {
      return true;
    }
    guard++;
  }
  return false;
}

/**
 * @param {Array} allNodes
 * @param {string} trashId
 * @param {Object.<string, Object>} nodesById
 * @returns {Array}
 */
function collectTrashSubtreeNodes_(allNodes, trashId, nodesById) {
  var result = [];
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (!node.id || node.id === trashId || isProtectedCatalogNode_(node)) {
      continue;
    }
    if (isNodeInTrashSubtree_(node, trashId, nodesById)) {
      result.push(node);
    }
  }
  return result;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {string} trashId
 */
function moveCatalogItemToTrash_(sheet, col, node, trashId) {
  if (node.parent_id === trashId) {
    return;
  }
  writeCatalogDataNodeFields_(sheet, col, node, {
    trash_parent_id: node.parent_id,
    parent_id: trashId
  });
}

/**
 * Целевой родитель при восстановлении: trash_parent_id или «Каталог», если родитель недоступен.
 *
 * @param {Object} root
 * @param {string} catalogRootId
 * @param {string} trashId
 * @param {Object.<string, Object>} nodesById
 * @returns {{id: string, node: Object|null}}
 */
function resolveRestoreParent_(root, catalogRootId, trashId, nodesById) {
  var restoreParentId = readCell_(root.trash_parent_id);
  var restoreParent = restoreParentId ? nodesById[restoreParentId] : null;
  if (!restoreParent || isNodeInTrashSubtree_(restoreParent, trashId, nodesById)) {
    restoreParentId = catalogRootId;
    restoreParent = nodesById[catalogRootId] || null;
  }
  return { id: restoreParentId, node: restoreParent };
}

/**
 * @param {string} parentId
 * @param {string} name
 * @param {string} nodeId
 * @param {Array} allNodes
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 * @returns {boolean}
 */
function isSiblingCatalogNameTaken_(parentId, name, nodeId, allNodes, reservedNames) {
  if (!name) {
    return false;
  }
  var excludeIds = {};
  excludeIds[nodeId] = true;
  var used = listSiblingFullNames_(parentId, allNodes, excludeIds);
  if (used[name]) {
    return true;
  }
  return !!(reservedNames[parentId] && reservedNames[parentId][name]);
}

/**
 * @param {string} parentId
 * @param {string} folderCode
 * @param {string} nodeId
 * @param {Array} allNodes
 * @returns {boolean}
 */
function isSiblingFolderCodeTaken_(parentId, folderCode, nodeId, allNodes) {
  folderCode = readCell_(folderCode);
  if (!folderCode) {
    return false;
  }
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.id === nodeId || node.parent_id !== parentId || node.type !== 'folder') {
      continue;
    }
    if (readCell_(node.folder_code) === folderCode) {
      return true;
    }
  }
  return false;
}

/**
 * Быстрое восстановление возможно, если достаточно вернуть parent_id (как при отправке в корзину).
 *
 * @param {Object} root
 * @param {string} restoreParentId
 * @param {Array} allNodes
 * @param {Object.<string, Object>} nodesById
 * @param {string} trashId
 * @param {string} catalogRootId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 * @returns {boolean}
 */
function canFastRestoreCatalogItem_(root, restoreParentId, allNodes, nodesById, trashId, catalogRootId, reservedNames) {
  if (!restoreParentId || !nodesById[restoreParentId]) {
    return false;
  }
  if (isNodeInTrashSubtree_(nodesById[restoreParentId], trashId, nodesById)) {
    return false;
  }
  if (isSiblingCatalogNameTaken_(restoreParentId, root.name, root.id, allNodes, reservedNames)) {
    return false;
  }
  if (root.type === 'folder') {
    if (isSiblingFolderCodeTaken_(restoreParentId, root.folder_code, root.id, allNodes)) {
      return false;
    }
    if (restoreParentId === catalogRootId) {
      try {
        assertSubtreeFitsDepth_(readCell_(root.folder_code), root.id, allNodes);
      } catch (depthErr) {
        return false;
      }
    }
  }
  return true;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {string} restoreParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function fastRestoreCatalogItem_(sheet, col, node, restoreParentId, reservedNames) {
  writeCatalogDataNodeFields_(sheet, col, node, {
    parent_id: restoreParentId,
    trash_parent_id: ''
  });
  if (!reservedNames[restoreParentId]) {
    reservedNames[restoreParentId] = {};
  }
  reservedNames[restoreParentId][node.name] = true;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {Object.<string, Object>} nodesById
 * @param {Object} root
 * @param {string} trashId
 * @param {string} catalogRootId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 * @returns {string}
 */
function restoreCatalogItemFromTrash_(sheet, col, allNodes, nodesById, root, trashId, catalogRootId, reservedNames) {
  var parentInfo = resolveRestoreParent_(root, catalogRootId, trashId, nodesById);
  var restoreParentId = parentInfo.id;
  var restoreParent = parentInfo.node;
  if (!restoreParent) {
    throw new Error('Не найдена папка для восстановления.');
  }

  if (canFastRestoreCatalogItem_(root, restoreParentId, allNodes, nodesById, trashId, catalogRootId, reservedNames)) {
    fastRestoreCatalogItem_(sheet, col, root, restoreParentId, reservedNames);
    return restoreParentId;
  }

  if (root.type === 'folder') {
    moveCatalogFolderRoot_(sheet, col, allNodes, root, restoreParent, restoreParentId, reservedNames);
  } else {
    moveCatalogFileNode_(sheet, col, allNodes, root, restoreParent, restoreParentId, reservedNames);
  }
  writeCatalogDataNodeFields_(sheet, col, root, { trash_parent_id: '' });
  return restoreParentId;
}

/**
 * Открывает диалог переноса / копирования выделенных элементов Catalog.
 */
function moveCopyDialog_() {
  var contextResult = getMoveCopyContext();
  var template = HtmlService.createTemplateFromFile('MoveCopyDialog');
  template.initialContext = JSON.stringify(contextResult);
  var html = template.evaluate()
    .setWidth(520)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Переместить / Копировать');
}

/**
 * Контекст для диалога переноса/копирования.
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function getMoveCopyContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = getCatalogSheets_(ss);

    var headerCheck = validateCatalogDataHeaders_(sheets.data);
    if (!headerCheck.ok) {
      return headerCheck;
    }

    if (!ss.getActiveSheet() || ss.getActiveSheet().getName() !== 'Catalog') {
      return {
        ok: false,
        error: 'WRONG_SHEET',
        message: 'Выберите строки на листе Catalog.'
      };
    }

    var selected = resolveSelectedCatalogEntries_(ss);
    if (!selected.length) {
      return {
        ok: false,
        error: 'NO_SELECTION',
        message: 'Выделите папки или файлы на листе Catalog.'
      };
    }

    var movable = selected.filter(function (entry) {
      return !isProtectedCatalogNode_(entry);
    });
    if (!movable.length) {
      return {
        ok: false,
        error: 'PROTECTED_ONLY',
        message: 'Нельзя перемещать или копировать «Каталог» и «Корзину».'
      };
    }

    var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
    var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
    var plan = buildMoveCopyPlan_(movable, allNodes);
    if (!plan.roots.length) {
      return {
        ok: false,
        error: 'EMPTY_PLAN',
        message: 'Не удалось определить элементы для операции.'
      };
    }

    var trashId = getTrashFolderId_(sheets.data, dataCol);
    var nodesById = mapCatalogNodesById_(allNodes);
    for (var t = 0; t < plan.roots.length; t++) {
      if (isNodeInTrashSubtree_(plan.roots[t], trashId, nodesById)) {
        return {
          ok: false,
          error: 'TRASH_SELECTION',
          message: 'Элементы из «Корзины» восстанавливают через меню «Восстановить из корзины».'
        };
      }
    }

    return {
      ok: true,
      data: {
        folder_count: plan.folderCount,
        file_count: plan.fileCount,
        summary_message: buildMoveCopySummaryMessage_(plan.folderCount, plan.fileCount),
        preview: plan.preview,
        target_folders: listMoveCopyTargetFolders_(plan.excludedTargetIds, sheets.data, dataCol)
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'MOVE_COPY_CONTEXT_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} mode move | copy
 * @param {string} targetParentId
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function moveCopyCatalogItems(mode, targetParentId) {
  assertCatalogImportNotBusy_();
  try {
    return moveCopyCatalogItemsCore_(mode, targetParentId);
  } catch (err) {
    return {
      ok: false,
      error: 'MOVE_COPY_FAILED',
      message: String(err)
    };
  }
}

/**
 * @param {string} mode
 * @param {string} targetParentId
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function moveCopyCatalogItemsCore_(mode, targetParentId) {
  var normalizedMode = String(mode || '').toLowerCase();
  if (normalizedMode !== 'move' && normalizedMode !== 'copy') {
    throw new Error('Неверный режим операции.');
  }
  if (!targetParentId) {
    throw new Error('Выберите целевую папку.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getCatalogSheets_(ss);

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    throw new Error(headerCheck.message || headerCheck.error);
  }

  var selected = resolveSelectedCatalogEntries_(ss).filter(function (entry) {
    return !isProtectedCatalogNode_(entry);
  });
  if (!selected.length) {
    throw new Error('Выделите папки или файлы на листе Catalog.');
  }

  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var allNodes = readCatalogDataNodes_(sheets.data, dataCol);
  var plan = buildMoveCopyPlan_(selected, allNodes);
  if (!plan.roots.length) {
    throw new Error('Нет элементов для операции.');
  }

  var trashId = getTrashFolderId_(sheets.data, dataCol);
  var nodesById = mapCatalogNodesById_(allNodes);
  for (var t = 0; t < plan.roots.length; t++) {
    if (isNodeInTrashSubtree_(plan.roots[t], trashId, nodesById)) {
      throw new Error('Элементы из «Корзины» восстанавливают через меню «Восстановить из корзины».');
    }
  }

  var targetParent = findCatalogDataEntryById_(sheets.data, dataCol, targetParentId);
  if (!targetParent || targetParent.type !== 'folder') {
    throw new Error('Целевая папка не найдена.');
  }
  if (targetParent.folder_code === FOLDER_CODE_TRASH) {
    throw new Error('Нельзя перемещать или копировать в «Корзину».');
  }
  if (plan.excludedTargetIds[targetParentId]) {
    throw new Error('Нельзя выбрать эту папку: она входит в выделение или его поддерево.');
  }

  if (normalizedMode === 'move') {
    for (var i = 0; i < plan.roots.length; i++) {
      if (plan.roots[i].parent_id === targetParentId) {
        throw new Error('Элемент «' + plan.roots[i].name + '» уже находится в выбранной папке.');
      }
    }
  }

  var targetParentId = targetParent.id;
  var reservedNames = {};
  var processed = { folders: 0, files: 0 };

  plan.roots.sort(compareCatalogNodes_);
  for (var r = 0; r < plan.roots.length; r++) {
    var root = plan.roots[r];
    if (normalizedMode === 'move') {
      if (root.type === 'folder') {
        moveCatalogFolderRoot_(sheets.data, dataCol, allNodes, root, targetParent, targetParentId, reservedNames);
        processed.folders++;
      } else {
        moveCatalogFileNode_(sheets.data, dataCol, allNodes, root, targetParent, targetParentId, reservedNames);
        processed.files++;
      }
    } else if (root.type === 'folder') {
      copyCatalogFolderRoot_(sheets.data, dataCol, allNodes, root, targetParent, targetParentId, reservedNames);
      processed.folders++;
    } else {
      copyCatalogFileNode_(sheets.data, dataCol, allNodes, root, targetParent, targetParentId, reservedNames);
      processed.files++;
    }
  }

  SpreadsheetApp.flush();
  renderCatalogViewCore_(sheets.view, sheets.data);

  var actionLabel = normalizedMode === 'move' ? 'Перенесено' : 'Скопировано';
  return {
    ok: true,
    data: {
      mode: normalizedMode,
      target_parent_id: targetParentId,
      folders: processed.folders,
      files: processed.files,
      affected_folders: plan.folderCount,
      affected_files: plan.fileCount
    },
    message: actionLabel + ': папок ' + plan.folderCount + ', файлов ' + plan.fileCount + '.'
  };
}

/**
 * @param {Array} selectedEntries
 * @param {Array} allNodes
 * @returns {{roots: Array, affected: Array, folderCount: number, fileCount: number, excludedTargetIds: Object.<string, boolean>, preview: string[]}}
 */
function buildMoveCopyPlan_(selectedEntries, allNodes) {
  var nodesById = mapCatalogNodesById_(allNodes);
  var childrenByParent = mapCatalogChildrenByParent_(allNodes);
  var selectedIds = {};
  for (var i = 0; i < selectedEntries.length; i++) {
    selectedIds[selectedEntries[i].id] = true;
  }

  var roots = [];
  for (var s = 0; s < selectedEntries.length; s++) {
    var entry = selectedEntries[s];
    if (!isDescendantOfSelectedNode_(entry.id, selectedIds, nodesById)) {
      roots.push(entry);
    }
  }

  var affected = [];
  var affectedIds = {};
  var excludedTargetIds = {};
  for (var r = 0; r < roots.length; r++) {
    var root = roots[r];
    affectedIds[root.id] = true;
    excludedTargetIds[root.id] = true;
    if (root.type === 'folder') {
      collectDescendantNodeIds_(root.id, childrenByParent, affectedIds);
      collectDescendantNodeIds_(root.id, childrenByParent, excludedTargetIds);
    }
  }

  for (var a = 0; a < allNodes.length; a++) {
    if (affectedIds[allNodes[a].id]) {
      affected.push(allNodes[a]);
    }
  }
  affected.sort(compareCatalogNodes_);

  var folderCount = 0;
  var fileCount = 0;
  for (var f = 0; f < affected.length; f++) {
    if (affected[f].type === 'folder') {
      folderCount++;
    } else {
      fileCount++;
    }
  }

  var preview = affected.slice(0, 12).map(function (node) {
    return '• ' + (node.name || node.type || node.id);
  });
  if (affected.length > 12) {
    preview.push('… и ещё ' + (affected.length - 12));
  }

  return {
    roots: roots,
    affected: affected,
    folderCount: folderCount,
    fileCount: fileCount,
    excludedTargetIds: excludedTargetIds,
    preview: preview
  };
}

/**
 * @param {number} folderCount
 * @param {number} fileCount
 * @returns {string}
 */
function buildMoveCopySummaryMessage_(folderCount, fileCount) {
  return 'Будет затронуто: папок — ' + folderCount + ', файлов — ' + fileCount + '.';
}

/**
 * @param {Object.<string, boolean>} excludedTargetIds
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {Object.<string, number>} dataCol
 * @returns {Array<{id: string, label: string}>}
 */
function listMoveCopyTargetFolders_(excludedTargetIds, dataSheet, dataCol) {
  var tree = buildCatalogTreeState_(dataSheet, dataCol);
  if (!tree) {
    return [];
  }

  var options = [];
  for (var i = 0; i < tree.sorted.length; i++) {
    var node = tree.sorted[i];
    if (node.type !== 'folder') {
      continue;
    }
    if (node.folder_code === FOLDER_CODE_TRASH) {
      continue;
    }
    if (excludedTargetIds[node.id]) {
      continue;
    }
    options.push({
      id: node.id,
      label: formatMoveCopyFolderOptionLabel_(node),
      folder_code: readCell_(node.folder_code),
      name: readCell_(node.name)
    });
  }
  return options;
}

/**
 * @param {Object} node
 * @returns {string}
 */
function formatMoveCopyFolderOptionLabel_(node) {
  var depth = node.folder_code === FOLDER_CODE_CATALOG
    ? 0
    : getFolderDisplayDepth_(node.folder_code);
  var indent = depth > 0 ? Array(depth + 1).join('— ') : '';
  return indent + (node.name || node.id);
}

/**
 * @param {Array} allNodes
 * @returns {Object.<string, Object>}
 */
function mapCatalogNodesById_(allNodes) {
  var nodesById = {};
  for (var i = 0; i < allNodes.length; i++) {
    if (allNodes[i].id) {
      nodesById[allNodes[i].id] = allNodes[i];
    }
  }
  return nodesById;
}

/**
 * @param {Array} allNodes
 * @returns {Object.<string, string[]>}
 */
function mapCatalogChildrenByParent_(allNodes) {
  var childrenByParent = {};
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (!node.id) {
      continue;
    }
    var parentKey = node.parent_id || '';
    if (!childrenByParent[parentKey]) {
      childrenByParent[parentKey] = [];
    }
    childrenByParent[parentKey].push(node.id);
  }
  return childrenByParent;
}

/**
 * @param {string} nodeId
 * @param {Object.<string, boolean>} selectedIds
 * @param {Object.<string, Object>} nodesById
 * @returns {boolean}
 */
function isDescendantOfSelectedNode_(nodeId, selectedIds, nodesById) {
  var current = nodesById[nodeId];
  while (current && current.parent_id) {
    if (selectedIds[current.parent_id]) {
      return true;
    }
    current = nodesById[current.parent_id];
  }
  return false;
}

/**
 * @param {string} rootId
 * @param {Object.<string, string[]>} childrenByParent
 * @param {Object.<string, boolean>} result
 */
function collectDescendantNodeIds_(rootId, childrenByParent, result) {
  var children = childrenByParent[rootId] || [];
  for (var i = 0; i < children.length; i++) {
    result[children[i]] = true;
    collectDescendantNodeIds_(children[i], childrenByParent, result);
  }
}

/**
 * @param {string} parentId
 * @param {Array} allNodes
 * @param {Object.<string, boolean>} excludeIds
 * @returns {Object.<string, boolean>}
 */
function listSiblingFullNames_(parentId, allNodes, excludeIds) {
  var used = {};
  for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.parent_id !== parentId) {
      continue;
    }
    if (excludeIds && excludeIds[node.id]) {
      continue;
    }
    if (node.name) {
      used[node.name] = true;
    }
  }
  return used;
}

/**
 * @param {string} parentId
 * @param {string} folderCode
 * @param {string} baseLabel
 * @param {Array} allNodes
 * @param {Object.<string, boolean>} excludeIds
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 * @returns {{label: string, name: string}}
 */
function allocateUniqueCatalogName_(parentId, folderCode, baseLabel, allNodes, excludeIds, reservedNames) {
  var used = listSiblingFullNames_(parentId, allNodes, excludeIds);
  if (reservedNames[parentId]) {
    var reserved = reservedNames[parentId];
    var reservedKeys = Object.keys(reserved);
    for (var i = 0; i < reservedKeys.length; i++) {
      used[reservedKeys[i]] = true;
    }
  }

  var label = sanitizeFolderLabel_(baseLabel);
  var suffix = 2;
  while (true) {
    var fullName = buildCatalogFullName_(folderCode, label);
    if (!used[fullName]) {
      if (!reservedNames[parentId]) {
        reservedNames[parentId] = {};
      }
      reservedNames[parentId][fullName] = true;
      return { label: label, name: fullName };
    }
    label = appendAutoSuffixToLabel_(baseLabel, suffix);
    suffix++;
  }
}

/**
 * @param {string} baseLabel
 * @param {number} suffix
 * @returns {string}
 */
function appendAutoSuffixToLabel_(baseLabel, suffix) {
  var label = sanitizeFolderLabel_(baseLabel);
  var dot = label.lastIndexOf('.');
  if (dot > 0) {
    return label.substring(0, dot) + ' (' + suffix + ')' + label.substring(dot);
  }
  return label + ' (' + suffix + ')';
}

/**
 * @param {string} parentId
 * @param {Array} allNodes
 * @returns {Array}
 */
function getSortedChildrenOfNode_(parentId, allNodes) {
  var children = [];
  for (var i = 0; i < allNodes.length; i++) {
    if (allNodes[i].parent_id === parentId) {
      children.push(allNodes[i]);
    }
  }
  children.sort(compareCatalogNodes_);
  return children;
}

/**
 * @param {Object} targetParent
 * @returns {string}
 */
function getChildFileFolderCode_(targetParent) {
  if (targetParent.folder_code === FOLDER_CODE_CATALOG) {
    return '';
  }
  return readCell_(targetParent.folder_code);
}

/**
 * @param {string} folderCode
 * @param {string} folderId
 * @param {Array} allNodes
 */
function assertSubtreeFitsDepth_(folderCode, folderId, allNodes) {
  var baseDepth = getFolderDisplayDepth_(folderCode);
  var below = getMaxFolderDepthBelow_(folderId, allNodes);
  if (baseDepth + below > CATALOG_TREE_DEPTH_MAX) {
    throw new Error(
      'Превышена максимальная глубина вложенности (' + CATALOG_TREE_DEPTH_MAX + ').'
    );
  }
}

/**
 * @param {string} folderId
 * @param {Array} allNodes
 * @returns {number}
 */
function getMaxFolderDepthBelow_(folderId, allNodes) {
  var children = getSortedChildrenOfNode_(folderId, allNodes);
  var max = 0;
  for (var i = 0; i < children.length; i++) {
    if (children[i].type !== 'folder') {
      continue;
    }
    var depth = 1 + getMaxFolderDepthBelow_(children[i].id, allNodes);
    if (depth > max) {
      max = depth;
    }
  }
  return max;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {Object} fields
 */
function writeCatalogDataNodeFields_(sheet, col, node, fields) {
  if (fields.parent_id !== undefined) {
    sheet.getRange(node.rowIndex, col.parent_id).setValue(readCell_(fields.parent_id));
    node.parent_id = readCell_(fields.parent_id);
  }
  if (fields.folder_code !== undefined) {
    sheet.getRange(node.rowIndex, col.folder_code).setValue(readCell_(fields.folder_code));
    node.folder_code = readCell_(fields.folder_code);
  }
  if (fields.name !== undefined) {
    sheet.getRange(node.rowIndex, col.name).setValue(fields.name);
    node.name = fields.name;
  }
  if (fields.file_id !== undefined) {
    sheet.getRange(node.rowIndex, col.file_id).setValue(fields.file_id);
    node.file_id = fields.file_id;
  }
  if (fields.trash_parent_id !== undefined) {
    sheet.getRange(node.rowIndex, col.trash_parent_id).setValue(readCell_(fields.trash_parent_id));
    node.trash_parent_id = readCell_(fields.trash_parent_id);
  }
}

/**
 * Права целевой папки → строка CatalogData (читаем родителя заново с листа).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {string} targetParentId
 */
function applyTargetFolderPermissions_(sheet, col, node, targetParentId) {
  if (node.type === 'file' && isCatalogNodeApproved_(node)) {
    return;
  }
  var parent = findCatalogDataEntryById_(sheet, col, targetParentId);
  if (!parent) {
    return;
  }
  var usersMap = readUsersMapFromSpreadsheet_(sheet.getParent());
  applyInheritedPermissionsToNode_(sheet, col, node, inheritPermissionsFromParent_(parent, usersMap));
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {string} targetParentId
 * @returns {Object}
 */
function readTargetFolderPermissions_(sheet, col, targetParentId) {
  var parent = findCatalogDataEntryById_(sheet, col, targetParentId);
  if (!parent) {
    return inheritPermissionsFromParent_({ permissions_json: '[]' }, {});
  }
  var usersMap = readUsersMapFromSpreadsheet_(sheet.getParent());
  return inheritPermissionsFromParent_(parent, usersMap);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Object} node
 * @param {Object} inherited
 */
function applyInheritedPermissionsToNode_(sheet, col, node, inherited) {
  if (!inherited) {
    return;
  }
  sheet.getRange(node.rowIndex, col.permissions_json).setValue(inherited.permissions_json || '[]');
  sheet.getRange(node.rowIndex, col.editors).setValue(inherited.editors || '');
  sheet.getRange(node.rowIndex, col.commenters).setValue(inherited.commenters || '');
  sheet.getRange(node.rowIndex, col.readers).setValue(inherited.readers || '');
  sheet.getRange(node.rowIndex, col.approved).setValue('');
  node.permissions_json = inherited.permissions_json || '[]';
  node.editors = inherited.editors || '';
  node.commenters = inherited.commenters || '';
  node.readers = inherited.readers || '';
  node.approved = '';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {Object} fileNode
 * @param {Object} targetParent
 * @param {Object} targetParent
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function moveCatalogFileNode_(sheet, col, allNodes, fileNode, targetParent, targetParentId, reservedNames) {
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var baseLabel = stripFolderPrefix_(fileNode.name, fileNode.folder_code);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    baseLabel,
    allNodes,
    {},
    reservedNames
  );

  if (fileNode.file_id) {
    Drive.Files.update({ name: unique.name }, fileNode.file_id, null, {
      supportsAllDrives: true,
      fields: 'id,name'
    });
  }

  writeCatalogDataNodeFields_(sheet, col, fileNode, {
    parent_id: targetParent.id,
    folder_code: fileFolderCode,
    name: unique.name
  });
  applyTargetFolderPermissions_(sheet, col, fileNode, targetParentId);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {Object} folderNode
 * @param {Object} targetParent
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function moveCatalogFolderRoot_(sheet, col, allNodes, folderNode, targetParent, targetParentId, reservedNames) {
  assertCanAddSubfolder_(readCell_(targetParent.folder_code));
  var newCode = allocateFolderCode_(sheet, col, targetParent);
  assertSubtreeFitsDepth_(newCode, folderNode.id, allNodes);
  var baseLabel = stripFolderPrefix_(folderNode.name, folderNode.folder_code);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    newCode,
    baseLabel,
    allNodes,
    {},
    reservedNames
  );

  writeCatalogDataNodeFields_(sheet, col, folderNode, {
    parent_id: targetParent.id,
    folder_code: newCode,
    name: unique.name
  });
  applyTargetFolderPermissions_(sheet, col, folderNode, targetParentId);

  recodeMovedFolderChildren_(sheet, col, allNodes, folderNode.id, newCode, targetParentId, reservedNames);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {string} folderId
 * @param {string} parentFolderCode
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function recodeMovedFolderChildren_(sheet, col, allNodes, folderId, parentFolderCode, targetParentId, reservedNames) {
  var children = getSortedChildrenOfNode_(folderId, allNodes);
  var folderDigit = 1;

  for (var i = 0; i < children.length; i++) {
    var child = children[i];
    if (child.type === 'folder') {
      var childCode = parentFolderCode + String(folderDigit);
      folderDigit++;
      var childLabel = stripFolderPrefix_(child.name, child.folder_code);
      var childUnique = allocateUniqueCatalogName_(
        folderId,
        childCode,
        childLabel,
        allNodes,
        {},
        reservedNames
      );

      writeCatalogDataNodeFields_(sheet, col, child, {
        folder_code: childCode,
        name: childUnique.name
      });
      applyTargetFolderPermissions_(sheet, col, child, targetParentId);
      recodeMovedFolderChildren_(sheet, col, allNodes, child.id, childCode, targetParentId, reservedNames);
    } else {
      var fileFolderCode = parentFolderCode;
      var fileLabel = stripFolderPrefix_(child.name, child.folder_code);
      var fileUnique = allocateUniqueCatalogName_(
        folderId,
        fileFolderCode,
        fileLabel,
        allNodes,
        {},
        reservedNames
      );

      if (child.file_id) {
        Drive.Files.update({ name: fileUnique.name }, child.file_id, null, {
          supportsAllDrives: true,
          fields: 'id,name'
        });
      }

      writeCatalogDataNodeFields_(sheet, col, child, {
        folder_code: fileFolderCode,
        name: fileUnique.name
      });
      applyTargetFolderPermissions_(sheet, col, child, targetParentId);
    }
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {Object} fileNode
 * @param {Object} targetParent
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function copyCatalogFileNode_(sheet, col, allNodes, fileNode, targetParent, targetParentId, reservedNames) {
  var fileFolderCode = getChildFileFolderCode_(targetParent);
  var baseLabel = stripFolderPrefix_(fileNode.name, fileNode.folder_code);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    fileFolderCode,
    baseLabel,
    allNodes,
    {},
    reservedNames
  );

  var newFileId = '';
  if (fileNode.file_id) {
    newFileId = copyDriveFile_(fileNode.file_id, unique.name);
  }

  var inherited = readTargetFolderPermissions_(sheet, col, targetParentId);
  appendCatalogDataRow_(sheet, col, {
    id: generateId_(),
    parent_id: targetParent.id,
    folder_code: fileFolderCode,
    name: unique.name,
    type: 'file',
    file_id: newFileId,
    mime_type: fileNode.mime_type || '',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  });
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {Object} folderNode
 * @param {Object} targetParent
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function copyCatalogFolderRoot_(sheet, col, allNodes, folderNode, targetParent, targetParentId, reservedNames) {
  assertCanAddSubfolder_(readCell_(targetParent.folder_code));
  var newCode = allocateFolderCode_(sheet, col, targetParent);
  assertSubtreeFitsDepth_(newCode, folderNode.id, allNodes);
  var baseLabel = stripFolderPrefix_(folderNode.name, folderNode.folder_code);
  var unique = allocateUniqueCatalogName_(
    targetParent.id,
    newCode,
    baseLabel,
    allNodes,
    {},
    reservedNames
  );

  var inherited = readTargetFolderPermissions_(sheet, col, targetParentId);
  var newFolderId = generateId_();
  appendCatalogDataRow_(sheet, col, {
    id: newFolderId,
    parent_id: targetParent.id,
    folder_code: newCode,
    name: unique.name,
    type: 'folder',
    permissions_json: inherited.permissions_json,
    editors: inherited.editors,
    commenters: inherited.commenters,
    readers: inherited.readers
  });

  copyCatalogFolderChildren_(sheet, col, allNodes, folderNode.id, newFolderId, newCode, targetParentId, reservedNames);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @param {Array} allNodes
 * @param {string} sourceFolderId
 * @param {string} targetFolderId
 * @param {string} targetFolderCode
 * @param {string} targetParentId
 * @param {Object.<string, Object.<string, boolean>>} reservedNames
 */
function copyCatalogFolderChildren_(sheet, col, allNodes, sourceFolderId, targetFolderId, targetFolderCode, targetParentId, reservedNames) {
  var children = getSortedChildrenOfNode_(sourceFolderId, allNodes);
  var folderDigit = 1;
  var inherited = readTargetFolderPermissions_(sheet, col, targetParentId);

  for (var i = 0; i < children.length; i++) {
    var child = children[i];
    if (child.type === 'folder') {
      var childCode = targetFolderCode + String(folderDigit);
      folderDigit++;
      var childLabel = stripFolderPrefix_(child.name, child.folder_code);
      var childUnique = allocateUniqueCatalogName_(
        targetFolderId,
        childCode,
        childLabel,
        allNodes,
        {},
        reservedNames
      );
      var newChildFolderId = generateId_();
      appendCatalogDataRow_(sheet, col, {
        id: newChildFolderId,
        parent_id: targetFolderId,
        folder_code: childCode,
        name: childUnique.name,
        type: 'folder',
        permissions_json: inherited.permissions_json,
        editors: inherited.editors,
        commenters: inherited.commenters,
        readers: inherited.readers
      });
      copyCatalogFolderChildren_(sheet, col, allNodes, child.id, newChildFolderId, childCode, targetParentId, reservedNames);
    } else {
      var fileFolderCode = targetFolderCode;
      var fileLabel = stripFolderPrefix_(child.name, child.folder_code);
      var fileUnique = allocateUniqueCatalogName_(
        targetFolderId,
        fileFolderCode,
        fileLabel,
        allNodes,
        {},
        reservedNames
      );
      var newFileId = '';
      if (child.file_id) {
        newFileId = copyDriveFile_(child.file_id, fileUnique.name);
      }
      appendCatalogDataRow_(sheet, col, {
        id: generateId_(),
        parent_id: targetFolderId,
        folder_code: fileFolderCode,
        name: fileUnique.name,
        type: 'file',
        file_id: newFileId,
        mime_type: child.mime_type || '',
        permissions_json: inherited.permissions_json,
        editors: inherited.editors,
        commenters: inherited.commenters,
        readers: inherited.readers
      });
    }
  }
}

/**
 * @param {string} fileId
 * @param {string} newName
 * @returns {string}
 */
function copyDriveFile_(fileId, newName) {
  var copied = Drive.Files.copy({ name: newName }, fileId, {
    supportsAllDrives: true,
    fields: 'id,name'
  });
  return copied.id;
}

/**
 * @param {string[]} ids
 * @param {Object.<string, Object>} nodesById
 * @returns {Array}
 */
function resolveCatalogEntriesByIds_(ids, nodesById) {
  var entries = [];
  for (var i = 0; i < ids.length; i++) {
    var entry = nodesById[ids[i]];
    if (entry && entry.id && entry.type) {
      entries.push(entry);
    }
  }
  return entries;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Array}
 */
function resolveSelectedCatalogEntries_(ss) {
  if (!ss.getActiveSheet() || ss.getActiveSheet().getName() !== 'Catalog') {
    throw new Error('Выберите строки на листе Catalog.');
  }

  var rowIndexes = getSelectedCatalogRowIndexes_(ss);
  if (!rowIndexes.length) {
    return [];
  }

  var sheets = getCatalogSheets_(ss);
  var indexSheet = sheets.index || getCatalogIndexSheet_(ss);
  var indexCol = getCatalogIndexColumnMap_();
  var dataCol = columnIndexMap_(SHEET_HEADERS.CatalogData);

  var entries = [];
  for (var i = 0; i < rowIndexes.length; i++) {
    var indexEntry = readCatalogIndexEntryAtViewRow_(indexSheet, indexCol, rowIndexes[i]);
    if (!indexEntry || !indexEntry.data_row_index) {
      continue;
    }
    var entry = readCatalogDataNodeAtRow_(sheets.data, dataCol, indexEntry.data_row_index);
    if (!entry || !entry.id || !entry.type) {
      continue;
    }
    entries.push(entry);
  }
  return entries;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {number[]}
 */
function getSelectedCatalogRowIndexes_(ss) {
  var rows = {};
  var rangeList = ss.getActiveRangeList();
  var ranges = rangeList ? rangeList.getRanges() : [];
  if (!ranges.length && ss.getActiveRange()) {
    ranges = [ss.getActiveRange()];
  }

  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    if (range.getSheet().getName() !== 'Catalog') {
      continue;
    }
    var startRow = range.getRow();
    var numRows = range.getNumRows();
    for (var offset = 0; offset < numRows; offset++) {
      var rowIndex = startRow + offset;
      if (rowIndex >= 2) {
        rows[rowIndex] = true;
      }
    }
  }

  return Object.keys(rows).map(function (key) {
    return parseInt(key, 10);
  }).sort(function (a, b) {
    return a - b;
  });
}

/**
 * @param {Object} node
 * @returns {boolean}
 */
function isProtectedCatalogNode_(node) {
  return node.folder_code === FOLDER_CODE_CATALOG || node.folder_code === FOLDER_CODE_TRASH;
}

/**
 * @param {string} fileId
 */
function deleteDriveFile_(fileId) {
  if (!fileId) {
    return;
  }
  Drive.Files.remove(fileId, { supportsAllDrives: true });
}

/**
 * @param {string} errText
 * @returns {boolean}
 */
function isDrivePermissionAlreadyExistsError_(errText) {
  var text = String(errText || '').toLowerCase();
  return text.indexOf('already exists') !== -1 ||
    text.indexOf('already has access') !== -1 ||
    text.indexOf('duplicate') !== -1;
}

/**
 * Drive не позволяет понизить прямое право ниже унаследованного с родительской папки.
 *
 * @param {string} errText
 * @returns {boolean}
 */
function isDriveInheritedAccessLimitError_(errText) {
  var text = String(errText || '').toLowerCase();
  return text.indexOf('inherited access') !== -1 ||
    text.indexOf('less than the inherited') !== -1 ||
    text.indexOf('limited-expansive-access') !== -1 ||
    text.indexOf('limited access') !== -1;
}

/**
 * @param {string} left
 * @param {string} right
 * @returns {number}
 */
function compareDriveRole_(left, right) {
  var order = { owner: 4, organizer: 4, fileorganizer: 4, writer: 3, commenter: 2, reader: 1 };
  return (order[String(left || '').toLowerCase()] || 0) - (order[String(right || '').toLowerCase()] || 0);
}

/**
 * Синхронизирует права файла на Drive с permissions_json (создание, обновление, удаление).
 * @param {string} fileId
 * @param {Array|string} permissionsInput
 */
function syncDrivePermissionsFromJson_(fileId, permissionsInput, ss) {
  if (!fileId) {
    return;
  }
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var permissions = [];
  if (typeof permissionsInput === 'string') {
    permissions = normalizeCatalogPermissions_(parseCatalogPermissionsJson_(permissionsInput));
  } else {
    permissions = normalizeCatalogPermissions_(permissionsInput || []);
  }
  permissions = expandCatalogGroupsInPermissions_(permissions, ss);

  var desiredMap = {};
  var syncErrors = [];
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (!perm.subject || perm.type === 'anyone' || perm.type === 'domain') {
      continue;
    }
    desiredMap[buildDrivePermissionKey_(perm.type, perm.subject)] = perm;
  }

  var current = listDrivePermissionsWithIds_(fileId);
  for (var c = 0; c < current.length; c++) {
    var currentPerm = current[c];
    var key = buildDrivePermissionKey_(currentPerm.type, currentPerm.subject);
    var desired = desiredMap[key];
    if (!desired) {
      if (currentPerm.inherited) {
        continue;
      }
      try {
        Drive.Permissions.remove(fileId, currentPerm.id, { supportsAllDrives: true });
      } catch (removeErr) {
        Logger.log('syncDrivePermissionsFromJson_ remove ' + currentPerm.id + ': ' + removeErr);
      }
      continue;
    }
    var desiredRole = mapLevelToDriveRole_(desired.level);
    if (currentPerm.inherited) {
      if (compareDriveRole_(currentPerm.role, desiredRole) >= 0) {
        delete desiredMap[key];
      }
      continue;
    }
    if (currentPerm.role !== desiredRole) {
      try {
        Drive.Permissions.update(
          { role: desiredRole },
          fileId,
          currentPerm.id,
          { supportsAllDrives: true, sendNotificationEmail: false }
        );
      } catch (updateErr) {
        var updateErrText = String(updateErr);
        if (isDriveInheritedAccessLimitError_(updateErrText) &&
            compareDriveRole_(currentPerm.role, desiredRole) > 0) {
          Logger.log('syncDrivePermissionsFromJson_ update skipped (inherited floor) ' +
            currentPerm.id + ': ' + updateErrText);
        } else {
          syncErrors.push(updateErrText);
          Logger.log('syncDrivePermissionsFromJson_ update ' + currentPerm.id + ': ' + updateErr);
        }
      }
    }
    delete desiredMap[key];
  }

  var remainingKeys = Object.keys(desiredMap);
  for (var r = 0; r < remainingKeys.length; r++) {
    var newPerm = desiredMap[remainingKeys[r]];
    var resource = {
      role: mapLevelToDriveRole_(newPerm.level),
      type: newPerm.type === 'group' ? 'group' : 'user',
      emailAddress: newPerm.subject
    };
    try {
      Drive.Permissions.create(resource, fileId, {
        supportsAllDrives: true,
        sendNotificationEmail: false
      });
    } catch (permErr) {
      var errText = String(permErr);
      if (isDrivePermissionAlreadyExistsError_(errText) ||
          isDriveInheritedAccessLimitError_(errText)) {
        Logger.log('syncDrivePermissionsFromJson_ create skipped: ' + errText);
        continue;
      }
      syncErrors.push(errText);
      Logger.log('syncDrivePermissionsFromJson_ create: ' + permErr);
    }
  }

  if (syncErrors.length) {
    throw new Error(syncErrors[0]);
  }

  try {
    enforceWritersCanShare_(fileId);
  } catch (shareErr) {
    Logger.log('writersCanShare: ' + shareErr);
  }
}

/**
 * @param {string} type
 * @param {string} subject
 * @returns {string}
 */
function buildDrivePermissionKey_(type, subject) {
  return String(type || 'user') + ':' + normalizeEmail_(subject);
}

/**
 * @param {string} fileId
 * @returns {Array<{id: string, subject: string, type: string, role: string, level: string}>}
 */
function listDrivePermissionsWithIds_(fileId) {
  var permissions = [];
  var pageToken = null;

  do {
    var response = Drive.Permissions.list(fileId, {
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      fields: 'nextPageToken, permissions(id,emailAddress,type,role,displayName,domain,permissionDetails(inherited))'
    });

    var batch = response.permissions || [];
    for (var i = 0; i < batch.length; i++) {
      var perm = batch[i];
      if (perm.role === 'owner') {
        continue;
      }
      var level = mapDriveRoleToLevel_(perm.role);
      if (!level) {
        continue;
      }

      var subject = '';
      var type = perm.type || 'user';
      if (type === 'user' || type === 'group') {
        subject = perm.emailAddress || '';
      } else if (type === 'domain') {
        subject = perm.domain || '';
      } else if (type === 'anyone') {
        subject = 'anyone';
      }
      if (!subject) {
        continue;
      }

      permissions.push({
        id: perm.id,
        subject: subject,
        type: type,
        role: perm.role,
        level: level,
        inherited: isDrivePermissionInherited_(perm)
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return permissions;
}

/**
 * @param {Object} perm
 * @returns {boolean}
 */
function isDrivePermissionInherited_(perm) {
  if (!perm || !perm.permissionDetails || !perm.permissionDetails.length) {
    return false;
  }
  for (var i = 0; i < perm.permissionDetails.length; i++) {
    if (perm.permissionDetails[i].inherited) {
      return true;
    }
  }
  return false;
}

/**
 * @deprecated Используйте syncDrivePermissionsFromJson_
 * @param {string} fileId
 * @param {string} permissionsJson
 */
function applyPermissionsFromJson_(fileId, permissionsJson) {
  syncDrivePermissionsFromJson_(fileId, permissionsJson);
}

/**
 * @param {string} level
 * @returns {string}
 */
function mapLevelToDriveRole_(level) {
  if (level === 'edit') {
    return 'writer';
  }
  if (level === 'comment') {
    return 'commenter';
  }
  return 'reader';
}

function firstScanCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return {
      ok: false,
      error: 'NO_SPREADSHEET',
      message: 'Нет активной таблицы.'
    };
  }

  var settingsSheet = ss.getSheetByName('Settings');
  var sheets = getCatalogSheets_(ss);
  var usersSheet = ss.getSheetByName('Users');
  if (!settingsSheet || !usersSheet) {
    return {
      ok: false,
      error: 'SCHEMA_MISSING',
      message: 'Сначала выполните setupSchema.'
    };
  }

  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    return {
      ok: false,
      error: 'NO_ROOT_FOLDER',
      message: 'Укажите root_folder_id на скрытом листе Settings (Данные → Скрытые листы).'
    };
  }

  var folderCheck = validateRootFolder_(rootFolderId);
  if (!folderCheck.ok) {
    return folderCheck;
  }

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    return headerCheck;
  }

  var col = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var catalogParentId = findIdByFolderCode_(sheets.data, col, FOLDER_CODE_CATALOG);
  if (!catalogParentId) {
    return {
      ok: false,
      error: 'NO_CATALOG_FOLDER',
      message: 'Не найдена системная папка «Каталог». Выполните setupSchema.'
    };
  }

  var knownFileIds = buildFileIdMap_(sheets.data, col);
  var skipFolders = {};
  try {
    var importFolderId = ensureCatalogImportDriveFolder_(ss);
    if (importFolderId) {
      skipFolders[importFolderId] = true;
    }
  } catch (importFolderErr) {
    Logger.log('firstScan import folder: ' + importFolderErr);
  }
  var driveFiles = listAllFilesInFolder_(rootFolderId, skipFolders);
  var myEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());

  var added = [];
  var skipped = [];
  var failed = [];
  var needsOwnerTransfer = [];
  var userEmails = {};

  for (var i = 0; i < driveFiles.length; i++) {
    var driveFile = driveFiles[i];
    if (driveFile.mimeType === DRIVE_SHORTCUT_MIME) {
      skipped.push(driveFile.id);
      continue;
    }
    if (knownFileIds[driveFile.id]) {
      skipped.push(driveFile.id);
      continue;
    }

    try {
      var permissions = [];
      try {
        permissions = readDrivePermissions_(driveFile.id);
      } catch (permErr) {
        Logger.log('permissions read failed for ' + driveFile.id + ': ' + permErr);
      }

      collectUserEmailsFromPermissions_(permissions, userEmails);
      var display = buildPermissionDisplay_(permissions, readUsersMap_(usersSheet));

      appendCatalogDataRow_(sheets.data, col, {
        id: generateId_(),
        file_id: driveFile.id,
        name: driveFile.name,
        type: 'file',
        folder_code: '',
        parent_id: catalogParentId,
        mime_type: driveFile.mimeType,
        permissions_json: JSON.stringify(permissions),
        editors: display.editors,
        commenters: display.commenters,
        readers: display.readers
      });

      try {
        enforceWritersCanShare_(driveFile.id);
      } catch (shareErr) {
        Logger.log('writersCanShare failed for ' + driveFile.id + ': ' + shareErr);
      }

      knownFileIds[driveFile.id] = true;
      added.push({ file_id: driveFile.id, name: driveFile.name });

      var ownerEmail = normalizeEmail_(driveFile.ownerEmail);
      if (ownerEmail && myEmail && ownerEmail !== myEmail) {
        needsOwnerTransfer.push({
          file_id: driveFile.id,
          name: driveFile.name,
          owner_email: ownerEmail
        });
      }
    } catch (fileErr) {
      failed.push({
        file_id: driveFile.id,
        name: driveFile.name,
        message: String(fileErr)
      });
    }
  }

  var usersAdded = seedUsersFromEmails_(usersSheet, userEmails, readUsersMap_(usersSheet));
  renderCatalogViewCore_(sheets.view, sheets.data);

  return {
    ok: failed.length === 0,
    data: {
      root_folder_id: rootFolderId,
      root_folder_name: folderCheck.data.name,
      files_found: driveFiles.length,
      files_added: added.length,
      files_skipped: skipped.length,
      files_failed: failed.length,
      users_added: usersAdded,
      added: added,
      failed: failed,
      needs_owner_transfer: needsOwnerTransfer
    },
    message: buildFirstScanMessage_(added.length, driveFiles.length, skipped.length, failed)
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{ok: boolean, error?: string, message?: string}}
 */
function validateCatalogDataHeaders_(sheet) {
  var expected = SHEET_HEADERS.CatalogData;
  var actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  for (var i = 0; i < expected.length; i++) {
    if (actual[i] !== expected[i]) {
      return {
        ok: false,
        error: 'CATALOG_DATA_HEADERS_MISMATCH',
        message: 'Заголовки CatalogData не совпадают со схемой. Выполните setupSchema.\n' +
          'Ожидалось: ' + expected[i] + ', найдено: ' + actual[i] + ' (колонка ' + (i + 1) + ').'
      };
    }
  }
  return { ok: true };
}

/**
 * @param {number} added
 * @param {number} found
 * @param {number} skipped
 * @param {Array<{name: string, message: string}>} failed
 * @returns {string}
 */
function buildFirstScanMessage_(added, found, skipped, failed) {
  var failedCount = failed ? failed.length : 0;
  var lines = [
    'Найдено файлов на Drive: ' + found,
    'Добавлено в каталог: ' + added,
    'Уже были в каталоге: ' + skipped
  ];
  if (failedCount > 0) {
    lines.push('Ошибок при добавлении: ' + failedCount);
    lines.push('');
    var limit = Math.min(failedCount, 3);
    for (var i = 0; i < limit; i++) {
      lines.push((failed[i].name || failed[i].file_id) + ':');
      lines.push(failed[i].message);
    }
    if (failedCount > limit) {
      lines.push('… и ещё ' + (failedCount - limit));
    }
  }
  if (found === 0) {
    lines.push('');
    lines.push('Папка пуста или нет доступа. Проверьте root_folder_id на листе Settings.');
  }
  return lines.join('\n');
}

/**
 * @param {string} folderId
 * @returns {{ok: boolean, data?: {name: string}, error?: string, message?: string}}
 */
function validateRootFolder_(folderId) {
  try {
    var folder = Drive.Files.get(folderId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,trashed'
    });
    if (folder.trashed) {
      return {
        ok: false,
        error: 'ROOT_FOLDER_TRASHED',
        message: 'Папка каталога на Drive (root_folder_id) находится в корзине.'
      };
    }
    if (folder.mimeType !== DRIVE_FOLDER_MIME) {
      return {
        ok: false,
        error: 'ROOT_NOT_FOLDER',
        message: 'root_folder_id указывает не на папку, а на файл.'
      };
    }
    return { ok: true, data: { name: folder.name } };
  } catch (err) {
    return {
      ok: false,
      error: 'ROOT_FOLDER_NOT_FOUND',
      message: formatRootFolderAccessError_(err)
    };
  }
}

/**
 * @param {*} err
 * @returns {string}
 */
function formatRootFolderAccessError_(err) {
  var details = String(err);
  var isAccessDenied = /not found|404|403|permission|insufficient|forbidden/i.test(details);
  if (isAccessDenied) {
    return 'Нет доступа к папке каталога на Google Drive (root_folder_id).\n\n' +
      'Скрипт выполняется от имени текущего пользователя («' +
      (Session.getEffectiveUser().getEmail() || 'неизвестно') +
      '»). Владелец каталога должен:\n' +
      '1. Открыть в Drive папку с ID из Settings → root_folder_id\n' +
      '2. Поделиться ею с этим пользователем с правом «Редактор»\n\n' +
      'Без этого копии файлов в каталог создать нельзя. ' +
      'Google Drive при отсутствии доступа часто отвечает «File not found».';
  }
  return 'Не удалось открыть папку каталога на Drive (root_folder_id). ' + details;
}

/**
 * @param {string} folderId
 * @param {Object.<string, boolean>=} skipFolderIds
 * @returns {Array<{id: string, name: string, mimeType: string, ownerEmail: string}>}
 */
function listAllFilesInFolder_(folderId, skipFolderIds) {
  var files = [];
  collectFilesRecursive_(folderId, files, skipFolderIds || {});
  return files;
}

/**
 * @param {string} folderId
 * @param {Array<{id: string, name: string, mimeType: string, ownerEmail: string}>} files
 * @param {Object.<string, boolean>=} skipFolderIds
 */
function collectFilesRecursive_(folderId, files, skipFolderIds) {
  var pageToken = null;
  var query = "'" + folderId + "' in parents and trashed = false";

  do {
    var response = Drive.Files.list({
      q: query,
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'nextPageToken, files(id,name,mimeType,size,writersCanShare,owners(emailAddress))'
    });

    var batch = response.files || [];
    for (var i = 0; i < batch.length; i++) {
      var item = batch[i];
      if (item.mimeType === DRIVE_FOLDER_MIME) {
        if (skipFolderIds && skipFolderIds[item.id]) {
          continue;
        }
        collectFilesRecursive_(item.id, files, skipFolderIds);
        continue;
      }
      if (item.mimeType === DRIVE_SHORTCUT_MIME) {
        continue;
      }
      var ownerEmail = '';
      if (item.owners && item.owners.length) {
        ownerEmail = item.owners[0].emailAddress || '';
      }
      files.push({
        id: item.id,
        name: item.name,
        mimeType: item.mimeType,
        ownerEmail: ownerEmail,
        size: parseInt(item.size, 10) || 0,
        writersCanShare: item.writersCanShare === true
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
}

/**
 * @param {{ok: boolean, message?: string, error?: string, data?: object}} result
 */
function showFirstScanResult_(result) {
  try {
    var ui = SpreadsheetApp.getUi();
    if (!ui) {
      return;
    }
    var title = result.ok ? 'Первое сканирование' : 'Ошибка сканирования';
    var text = result.message || result.error || 'Неизвестная ошибка';
    ui.alert(title, text, ui.ButtonSet.OK);
  } catch (err) {
    Logger.log('firstScan result: ' + JSON.stringify(result));
  }
}

var INTEGRITY_REPORT_SHEET_NAME = 'Отчёт целостности';
var INTEGRITY_REPORT_CACHE_KEY = 'catalog_integrity_report_v1';

/**
 * Проверка целостности каталога: соответствие Drive ↔ CatalogData, права, владельцы, структура.
 * Запуск: меню «Каталогизатор → Проверка целостности».
 *
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function runCatalogIntegrityCheck() {
  try {
    var result = runCatalogIntegrityCheckCore_();
    showIntegrityCheckResult_(result);
    return result;
  } catch (err) {
    var fail = {
      ok: false,
      error: 'INTEGRITY_CHECK_FAILED',
      message: String(err)
    };
    showIntegrityCheckResult_(fail);
    return fail;
  }
}

/**
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function runCatalogIntegrityCheckCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return {
      ok: false,
      error: 'NO_SPREADSHEET',
      message: 'Нет активной таблицы.'
    };
  }

  var settingsSheet = ss.getSheetByName('Settings');
  var sheets = getCatalogSheets_(ss);
  var usersSheet = ss.getSheetByName('Users');
  if (!settingsSheet || !usersSheet) {
    return {
      ok: false,
      error: 'SCHEMA_MISSING',
      message: 'Сначала выполните setupSchema.'
    };
  }

  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    return {
      ok: false,
      error: 'NO_ROOT_FOLDER',
      message: 'Укажите root_folder_id на скрытом листе Settings (Данные → Скрытые листы).'
    };
  }

  var folderCheck = validateRootFolder_(rootFolderId);
  if (!folderCheck.ok) {
    return folderCheck;
  }

  var headerCheck = validateCatalogDataHeaders_(sheets.data);
  if (!headerCheck.ok) {
    return headerCheck;
  }

  var col = columnIndexMap_(SHEET_HEADERS.CatalogData);
  var tree = buildCatalogTreeState_(sheets.data, col);
  if (!tree) {
    return {
      ok: false,
      error: 'EMPTY_CATALOG',
      message: 'CatalogData пуст. Выполните setupSchema или первое сканирование.'
    };
  }

  var catalogOwnerEmail = getSpreadsheetOwnerEmail_(ss);
  var usersMap = readUsersMap_(usersSheet);
  var importFolderId = '';
  try {
    importFolderId = ensureCatalogImportDriveFolder_(ss) || '';
  } catch (importErr) {
    Logger.log('integrity import folder: ' + importErr);
  }

  var skipFolders = {};
  if (importFolderId) {
    skipFolders[importFolderId] = true;
  }

  var driveFilesInRoot = listAllFilesInFolder_(rootFolderId, skipFolders);
  var driveById = {};
  var driveMetaById = {};
  for (var d = 0; d < driveFilesInRoot.length; d++) {
    driveById[driveFilesInRoot[d].id] = true;
    driveMetaById[driveFilesInRoot[d].id] = driveFilesInRoot[d];
  }

  var importDriveFiles = importFolderId ? listFilesInSingleFolder_(importFolderId) : [];
  var importById = {};
  for (var im = 0; im < importDriveFiles.length; im++) {
    importById[importDriveFiles[im].id] = importDriveFiles[im];
    driveMetaById[importDriveFiles[im].id] = importDriveFiles[im];
  }

  var trashId = getTrashFolderId_(sheets.data, col);
  var structureIssues = collectCatalogStructureIssues_(tree);
  var fileRows = [];
  var issues = structureIssues.slice();
  issues = issues.concat(collectIntegrityTrashIssues_(tree, trashId, driveMetaById));
  issues = issues.concat(collectIntegrityStalePermissionIssues_(tree, ss));
  issues = issues.concat(collectIntegrityPhysicalFolderIssues_(rootFolderId, importFolderId));
  issues = issues.concat(collectIntegrityRootAclIssues_(rootFolderId, catalogOwnerEmail));

  var catalogFileIdToNode = {};
  var catalogFileNodes = [];
  for (var n = 0; n < tree.sorted.length; n++) {
    var node = tree.sorted[n];
    if (node.type !== 'file') {
      continue;
    }
    catalogFileNodes.push(node);
    if (node.file_id) {
      if (catalogFileIdToNode[node.file_id]) {
        issues.push({
          severity: 'error',
          category: 'duplicate_file_id',
          item: node.name || node.file_id,
          details: 'file_id ' + node.file_id + ' дублируется (строки ' +
            catalogFileIdToNode[node.file_id].rowIndex + ' и ' + node.rowIndex + ')'
        });
      } else {
        catalogFileIdToNode[node.file_id] = node;
      }
    }
  }

  var permissionCache = {};
  var approvedEditorsCache = {};
  var totalCatalogBytes = 0;

  for (var f = 0; f < catalogFileNodes.length; f++) {
    var fileNode = catalogFileNodes[f];
    var path = buildCatalogVirtualPath_(fileNode.id, tree.nodesById);
    var rowIssues = [];
    var driveMeta = null;
    var driveLocation = 'missing';

    if (!fileNode.file_id) {
      rowIssues.push('нет file_id');
      issues.push({
        severity: 'error',
        category: 'missing_file_id',
        item: path || fileNode.name,
        details: 'Строка ' + fileNode.rowIndex + ': файл в каталоге без file_id'
      });
    } else if (driveById[fileNode.file_id]) {
      driveLocation = 'root';
      for (var dr = 0; dr < driveFilesInRoot.length; dr++) {
        if (driveFilesInRoot[dr].id === fileNode.file_id) {
          driveMeta = driveFilesInRoot[dr];
          break;
        }
      }
    } else if (importById[fileNode.file_id]) {
      driveLocation = 'import';
      driveMeta = importById[fileNode.file_id];
      rowIssues.push('файл в _Import, не в каталоге на Drive');
      issues.push({
        severity: 'warning',
        category: 'stuck_in_import',
        item: path || fileNode.name,
        details: 'file_id ' + fileNode.file_id + ' всё ещё в папке _Import'
      });
    } else {
      rowIssues.push('нет на Drive');
      issues.push({
        severity: 'error',
        category: 'catalog_orphan',
        item: path || fileNode.name,
        details: 'file_id ' + fileNode.file_id + ' не найден в root_folder_id и _Import'
      });
    }

    var ownerEmail = '';
    var fileSize = 0;
    if (driveMeta) {
      ownerEmail = normalizeEmail_(driveMeta.ownerEmail);
      fileSize = driveMeta.size || 0;
      totalCatalogBytes += fileSize;
      if (catalogOwnerEmail && ownerEmail && ownerEmail !== catalogOwnerEmail) {
        rowIssues.push('владелец: ' + ownerEmail);
        issues.push({
          severity: 'error',
          category: 'wrong_owner',
          item: path || fileNode.name,
          details: 'Владелец ' + ownerEmail + ', ожидается ' + catalogOwnerEmail
        });
      }
      if (driveMeta.writersCanShare === true) {
        rowIssues.push('writersCanShare');
        issues.push({
          severity: 'warning',
          category: 'writers_can_share',
          item: path || fileNode.name,
          details: 'writersCanShare=true на Drive (ожидается false)'
        });
      }
      var catalogDriveName = readCell_(fileNode.name);
      if (catalogDriveName && driveMeta.name && driveMeta.name !== catalogDriveName) {
        rowIssues.push('имя на Drive');
        issues.push({
          severity: 'warning',
          category: 'filename_mismatch',
          item: path || fileNode.name,
          details: 'Каталог: «' + catalogDriveName + '», Drive: «' + driveMeta.name + '»'
        });
      }
      if (fileNode.mime_type && driveMeta.mimeType && fileNode.mime_type !== driveMeta.mimeType) {
        rowIssues.push('mime_type');
        issues.push({
          severity: 'warning',
          category: 'mime_type_drift',
          item: path || fileNode.name,
          details: 'Каталог: ' + fileNode.mime_type + ', Drive: ' + driveMeta.mimeType
        });
      }
    } else if (fileNode.file_id) {
      ownerEmail = getDriveFileOwnerEmail_(fileNode.file_id);
      if (catalogOwnerEmail && ownerEmail && ownerEmail !== catalogOwnerEmail) {
        rowIssues.push('владелец: ' + ownerEmail);
        issues.push({
          severity: 'warning',
          category: 'wrong_owner',
          item: path || fileNode.name,
          details: 'Владелец ' + ownerEmail + ', ожидается ' + catalogOwnerEmail
        });
      }
    }

    var permsMatch = '—';
    if (fileNode.file_id && driveLocation !== 'missing') {
      var permResult = compareCatalogDrivePermissions_(fileNode, ss, permissionCache);
      permsMatch = permResult.match ? 'да' : 'нет';
      if (!permResult.match) {
        rowIssues.push('права: ' + permResult.summary);
        issues.push({
          severity: 'warning',
          category: 'permissions_mismatch',
          item: path || fileNode.name,
          details: permResult.summary
        });
      }
    }

    var display = permissionsJsonToDisplay_(fileNode.permissions_json, usersMap);
    fileRows.push({
      path: path,
      name: fileNode.name,
      catalog_id: fileNode.id,
      file_id: fileNode.file_id || '',
      drive_location: driveLocation,
      size: fileSize,
      owner: ownerEmail || '—',
      editors: display.editors,
      commenters: display.commenters,
      readers: display.readers,
      permissions_match: permsMatch,
      issues: rowIssues.join('; ')
    });
  }

  issues = issues.concat(
    collectIntegrityApprovedEditorIssues_(catalogFileNodes, tree, approvedEditorsCache)
  );

  for (var di = 0; di < driveFilesInRoot.length; di++) {
    var driveFile = driveFilesInRoot[di];
    if (!catalogFileIdToNode[driveFile.id]) {
      issues.push({
        severity: 'warning',
        category: 'drive_orphan',
        item: driveFile.name,
        details: 'file_id ' + driveFile.id + ' на Drive, но нет в CatalogData',
        repair_context: { file_id: driveFile.id, name: driveFile.name }
      });
      if (driveFile.writersCanShare) {
        issues.push({
          severity: 'warning',
          category: 'writers_can_share',
          item: driveFile.name,
          details: 'file_id ' + driveFile.id + ': writersCanShare должен быть false'
        });
      }
    }
  }

  for (var ii = 0; ii < importDriveFiles.length; ii++) {
    var importFile = importDriveFiles[ii];
    if (!catalogFileIdToNode[importFile.id]) {
      issues.push({
        severity: 'info',
        category: 'import_orphan',
        item: importFile.name,
        details: 'file_id ' + importFile.id + ' в _Import без записи в каталоге',
        repair_context: { file_id: importFile.id, name: importFile.name, in_import: true }
      });
      if (importFile.writersCanShare) {
        issues.push({
          severity: 'warning',
          category: 'writers_can_share',
          item: importFile.name,
          details: 'file_id ' + importFile.id + ': writersCanShare должен быть false'
        });
      }
    }
  }

  var folderStats = buildIntegrityFolderStatistics_(tree, fileRows);
  var issueCounts = countIntegrityIssuesByCategory_(issues);
  var errorCount = issueCounts.error || 0;
  var warningCount = issueCounts.warning || 0;

  var reportData = {
    generated_at: new Date(),
    root_folder_id: rootFolderId,
    root_folder_name: folderCheck.data.name,
    catalog_owner: catalogOwnerEmail,
    catalog_files: catalogFileNodes.length,
    drive_files_in_root: driveFilesInRoot.length,
    drive_files_in_import: importDriveFiles.length,
    total_size_bytes: totalCatalogBytes,
    error_count: errorCount,
    warning_count: warningCount,
    info_count: issueCounts.info || 0,
    issues: issues,
    files: fileRows,
    folder_stats: folderStats
  };

  var sheetName = writeIntegrityReportSheet_(ss, reportData);
  reportData.report_sheet = sheetName;
  saveIntegrityReportSnapshot_(reportData);

  return {
    ok: errorCount === 0,
    data: reportData,
    message: buildIntegrityCheckMessage_(reportData)
  };
}

/**
 * @param {{ok: boolean, message?: string, error?: string, data?: object}} result
 */
function showIntegrityCheckResult_(result) {
  try {
    var ui = SpreadsheetApp.getUi();
    if (!ui) {
      return;
    }
    var title = result.ok ? 'Проверка целостности' : 'Проблемы целостности';
    var text = result.message || result.error || 'Неизвестная ошибка';
    ui.alert(title, text, ui.ButtonSet.OK);
  } catch (err) {
    Logger.log('integrity check: ' + JSON.stringify(result));
  }
}

/**
 * Мастер исправления по последнему отчёту целостности (только владелец таблицы).
 */
function runCatalogIntegrityRepairMenu_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }
  if (!isSpreadsheetOwner_(ss)) {
    SpreadsheetApp.getUi().alert(
      'Исправление',
      'Исправлять проблемы может только владелец таблицы.\n\nПроверку целостности может запустить любой пользователь.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var reportData = loadIntegrityReportSnapshot_();
  if (!reportData || !reportData.issues || !reportData.issues.length) {
    var rerun = SpreadsheetApp.getUi().alert(
      'Исправление',
      'Нет сохранённого отчёта целостности.\n\nЗапустить проверку сейчас?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (rerun !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    var checkResult = runCatalogIntegrityCheckCore_();
    if (!checkResult.data) {
      SpreadsheetApp.getUi().alert(
        'Исправление',
        checkResult.message || checkResult.error || 'Проверка не удалась.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    reportData = checkResult.data;
  }

  var repairResult = runCatalogIntegrityRepairWizard_(ss, reportData);
  var summary = [
    'Исправлено: ' + repairResult.fixed,
    'Пропущено: ' + repairResult.skipped,
    'Ошибок: ' + repairResult.failed
  ];
  if (repairResult.stopped) {
    summary.push('Мастер остановлен досрочно.');
  }
  if (repairResult.fixed > 0) {
    summary.push('');
    summary.push('Рекомендуется повторить проверку целостности.');
  }
  SpreadsheetApp.getUi().alert('Исправление завершено', summary.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * @param {object} reportData
 */
function saveIntegrityReportSnapshot_(reportData) {
  try {
    var cache = CacheService.getDocumentCache();
    if (!cache) {
      return;
    }
    var snapshot = {
      generated_at: reportData.generated_at ? String(reportData.generated_at) : '',
      root_folder_id: reportData.root_folder_id || '',
      catalog_owner: reportData.catalog_owner || '',
      issues: reportData.issues || [],
      files: reportData.files || []
    };
    var payload = JSON.stringify(snapshot);
    if (payload.length > 95000) {
      snapshot.files = [];
      payload = JSON.stringify(snapshot);
    }
    cache.put(INTEGRITY_REPORT_CACHE_KEY, payload, 21600);
  } catch (err) {
    Logger.log('saveIntegrityReportSnapshot_: ' + err);
  }
}

/**
 * @returns {object|null}
 */
function loadIntegrityReportSnapshot_() {
  try {
    var cache = CacheService.getDocumentCache();
    if (!cache) {
      return null;
    }
    var raw = cache.get(INTEGRITY_REPORT_CACHE_KEY);
    if (!raw) {
      return null;
    }
    return JSON.parse(raw);
  } catch (err) {
    Logger.log('loadIntegrityReportSnapshot_: ' + err);
    return null;
  }
}

/**
 * @param {string} category
 * @returns {boolean}
 */
function isIntegrityIssueRepairable_(category) {
  return !!getIntegrityRepairActionText_(category);
}

/**
 * @param {string} category
 * @returns {string}
 */
function getIntegrityRepairActionText_(category) {
  var map = {
    root_acl_public: 'Снять публичный/доменный доступ с корневой папки Drive',
    root_acl_broad: 'Снять весь прямой доступ с корня (кроме владельца каталога)',
    physical_drive_folder: 'Удалить пустую физическую папку на Drive',
    drive_orphan: 'Удалить файл с Drive (нет в каталоге)',
    import_orphan: 'Удалить файл из _Import (нет в каталоге)'
  };
  return map[category] || '';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {object} reportData
 * @returns {{fixed: number, skipped: number, failed: number, stopped: boolean}}
 */
function runCatalogIntegrityRepairWizard_(ss, reportData) {
  var ui = SpreadsheetApp.getUi();
  var issues = reportData.issues || [];
  var repairable = [];
  for (var i = 0; i < issues.length; i++) {
    if (isIntegrityIssueRepairable_(issues[i].category)) {
      repairable.push(issues[i]);
    }
  }

  if (!repairable.length) {
    ui.alert(
      'Исправление',
      'В отчёте нет проблем, которые можно исправить автоматически (уровень «владелец + подтверждение»).',
      ui.ButtonSet.OK
    );
    return { fixed: 0, skipped: 0, failed: 0, stopped: false };
  }

  var intro = ui.alert(
    'Исправление',
    'Найдено проблем в отчёте: ' + issues.length + '.\n' +
      'Можно исправить автоматически: ' + repairable.length + '.\n\n' +
      'По каждой — вопрос: исправить, пропустить или остановить мастер.\n\nНачать?',
    ui.ButtonSet.YES_NO
  );
  if (intro !== ui.Button.YES) {
    return { fixed: 0, skipped: 0, failed: 0, stopped: false };
  }

  var stats = { fixed: 0, skipped: 0, failed: 0, stopped: false };
  var ownerEmail = getSpreadsheetOwnerEmail_(ss);

  for (var r = 0; r < repairable.length; r++) {
    var issue = repairable[r];
    var actionText = getIntegrityRepairActionText_(issue.category);
    var prompt = [
      'Проблема ' + (r + 1) + ' из ' + repairable.length,
      '',
      'Уровень: ' + issue.severity,
      'Категория: ' + issue.category,
      'Объект: ' + issue.item,
      '',
      issue.details,
      '',
      'Действие: ' + actionText,
      '',
      'Исправить?'
    ].join('\n');

    var response = ui.alert('Исправление', prompt, ui.ButtonSet.YES_NO_CANCEL);
    if (response === ui.Button.CANCEL) {
      stats.stopped = true;
      break;
    }
    if (response === ui.Button.NO) {
      stats.skipped++;
      continue;
    }

    var applyResult = applyIntegrityRepair_(issue, reportData, ss, ownerEmail);
    if (applyResult.ok) {
      stats.fixed++;
    } else {
      stats.failed++;
      ui.alert(
        'Ошибка исправления',
        (issue.item || issue.category) + ':\n' + (applyResult.message || 'неизвестная ошибка'),
        ui.ButtonSet.OK
      );
    }
  }

  return stats;
}

/**
 * @param {object} issue
 * @param {object} reportData
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} ownerEmail
 * @returns {{ok: boolean, message?: string}}
 */
function applyIntegrityRepair_(issue, reportData, ss, ownerEmail) {
  var ctx = issue.repair_context || {};
  var category = issue.category;

  try {
    if (category === 'root_acl_public') {
      return repairIntegrityRemoveRootPermission_(ctx.root_folder_id || reportData.root_folder_id, ctx.permission_id);
    }
    if (category === 'root_acl_broad') {
      return repairIntegrityRemoveRootBroadAcl_(
        ctx.root_folder_id || reportData.root_folder_id,
        ownerEmail || reportData.catalog_owner
      );
    }
    if (category === 'physical_drive_folder') {
      return repairIntegrityDeleteEmptyDriveFolder_(ctx.folder_id, ctx.folder_name || issue.item);
    }
    if (category === 'drive_orphan' || category === 'import_orphan') {
      return repairIntegrityDeleteDriveFile_(ctx.file_id || parseIntegrityIssueFileId_(issue.details), issue.item);
    }
    return { ok: false, message: 'Нет обработчика для категории ' + category };
  } catch (err) {
    return { ok: false, message: String(err) };
  }
}

/**
 * @param {string} details
 * @returns {string}
 */
function parseIntegrityIssueFileId_(details) {
  var match = String(details || '').match(/file_id ([A-Za-z0-9_-]+)/i);
  return match ? match[1] : '';
}

/**
 * @param {string} rootFolderId
 * @param {string} permissionId
 * @returns {{ok: boolean, message?: string}}
 */
function repairIntegrityRemoveRootPermission_(rootFolderId, permissionId) {
  if (!rootFolderId || !permissionId) {
    return { ok: false, message: 'Нет root_folder_id или permission_id' };
  }
  Drive.Permissions.remove(rootFolderId, permissionId, { supportsAllDrives: true });
  return { ok: true };
}

/**
 * @param {string} rootFolderId
 * @param {string} catalogOwnerEmail
 * @returns {{ok: boolean, message?: string}}
 */
function repairIntegrityRemoveRootBroadAcl_(rootFolderId, catalogOwnerEmail) {
  if (!rootFolderId) {
    return { ok: false, message: 'Нет root_folder_id' };
  }
  var ownerEmail = normalizeEmail_(catalogOwnerEmail);
  var perms = listDrivePermissionsWithIds_(rootFolderId);
  var removed = 0;
  var errors = [];
  for (var i = 0; i < perms.length; i++) {
    var perm = perms[i];
    if (perm.inherited) {
      continue;
    }
    if (perm.type === 'user' && normalizeEmail_(perm.subject) === ownerEmail) {
      continue;
    }
    try {
      Drive.Permissions.remove(rootFolderId, perm.id, { supportsAllDrives: true });
      removed++;
    } catch (err) {
      errors.push(String(err));
    }
  }
  if (errors.length && !removed) {
    return { ok: false, message: errors[0] };
  }
  return { ok: true, message: 'Снято прав: ' + removed };
}

/**
 * @param {string} folderId
 * @param {string} folderName
 * @returns {{ok: boolean, message?: string}}
 */
function repairIntegrityDeleteEmptyDriveFolder_(folderId, folderName) {
  if (!folderId) {
    return { ok: false, message: 'Нет folder_id' };
  }
  var files = listFilesInSingleFolder_(folderId);
  var subfolders = listDirectDriveSubfolders_(folderId);
  if (files.length || subfolders.length) {
    return {
      ok: false,
      message: 'Папка «' + (folderName || folderId) + '» не пуста (' +
        files.length + ' файлов, ' + subfolders.length + ' подпапок). Удаление отменено.'
    };
  }
  Drive.Files.remove(folderId, { supportsAllDrives: true });
  return { ok: true };
}

/**
 * @param {string} fileId
 * @param {string} label
 * @returns {{ok: boolean, message?: string}}
 */
function repairIntegrityDeleteDriveFile_(fileId, label) {
  if (!fileId) {
    return { ok: false, message: 'Не удалось определить file_id' };
  }
  Drive.Files.remove(fileId, { supportsAllDrives: true });
  return { ok: true, message: 'Удалён: ' + (label || fileId) };
}

/**
 * @param {object} reportData
 * @returns {string}
 */
function buildIntegrityCheckMessage_(reportData) {
  var lines = [
    'Файлов в каталоге: ' + reportData.catalog_files,
    'Файлов на Drive (корень): ' + reportData.drive_files_in_root,
    'В _Import: ' + reportData.drive_files_in_import,
    'Общий размер (файлы в каталоге на Drive): ' + formatFileSize_(reportData.total_size_bytes),
    '',
    'Ошибок: ' + reportData.error_count,
    'Предупреждений: ' + reportData.warning_count
  ];
  if (reportData.info_count) {
    lines.push('Информационных: ' + reportData.info_count);
  }
  lines.push('');
  lines.push('Отчёт сохранён на листе «' + reportData.report_sheet + '».');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && isSpreadsheetOwner_(ss)) {
      lines.push('Владелец: меню «🔧 Исправление по отчёту» — пошаговое исправление.');
    }
  } catch (ownerHintErr) {
    Logger.log('integrity owner hint: ' + ownerHintErr);
  }
  return lines.join('\n');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {object} reportData
 * @returns {string}
 */
function writeIntegrityReportSheet_(ss, reportData) {
  var sheetName = INTEGRITY_REPORT_SHEET_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  var rows = [];
  var stamp = Utilities.formatDate(
    reportData.generated_at,
    Session.getScriptTimeZone() || 'Europe/Moscow',
    'yyyy-MM-dd HH:mm:ss'
  );

  rows.push(['Проверка целостности каталога', stamp]);
  rows.push([]);
  rows.push(['Сводка']);
  rows.push(['Корневая папка Drive', reportData.root_folder_name + ' (' + reportData.root_folder_id + ')']);
  rows.push(['Владелец каталога', reportData.catalog_owner || '—']);
  rows.push(['Файлов в каталоге', reportData.catalog_files]);
  rows.push(['Файлов на Drive (корень)', reportData.drive_files_in_root]);
  rows.push(['Файлов в _Import', reportData.drive_files_in_import]);
  rows.push(['Общий размер', formatFileSize_(reportData.total_size_bytes)]);
  rows.push(['Ошибок', reportData.error_count]);
  rows.push(['Предупреждений', reportData.warning_count]);
  rows.push(['Информационных', reportData.info_count || 0]);
  rows.push([]);

  rows.push(['Проблемы']);
  rows.push(['Уровень', 'Категория', 'Объект', 'Описание']);
  var issues = reportData.issues || [];
  if (!issues.length) {
    rows.push(['—', '—', '—', 'Проблем не обнаружено']);
  } else {
    for (var i = 0; i < issues.length; i++) {
      var issue = issues[i];
      rows.push([issue.severity, issue.category, issue.item, issue.details]);
    }
  }
  rows.push([]);

  rows.push(['Файлы каталога']);
  rows.push([
    'Путь', 'Имя', 'catalog_id', 'file_id', 'На Drive', 'Размер', 'Владелец',
    'Редактор', 'Комментатор', 'Чтение', 'Права совпадают', 'Замечания'
  ]);
  var files = reportData.files || [];
  for (var f = 0; f < files.length; f++) {
    var file = files[f];
    rows.push([
      file.path,
      file.name,
      file.catalog_id,
      file.file_id,
      file.drive_location,
      formatFileSize_(file.size),
      file.owner,
      file.editors,
      file.commenters,
      file.readers,
      file.permissions_match,
      file.issues
    ]);
  }
  rows.push([]);

  rows.push(['Статистика по папкам']);
  rows.push(['Путь', 'Код папки', 'Файлов', 'Размер']);
  var folderStats = reportData.folder_stats || [];
  for (var s = 0; s < folderStats.length; s++) {
    var stat = folderStats[s];
    rows.push([stat.path, stat.folder_code, stat.file_count, formatFileSize_(stat.total_size)]);
  }
  rows.push([]);
  rows.push(['Итого по каталогу', '', reportData.catalog_files, formatFileSize_(reportData.total_size_bytes)]);

  if (rows.length) {
    var maxWidth = 0;
    for (var w = 0; w < rows.length; w++) {
      if (rows[w].length > maxWidth) {
        maxWidth = rows[w].length;
      }
    }
    sheet.getRange(1, 1, rows.length, maxWidth).setValues(
      padIntegrityReportRows_(rows, maxWidth)
    );
  }
  sheet.setFrozenRows(1);
  try {
    sheet.autoResizeColumns(1, 12);
  } catch (resizeErr) {
    Logger.log('integrity report resize: ' + resizeErr);
  }
  return sheetName;
}

/**
 * @param {Array<Array>} rows
 * @param {number} width
 * @returns {Array<Array>}
 */
function padIntegrityReportRows_(rows, width) {
  var padded = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i].slice();
    while (row.length < width) {
      row.push('');
    }
    padded.push(row);
  }
  return padded;
}

/**
 * @param {object} tree
 * @param {Object.<string, number>} col
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectCatalogStructureIssues_(tree) {
  var issues = [];
  var nodes = tree.sorted;
  var nodesById = tree.nodesById;
  var idCounts = {};
  var reachable = {};
  var visited = {};

  function markReachable(id) {
    if (!id || visited[id] || !nodesById[id]) {
      return;
    }
    visited[id] = true;
    reachable[id] = true;
    var children = tree.childrenByParent[id] || [];
    for (var c = 0; c < children.length; c++) {
      markReachable(children[c]);
    }
  }
  markReachable(tree.catalogRootId);

  var codesByParent = {};
  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    if (!node.id) {
      issues.push({
        severity: 'error',
        category: 'missing_id',
        item: 'строка ' + node.rowIndex,
        details: 'Пустой id в CatalogData'
      });
      continue;
    }
    idCounts[node.id] = (idCounts[node.id] || 0) + 1;

    if (node.parent_id && !nodesById[node.parent_id]) {
      issues.push({
        severity: 'error',
        category: 'missing_parent',
        item: node.name || node.id,
        details: 'parent_id ' + node.parent_id + ' не найден (строка ' + node.rowIndex + ')'
      });
    }

    if (!reachable[node.id] && node.id !== tree.catalogRootId) {
      issues.push({
        severity: 'error',
        category: 'unreachable_node',
        item: node.name || node.id,
        details: 'Узел не достижим от корня «Каталог» (строка ' + node.rowIndex + ')'
      });
    }

    if (node.type === 'folder' && node.file_id) {
      issues.push({
        severity: 'warning',
        category: 'folder_has_file_id',
        item: node.name || node.folder_code,
        details: 'Папка не должна иметь file_id (строка ' + node.rowIndex + ')'
      });
    }

    if (node.type === 'file' && node.folder_code && !node.trash_parent_id) {
      var code = String(node.folder_code);
      if (!/^\d{2,}$/.test(code) && code !== '') {
        issues.push({
          severity: 'warning',
          category: 'invalid_folder_code',
          item: node.name || node.id,
          details: 'Некорректный folder_code «' + code + '» (строка ' + node.rowIndex + ')'
        });
      }
    }

    if (node.type === 'folder' && node.folder_code) {
      var parentKey = node.parent_id || '';
      if (!codesByParent[parentKey]) {
        codesByParent[parentKey] = {};
      }
      var folderCode = String(node.folder_code);
      if (codesByParent[parentKey][folderCode]) {
        issues.push({
          severity: 'error',
          category: 'duplicate_folder_code',
          item: node.name || folderCode,
          details: 'Дублирующийся folder_code «' + folderCode + '» у одного родителя'
        });
      } else {
        codesByParent[parentKey][folderCode] = true;
      }
    }

    if (node.trash_parent_id && !nodesById[node.trash_parent_id]) {
      issues.push({
        severity: 'error',
        category: 'invalid_trash_parent',
        item: node.name || node.id,
        details: 'trash_parent_id ' + node.trash_parent_id + ' не найден (строка ' + node.rowIndex + ')'
      });
    }
  }

  var dupIds = Object.keys(idCounts).filter(function (id) {
    return idCounts[id] > 1;
  });
  for (var d = 0; d < dupIds.length; d++) {
    issues.push({
      severity: 'error',
      category: 'duplicate_id',
      item: dupIds[d],
      details: 'id встречается ' + idCounts[dupIds[d]] + ' раз(а)'
    });
  }

  return issues;
}

/**
 * @param {object} tree
 * @param {string} trashId
 * @param {Object.<string, {id: string, name: string}>} driveMetaById
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectIntegrityTrashIssues_(tree, trashId, driveMetaById) {
  var issues = [];
  if (!tree || !tree.sorted) {
    return issues;
  }
  var nodesById = tree.nodesById;
  var nodes = tree.sorted;

  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    var trashParentId = readCell_(node.trash_parent_id);
    if (!trashParentId) {
      continue;
    }
    if (!trashId || node.parent_id !== trashId) {
      issues.push({
        severity: 'error',
        category: 'trash_inconsistent',
        item: node.name || node.id,
        details: 'trash_parent_id задан, но parent_id не указывает на «Корзину» (строка ' + node.rowIndex + ')'
      });
    }
    if (node.type === 'file' && node.file_id && !driveMetaById[node.file_id]) {
      issues.push({
        severity: 'warning',
        category: 'trash_drive_missing',
        item: node.name || node.id,
        details: 'Файл в корзине каталога, но file_id ' + node.file_id + ' отсутствует на Drive'
      });
    }
  }

  if (trashId) {
    for (var j = 0; j < nodes.length; j++) {
      var trashNode = nodes[j];
      if (trashNode.id === trashId || isProtectedCatalogNode_(trashNode)) {
        continue;
      }
      if (isNodeInTrashSubtree_(trashNode, trashId, nodesById) && !readCell_(trashNode.trash_parent_id)) {
        issues.push({
          severity: 'warning',
          category: 'trash_missing_restore_parent',
          item: trashNode.name || trashNode.id,
          details: 'Элемент в «Корзине», но trash_parent_id пуст (строка ' + trashNode.rowIndex + ')'
        });
      }
    }
  }

  return issues;
}

/**
 * @param {object} tree
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectIntegrityStalePermissionIssues_(tree, ss) {
  var issues = [];
  if (!tree || !tree.sorted) {
    return issues;
  }
  var usersSheet = ss.getSheetByName('Users');
  var groupsSheet = ss.getSheetByName('Groups');
  var activeUsers = readActiveUsersMap_(usersSheet);
  var allUsers = readUsersMap_(usersSheet);
  var groupsMap = readCatalogGroupsMap_(groupsSheet);
  var nodes = tree.sorted;

  for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];
    var perms = parseCatalogPermissionsJson_(node.permissions_json);
    if (!perms.length) {
      continue;
    }
    var itemLabel = node.name || node.id;
    for (var p = 0; p < perms.length; p++) {
      var perm = perms[p];
      if (!perm || !perm.subject) {
        continue;
      }
      if (perm.type === 'user') {
        var email = normalizeEmail_(perm.subject);
        if (!email) {
          continue;
        }
        if (!allUsers[email]) {
          issues.push({
            severity: 'warning',
            category: 'stale_permission_user',
            item: itemLabel,
            details: 'Пользователь ' + email + ' не найден в Users (строка ' + node.rowIndex + ')'
          });
        } else if (!activeUsers[email]) {
          issues.push({
            severity: 'warning',
            category: 'stale_permission_inactive_user',
            item: itemLabel,
            details: 'Неактивный пользователь ' + email + ' в permissions_json (строка ' + node.rowIndex + ')'
          });
        }
      } else if (perm.type === 'group') {
        var groupId = normalizeGroupId_(perm.subject);
        if (!groupsMap[groupId]) {
          issues.push({
            severity: 'warning',
            category: 'stale_permission_group',
            item: itemLabel,
            details: 'Группа ' + perm.subject + ' не найдена в Groups (строка ' + node.rowIndex + ')'
          });
        } else {
          var members = readCatalogGroupMembers_(groupsSheet, groupId);
          if (!members.length) {
            issues.push({
              severity: 'warning',
              category: 'stale_permission_empty_group',
              item: itemLabel,
              details: 'Группа «' + groupsMap[groupId].group_name + '» без участников (строка ' + node.rowIndex + ')'
            });
          }
        }
      }
    }
  }

  return issues;
}

/**
 * @param {string} rootFolderId
 * @param {string} importFolderId
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectIntegrityPhysicalFolderIssues_(rootFolderId, importFolderId) {
  var issues = [];
  if (!rootFolderId) {
    return issues;
  }
  var subfolders = listDirectDriveSubfolders_(rootFolderId);
  for (var i = 0; i < subfolders.length; i++) {
    var folder = subfolders[i];
    if (importFolderId && folder.id === importFolderId) {
      continue;
    }
    if (folder.name === IMPORT_DRIVE_FOLDER_NAME) {
      continue;
    }
    issues.push({
      severity: 'warning',
      category: 'physical_drive_folder',
      item: folder.name,
      details: 'Физическая подпапка в root_folder_id (' + folder.id +
        '); виртуальные папки каталога не создают папки на Drive',
      repair_context: { folder_id: folder.id, folder_name: folder.name }
    });
  }
  return issues;
}

/**
 * @param {string} rootFolderId
 * @param {string} catalogOwnerEmail
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectIntegrityRootAclIssues_(rootFolderId, catalogOwnerEmail) {
  var issues = [];
  if (!rootFolderId) {
    return issues;
  }
  var perms;
  try {
    perms = listDrivePermissionsWithIds_(rootFolderId);
  } catch (err) {
    issues.push({
      severity: 'warning',
      category: 'root_acl_unreadable',
      item: rootFolderId,
      details: 'Не удалось прочитать ACL корневой папки: ' + err
    });
    return issues;
  }
  var subjects = [];
  var ownerEmail = normalizeEmail_(catalogOwnerEmail);
  for (var i = 0; i < perms.length; i++) {
    var perm = perms[i];
    if (perm.inherited) {
      continue;
    }
    if (perm.type === 'anyone' || perm.type === 'domain') {
      issues.push({
        severity: 'error',
        category: 'root_acl_public',
        item: rootFolderId,
        details: 'Публичный/доменный доступ на корневой папке: ' + perm.type + ':' + perm.subject,
        repair_context: {
          root_folder_id: rootFolderId,
          permission_id: perm.id
        }
      });
      continue;
    }
    if (perm.type === 'user' && normalizeEmail_(perm.subject) === ownerEmail) {
      continue;
    }
    var label = perm.type === 'group'
      ? 'group:' + perm.subject
      : (perm.subject || perm.type);
    subjects.push(label + ' (' + perm.role + ')');
  }
  if (subjects.length) {
    issues.push({
      severity: 'warning',
      category: 'root_acl_broad',
      item: rootFolderId,
      details: 'Прямой доступ к корневой папке Drive у: ' + subjects.join(', '),
      repair_context: { root_folder_id: rootFolderId }
    });
  }
  return issues;
}

/**
 * @param {Array} catalogFileNodes
 * @param {object} tree
 * @param {Object.<string, Array<string>>} cache
 * @returns {Array<{severity: string, category: string, item: string, details: string}>}
 */
function collectIntegrityApprovedEditorIssues_(catalogFileNodes, tree, cache) {
  var issues = [];
  cache = cache || {};
  var nodesById = tree.nodesById;

  for (var i = 0; i < catalogFileNodes.length; i++) {
    var node = catalogFileNodes[i];
    if (!isCatalogNodeApproved_(node) || !node.file_id) {
      continue;
    }
    var editors = cache[node.file_id];
    if (!editors) {
      var perms = listDrivePermissionsWithIds_(node.file_id);
      editors = [];
      for (var p = 0; p < perms.length; p++) {
        if (perms[p].role === 'writer' && !perms[p].inherited) {
          editors.push(perms[p].subject || perms[p].type);
        }
      }
      cache[node.file_id] = editors;
    }
    if (!editors.length) {
      continue;
    }
    var path = buildCatalogVirtualPath_(node.id, nodesById);
    issues.push({
      severity: 'warning',
      category: 'approved_has_editors',
      item: path || node.name,
      details: 'Утверждён (' + readCell_(node.approved) + '), но на Drive есть редакторы: ' + editors.join(', ')
    });
  }

  return issues;
}

/**
 * @param {string} parentFolderId
 * @returns {Array<{id: string, name: string}>}
 */
function listDirectDriveSubfolders_(parentFolderId) {
  var folders = [];
  if (!parentFolderId) {
    return folders;
  }
  var pageToken = null;
  var query = "'" + parentFolderId + "' in parents and mimeType = '" + DRIVE_FOLDER_MIME +
    "' and trashed = false";

  do {
    var response = Drive.Files.list({
      q: query,
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'nextPageToken, files(id,name)'
    });
    var batch = response.files || [];
    for (var i = 0; i < batch.length; i++) {
      folders.push({ id: batch[i].id, name: batch[i].name });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return folders;
}

/**
 * @param {string} nodeId
 * @param {Object.<string, Object>} nodesById
 * @returns {string}
 */
function buildCatalogVirtualPath_(nodeId, nodesById) {
  var parts = [];
  var current = nodesById[nodeId];
  var guard = 0;
  while (current && guard < 20) {
    guard++;
    if (!isCatalogSystemNode_(current)) {
      parts.unshift(current.name || current.folder_code || current.id);
    }
    if (!current.parent_id) {
      break;
    }
    current = nodesById[current.parent_id];
  }
  return parts.join(' / ');
}

/**
 * @param {Object} fileNode
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object.<string, {match: boolean, summary: string}>} cache
 * @returns {{match: boolean, summary: string}}
 */
function compareCatalogDrivePermissions_(fileNode, ss, cache) {
  if (!fileNode.file_id) {
    return { match: true, summary: '' };
  }
  if (cache[fileNode.file_id]) {
    return cache[fileNode.file_id];
  }

  var catalogPerms = expandCatalogGroupsInPermissions_(
    parseCatalogPermissionsJson_(fileNode.permissions_json),
    ss
  );
  var catalogCanonical = canonicalizePermissionsJson_(JSON.stringify(catalogPerms));

  var drivePerms = getNonInheritedDrivePermissionsForCheck_(fileNode.file_id);
  var driveCanonical = canonicalizePermissionsJson_(JSON.stringify(drivePerms));
  var match = catalogCanonical === driveCanonical;
  var summary = match ? '' : describePermissionsDiff_(catalogPerms, drivePerms);
  var result = { match: match, summary: summary };
  cache[fileNode.file_id] = result;
  return result;
}

/**
 * @param {string} fileId
 * @returns {Array}
 */
function getNonInheritedDrivePermissionsForCheck_(fileId) {
  var all = listDrivePermissionsWithIds_(fileId);
  var result = [];
  for (var i = 0; i < all.length; i++) {
    if (all[i].inherited) {
      continue;
    }
    result.push({
      subject: all[i].subject,
      type: all[i].type,
      level: all[i].level,
      display_name: all[i].subject
    });
  }
  return normalizeCatalogPermissions_(result);
}

/**
 * @param {Array} catalogPerms
 * @param {Array} drivePerms
 * @returns {string}
 */
function describePermissionsDiff_(catalogPerms, drivePerms) {
  var catalogMap = {};
  for (var c = 0; c < catalogPerms.length; c++) {
    var cp = catalogPerms[c];
    catalogMap[(cp.type || 'user') + ':' + normalizeEmail_(cp.subject)] = cp.level;
  }
  var driveMap = {};
  for (var d = 0; d < drivePerms.length; d++) {
    var dp = drivePerms[d];
    driveMap[(dp.type || 'user') + ':' + normalizeEmail_(dp.subject)] = dp.level;
  }

  var parts = [];
  var catalogKeys = Object.keys(catalogMap);
  for (var i = 0; i < catalogKeys.length; i++) {
    var key = catalogKeys[i];
    if (!driveMap[key]) {
      parts.push('нет на Drive: ' + key);
    } else if (driveMap[key] !== catalogMap[key]) {
      parts.push(key + ': каталог=' + catalogMap[key] + ', Drive=' + driveMap[key]);
    }
  }
  var driveKeys = Object.keys(driveMap);
  for (var j = 0; j < driveKeys.length; j++) {
    var dkey = driveKeys[j];
    if (!catalogMap[dkey]) {
      parts.push('лишнее на Drive: ' + dkey + '=' + driveMap[dkey]);
    }
  }
  return parts.join('; ') || 'несовпадение прав';
}

/**
 * @param {object} tree
 * @param {Array} fileRows
 * @returns {Array<{path: string, folder_code: string, file_count: number, total_size: number}>}
 */
function buildIntegrityFolderStatistics_(tree, fileRows) {
  var folderNodes = [];
  for (var i = 0; i < tree.sorted.length; i++) {
    var node = tree.sorted[i];
    if (node.type === 'folder' && !isCatalogSystemNode_(node)) {
      folderNodes.push(node);
    }
  }

  var sizeByFileId = {};
  for (var f = 0; f < fileRows.length; f++) {
    if (fileRows[f].file_id) {
      sizeByFileId[fileRows[f].file_id] = fileRows[f].size || 0;
    }
  }

  var stats = [];
  for (var p = 0; p < folderNodes.length; p++) {
    var folder = folderNodes[p];
    var subtreeCount = 0;
    var subtreeSize = 0;
    for (var s = 0; s < tree.sorted.length; s++) {
      var desc = tree.sorted[s];
      if (desc.type !== 'file') {
        continue;
      }
      if (isDescendantOfFolderNode_(desc, folder.id, tree.nodesById)) {
        subtreeCount++;
        if (desc.file_id && sizeByFileId[desc.file_id] !== undefined) {
          subtreeSize += sizeByFileId[desc.file_id];
        }
      }
    }

    stats.push({
      path: buildCatalogVirtualPath_(folder.id, tree.nodesById),
      folder_code: folder.folder_code || '',
      file_count: subtreeCount,
      total_size: subtreeSize
    });
  }

  stats.sort(function (a, b) {
    return String(a.path).localeCompare(String(b.path), 'ru');
  });
  return stats;
}

/**
 * @param {Array} issues
 * @returns {Object.<string, number>}
 */
function countIntegrityIssuesByCategory_(issues) {
  var counts = { error: 0, warning: 0, info: 0 };
  for (var i = 0; i < issues.length; i++) {
    var sev = issues[i].severity || 'warning';
    counts[sev] = (counts[sev] || 0) + 1;
  }
  return counts;
}

/**
 * @param {number} bytes
 * @returns {number}
 */
function bytesToMegabytes_(bytes) {
  var n = parseInt(bytes, 10) || 0;
  return Math.round(n / (1024 * 1024) * 10) / 10;
}

/**
 * @param {number} bytes
 * @returns {string}
 */
function formatFileSize_(bytes) {
  return bytesToMegabytes_(bytes).toFixed(1) + ' MB';
}

/**
 * @param {string} folderId
 * @returns {Array<{id: string, name: string, mimeType: string, ownerEmail: string, size: number}>}
 */
function listFilesInSingleFolder_(folderId) {
  var files = [];
  var pageToken = null;
  var query = "'" + folderId + "' in parents and trashed = false and mimeType != '" + DRIVE_FOLDER_MIME + "'";

  do {
    var response = Drive.Files.list({
      q: query,
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'nextPageToken, files(id,name,mimeType,size,writersCanShare,owners(emailAddress))'
    });

    var batch = response.files || [];
    for (var i = 0; i < batch.length; i++) {
      var item = batch[i];
      if (item.mimeType === DRIVE_SHORTCUT_MIME) {
        continue;
      }
      var ownerEmail = '';
      if (item.owners && item.owners.length) {
        ownerEmail = item.owners[0].emailAddress || '';
      }
      files.push({
        id: item.id,
        name: item.name,
        mimeType: item.mimeType,
        ownerEmail: ownerEmail,
        size: parseInt(item.size, 10) || 0,
        writersCanShare: item.writersCanShare === true
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return files;
}

/**
 * @param {string} folderId
 * @returns {Array<{id: string, name: string, mimeType: string, ownerEmail: string}>}
 */
function listRootFolderFiles_(folderId) {
  return listAllFilesInFolder_(folderId);
}

/**
 * @param {string} fileId
 * @returns {Array<{subject: string, type: string, level: string, display_name: string}>}
 */
function readDrivePermissions_(fileId) {
  var permissions = [];
  var pageToken = null;

  do {
    var response = Drive.Permissions.list(fileId, {
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      fields: 'nextPageToken, permissions(emailAddress,type,role,displayName,domain)'
    });

    var batch = response.permissions || [];
    for (var i = 0; i < batch.length; i++) {
      var perm = batch[i];
      if (perm.role === 'owner') {
        continue;
      }
      var level = mapDriveRoleToLevel_(perm.role);
      if (!level) {
        continue;
      }

      var subject = '';
      var type = perm.type || 'user';
      if (type === 'user' || type === 'group') {
        subject = perm.emailAddress || '';
      } else if (type === 'domain') {
        subject = perm.domain || '';
      } else if (type === 'anyone') {
        subject = 'anyone';
      }
      if (!subject) {
        continue;
      }

      permissions.push({
        subject: subject,
        type: type,
        level: level,
        display_name: perm.displayName || subject
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return permissions;
}

/**
 * @param {string} role
 * @returns {string|null}
 */
function mapDriveRoleToLevel_(role) {
  if (role === 'writer' || role === 'fileOrganizer' || role === 'organizer') {
    return 'edit';
  }
  if (role === 'commenter') {
    return 'comment';
  }
  if (role === 'reader') {
    return 'read';
  }
  return null;
}

/**
 * @param {Array} permissions
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 * @returns {{editors: string, commenters: string, readers: string}}
 */
function buildPermissionDisplay_(permissions, usersByEmail) {
  var buckets = { edit: [], comment: [], read: [] };
  for (var i = 0; i < permissions.length; i++) {
    var perm = permissions[i];
    if (!buckets[perm.level]) {
      continue;
    }
    buckets[perm.level].push({
      type: perm.type || 'user',
      label: formatPermissionSubject_(perm, usersByEmail)
    });
  }
  return {
    editors: joinSortedPermissionDisplayLabels_(buckets.edit),
    commenters: joinSortedPermissionDisplayLabels_(buckets.comment),
    readers: joinSortedPermissionDisplayLabels_(buckets.read)
  };
}

/**
 * @param {Array<{type: string, label: string}>} items
 * @returns {string}
 */
function joinSortedPermissionDisplayLabels_(items) {
  items.sort(comparePermissionDisplayItems_);
  var labels = [];
  for (var i = 0; i < items.length; i++) {
    labels.push(items[i].label);
  }
  return labels.join(', ');
}

/**
 * @param {{type: string, label: string}} left
 * @param {{type: string, label: string}} right
 * @returns {number}
 */
function comparePermissionDisplayItems_(left, right) {
  var leftIsGroup = left.type === 'group';
  var rightIsGroup = right.type === 'group';
  if (leftIsGroup !== rightIsGroup) {
    return leftIsGroup ? -1 : 1;
  }
  return String(left.label).localeCompare(String(right.label), 'ru');
}

/**
 * @param {string} groupName
 * @returns {string}
 */
function formatCatalogGroupDisplayName_(groupName) {
  var name = String(groupName || '').trim();
  if (!name) {
    return '#';
  }
  if (name.charAt(0) === '#') {
    return name;
  }
  return '#' + name;
}

/**
 * @param {{subject: string, type: string, display_name: string}} perm
 * @param {Object.<string, {email: string, name: string}>} usersByEmail
 * @returns {string}
 */
function formatPermissionSubject_(perm, usersByEmail) {
  if (perm.type === 'group') {
    return formatCatalogGroupDisplayName_(perm.display_name || perm.subject);
  }
  if (perm.type === 'domain') {
    return 'Домен: ' + perm.subject;
  }
  if (perm.type === 'anyone') {
    return 'Все, у кого есть ссылка';
  }
  var user = usersByEmail[normalizeEmail_(perm.subject)];
  return user ? user.name : (perm.display_name || perm.subject);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, {email: string, name: string}>}
 */
function readUsersMap_(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return map;
  }
  var rows = sheet.getRange(2, 1, lastRow, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var email = normalizeEmail_(rows[i][0]);
    if (!email) {
      continue;
    }
    map[email] = {
      email: email,
      name: String(rows[i][1] || email)
    };
  }
  return map;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, boolean>} emails
 * @param {Object.<string, {email: string, name: string}>} existingUsers
 * @returns {number}
 */
function seedUsersFromEmails_(sheet, emails, existingUsers) {
  var added = 0;
  var keys = Object.keys(emails);
  for (var i = 0; i < keys.length; i++) {
    var email = keys[i];
    if (existingUsers[email]) {
      continue;
    }
    sheet.appendRow([email, email, 'true']);
    existingUsers[email] = { email: email, name: email };
    added++;
  }
  return added;
}

/**
 * @param {Array} permissions
 * @param {Object.<string, boolean>} emails
 */
function collectUserEmailsFromPermissions_(permissions, emails) {
  for (var i = 0; i < permissions.length; i++) {
    if (permissions[i].type === 'user') {
      var email = normalizeEmail_(permissions[i].subject);
      if (email) {
        emails[email] = true;
      }
    }
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object.<string, number>} col
 * @returns {Object.<string, boolean>}
 */
function buildFileIdMap_(sheet, col) {
  var map = {};
  var values = readColumnValues_(sheet, col.file_id);
  for (var i = 0; i < values.length; i++) {
    if (values[i]) {
      map[values[i]] = true;
    }
  }
  return map;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} key
 * @returns {string}
 */
function getSetting_(sheet, key) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return '';
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      return String(rows[i][1] || '').trim();
    }
  }
  return '';
}

/**
 * @param {string} fileId
 */
function enforceWritersCanShare_(fileId) {
  try {
    var file = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,writersCanShare'
    });
    if (file.writersCanShare === false || file.writersCanShare === undefined) {
      return;
    }
    Drive.Files.update({ writersCanShare: false }, fileId, null, {
      supportsAllDrives: true,
      fields: 'id,writersCanShare'
    });
  } catch (err) {
    Logger.log('enforceWritersCanShare_ ' + fileId + ': ' + err);
  }
}

/**
 * @param {string} email
 * @returns {string}
 */
function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function showOwnerTransferDialog_() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OwnerTransferDialog')
      .setWidth(520)
      .setHeight(420);
    SpreadsheetApp.getUi().showModalDialog(html, 'Смена владельца');
  } catch (err) {
    // Запуск из редактора Apps Script — UI недоступен, список в результате firstScan.
  }
}

/**
 * Создаёт import_webapp_secret при необходимости.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet
 */
function ensureImportWebAppSettings_(settingsSheet) {
  if (!getSetting_(settingsSheet, 'import_webapp_secret')) {
    upsertSetting_(settingsSheet, 'import_webapp_secret', Utilities.getUuid());
  }
  if (getSetting_(settingsSheet, 'import_webapp_url') === '') {
    upsertSetting_(settingsSheet, 'import_webapp_url', '');
  }
  syncImportWebAppSecretToScriptProperties_(settingsSheet);
}

/**
 * Копирует секрет веб-импорта в ScriptProperties (для doPost без openById).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet
 */
function syncImportWebAppSecretToScriptProperties_(settingsSheet) {
  var secret = getSetting_(settingsSheet, 'import_webapp_secret');
  if (secret) {
    PropertiesService.getScriptProperties().setProperty(IMPORT_WEBAPP_SECRET_KEY, secret);
  }
}

/**
 * @returns {{ok: boolean, error?: string, message?: string}}
 */
function checkImportWebAppConfigured_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss ? ss.getSheetByName('Settings') : null;
  if (!settingsSheet) {
    return {
      ok: false,
      error: 'NO_SETTINGS',
      message: 'Сначала выполните setupSchema.'
    };
  }
  syncImportWebAppSecretToScriptProperties_(settingsSheet);
  var url = getSetting_(settingsSheet, 'import_webapp_url');
  var secret = getSetting_(settingsSheet, 'import_webapp_secret');
  if (!url) {
    return {
      ok: false,
      error: 'NO_WEBAPP_URL',
      message: formatImportWebAppNotConfiguredMessage_()
    };
  }
  if (!secret) {
    return {
      ok: false,
      error: 'NO_WEBAPP_SECRET',
      message: 'Нет import_webapp_secret. Выполните setupSchema.'
    };
  }
  return { ok: true };
}

/**
 * @returns {string}
 */
function formatImportWebAppNotConfiguredMessage_() {
  return 'Импорт с Drive не настроен: нет import_webapp_url.\n' +
    'Владелец каталога: разверните веб-приложение (Execute as: Me, доступ: Все), ' +
    'затем снова выполните setupSchema в редакторе Apps Script.\n\n' +
    buildImportWebAppSetupInstructions_();
}

/**
 * Настройка веб-импорта при инициализации каталога (setupSchema).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet
 * @returns {{configured: boolean, saved?: boolean, url?: string, message?: string, instructions?: string}}
 */
function configureImportWebAppDuringSetup_(settingsSheet) {
  ensureImportWebAppSettings_(settingsSheet);
  var url = getSetting_(settingsSheet, 'import_webapp_url');
  if (url) {
    return {
      configured: true,
      saved: false,
      url: url
    };
  }

  var saveResult = saveImportWebAppUrlFromDeployment_();
  if (saveResult.ok && saveResult.data && saveResult.data.url) {
    return {
      configured: true,
      saved: true,
      url: saveResult.data.url
    };
  }

  return {
    configured: false,
    saved: false,
    message: saveResult.message,
    instructions: buildImportWebAppSetupInstructions_()
  };
}

/**
 * @param {{configured: boolean, saved?: boolean, instructions?: string}} importWebApp
 * @returns {string}
 */
function buildSetupSchemaMessage_(importWebApp) {
  var lines = ['Схема каталога v' + SCHEMA_VERSION + ' готова. Данные — CatalogData, дерево — Catalog.'];
  if (importWebApp.configured) {
    if (importWebApp.saved) {
      lines.push('Веб-импорт: URL развёртывания сохранён в Settings.');
    }
    return lines.join('\n');
  }
  lines.push('');
  lines.push('Веб-импорт с Drive пока не настроен (нет URL развёртывания).');
  if (importWebApp.message) {
    lines.push(importWebApp.message);
  }
  lines.push('');
  lines.push(importWebApp.instructions || buildImportWebAppSetupInstructions_());
  return lines.join('\n');
}

/**
 * @returns {string}
 */
function buildImportWebAppSetupInstructions_() {
  return [
    '1. Развернуть → Новое развёртывание → Веб-приложение',
    '2. Запуск от имени: Я (владелец каталога)',
    '3. Доступ: Все (защита — import_webapp_secret в Settings)',
    '4. Снова выполнить setupSchema — URL подтянется автоматически'
  ].join('\n');
}

/**
 * Сохраняет URL активного развёртывания веб-приложения в Settings.
 * @returns {{ok: boolean, message?: string, data?: {url: string}}}
 */
function saveImportWebAppUrlFromDeployment_() {
  try {
    var url = normalizeImportWebAppUrl_(ScriptApp.getService().getUrl());
    if (!url) {
      return {
        ok: false,
        message: 'Нет URL развёртывания. Создайте развёртывание типа «Веб-приложение».'
      };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      return { ok: false, message: 'Сначала выполните setupSchema.' };
    }
    ensureImportWebAppSettings_(settingsSheet);
    upsertSetting_(settingsSheet, 'import_webapp_url', url);
    return { ok: true, data: { url: url }, message: 'URL сохранён: ' + url };
  } catch (err) {
    return { ok: false, message: String(err) };
  }
}

/**
 * Выполняет copy/promote от имени владельца каталога без HTTP (в том же скрипте).
 *
 * @param {string} action
 * @param {Object} payload
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function invokeImportWebAppActionLocally_(action, payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Лист Settings не найден.');
  }
  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!rootFolderId) {
    throw new Error('Укажите root_folder_id на листе Settings.');
  }
  var importFolderId = ensureCatalogImportDriveFolder_(ss);
  if (action === 'copy') {
    return handleImportWebAppCopy_(rootFolderId, importFolderId, payload || {});
  }
  if (action === 'promote') {
    return handleImportWebAppPromote_(rootFolderId, importFolderId, payload || {});
  }
  throw new Error('Неизвестное действие веб-импорта: ' + action);
}

/**
 * @param {string} spreadsheetId
 * @param {string} action
 * @param {Object} payload
 * @returns {{ok: boolean, data?: object, message?: string}}
 */
function invokeOwnerImportWebApp_(spreadsheetId, action, payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || ss.getId() !== spreadsheetId) {
    throw new Error('Импорт доступен только из привязанной таблицы каталога.');
  }

  if (isSpreadsheetOwner_(ss)) {
    var localResult = invokeImportWebAppActionLocally_(action, payload);
    if (!localResult.ok) {
      throw new Error(localResult.message || localResult.error || 'Ошибка импорта');
    }
    return localResult;
  }

  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Лист Settings не найден.');
  }
  syncImportWebAppSecretToScriptProperties_(settingsSheet);
  var url = normalizeImportWebAppUrl_(getSetting_(settingsSheet, 'import_webapp_url'));
  var secret = getSetting_(settingsSheet, 'import_webapp_secret');
  var rootFolderId = getSetting_(settingsSheet, 'root_folder_id');
  if (!url) {
    throw new Error(formatImportWebAppNotConfiguredMessage_());
  }
  if (!secret) {
    throw new Error('Нет import_webapp_secret. Выполните setupSchema.');
  }
  if (!rootFolderId) {
    throw new Error('Укажите root_folder_id на листе Settings.');
  }
  var importFolderId = ensureCatalogImportDriveFolder_(ss);

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      secret: secret,
      spreadsheet_id: spreadsheetId,
      action: action,
      root_folder_id: rootFolderId,
      import_folder_id: importFolderId,
      payload: payload
    }),
    muteHttpExceptions: true,
    followRedirects: false
  });

  var code = response.getResponseCode();
  var text = response.getContentText() || '';
  var parsed;
  try {
    parsed = JSON.parse(text);
  } catch (parseErr) {
    throw new Error(formatImportWebAppHttpError_(code, url, text));
  }
  if (!parsed.ok) {
    throw new Error(parsed.message || parsed.error || 'Ошибка веб-импорта');
  }
  return parsed;
}

/**
 * @param {string} url
 * @returns {string}
 */
function normalizeImportWebAppUrl_(url) {
  url = String(url || '').trim();
  if (!url) {
    return '';
  }
  if (url.indexOf('/dev') !== -1) {
    url = url.replace('/dev', '/exec');
  }
  return url;
}

/**
 * @param {number} code
 * @param {string} url
 * @param {string} text
 * @returns {string}
 */
function formatImportWebAppHttpError_(code, url, text) {
  var hint = 'Проверьте import_webapp_url и развёртывание веб-приложения.';
  if (code === 404) {
    hint = 'URL веб-приложения не найден (HTTP 404). Владелец каталога: Развернуть → Управление развёртываниями → ' +
      'создайте или обновите «Веб-приложение» (Запуск от имени: Я, доступ: Все), затем выполните setupSchema.';
  } else if (code === 302 || code === 301) {
    hint = 'Веб-приложение перенаправило запрос (HTTP ' + code + '). Используйте URL с окончанием /exec, не /dev.';
  }
  return 'Веб-импорт вернул не JSON (HTTP ' + code + '). ' + hint +
    '\nURL: ' + url + '\n' + text.substring(0, 200);
}

/**
 * Шаг 1: от имени импортёра — чтение исходника для владельца каталога.
 * @param {string} sourceFileId
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @returns {{status: string}}
 */
function grantCatalogOwnerReadAccess_(sourceFileId, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var catalogOwnerEmail = getSpreadsheetOwnerEmail_(ss);
  if (!catalogOwnerEmail || !sourceFileId) {
    throw new Error('Не удалось определить владельца каталога или исходный файл.');
  }
  var importerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (importerEmail && importerEmail === catalogOwnerEmail) {
    return { status: 'already' };
  }
  if (catalogOwnerHasDriveFileAccess_(sourceFileId, catalogOwnerEmail)) {
    return { status: 'already' };
  }
  try {
    Drive.Permissions.create({
      type: 'user',
      role: 'reader',
      emailAddress: catalogOwnerEmail
    }, sourceFileId, {
      supportsAllDrives: true,
      sendNotificationEmail: false
    });
    return { status: 'granted' };
  } catch (err) {
    throw new Error(
      'Импортёр должен открыть доступ владельцу каталога (' + catalogOwnerEmail +
      ') к файлу/папке на Drive. ' + err
    );
  }
}

/**
 * @deprecated Используйте grantCatalogOwnerReadAccess_
 * @param {string} sourceFileId
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 */
function ensureCatalogOwnerCanReadSource_(sourceFileId, ss) {
  grantCatalogOwnerReadAccess_(sourceFileId, ss);
}

/**
 * @param {string} fileId
 * @param {string} catalogOwnerEmail
 * @returns {boolean}
 */
function catalogOwnerHasDriveFileAccess_(fileId, catalogOwnerEmail) {
  catalogOwnerEmail = normalizeEmail_(catalogOwnerEmail);
  if (getDriveFileOwnerEmail_(fileId) === catalogOwnerEmail) {
    return true;
  }
  var pageToken = null;
  do {
    var response = Drive.Permissions.list(fileId, {
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      fields: 'nextPageToken, permissions(emailAddress,type,role)'
    });
    var batch = response.permissions || [];
    for (var i = 0; i < batch.length; i++) {
      if (normalizeEmail_(batch[i].emailAddress || '') === catalogOwnerEmail) {
        return true;
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return false;
}

/**
 * Точка входа веб-приложения (развёртывание: выполнять от имени владельца каталога).
 * @param {Object} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var data = JSON.parse(raw);
    var result = handleImportWebAppRequest_(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: 'WEBAPP_FAILED',
      message: String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * @param {Object} data
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function handleImportWebAppRequest_(data) {
  var secret = data.secret;
  if (!secret) {
    return {
      ok: false,
      error: 'BAD_REQUEST',
      message: 'Нет secret.'
    };
  }

  var expectedSecret = PropertiesService.getScriptProperties().getProperty(IMPORT_WEBAPP_SECRET_KEY);
  if (!expectedSecret || secret !== expectedSecret) {
    return { ok: false, error: 'UNAUTHORIZED', message: 'Неверный import_webapp_secret.' };
  }

  if (data.action === 'copy') {
    return handleImportWebAppCopy_(data.root_folder_id, data.import_folder_id, data.payload || {});
  }
  if (data.action === 'promote') {
    return handleImportWebAppPromote_(data.root_folder_id, data.import_folder_id, data.payload || {});
  }

  return {
    ok: false,
    error: 'UNKNOWN_ACTION',
    message: 'Неизвестное действие: ' + data.action
  };
}

/**
 * Копирует файл на Drive от имени владельца каталога (веб-приложение).
 * @param {string} rootFolderId
 * @param {string} importFolderId
 * @param {Object} payload
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function handleImportWebAppCopy_(rootFolderId, importFolderId, payload) {
  try {
    var sourceFileId = payload.source_file_id;
    var targetName = payload.target_name;
    var permissionsJson = payload.permissions_json || '[]';
    if (!rootFolderId) {
      return { ok: false, error: 'NO_ROOT', message: 'Нет root_folder_id в запросе.' };
    }
    if (!importFolderId) {
      return { ok: false, error: 'NO_IMPORT', message: 'Нет import_folder_id в запросе.' };
    }
    if (!sourceFileId || !targetName) {
      return { ok: false, error: 'BAD_PAYLOAD', message: 'Нет source_file_id или target_name.' };
    }

    var folderCheck = validateRootFolder_(rootFolderId);
    if (!folderCheck.ok) {
      return folderCheck;
    }
    var importCheck = validateRootFolder_(importFolderId);
    if (!importCheck.ok) {
      return importCheck;
    }

    var stagingName = buildImportStagingFileName_(targetName);
    var driveFile;
    try {
      driveFile = Drive.Files.copy({
        name: stagingName,
        parents: [importFolderId]
      }, sourceFileId, {
        supportsAllDrives: true,
        fields: 'id,name,mimeType'
      });
    } catch (copyErr) {
      return {
        ok: false,
        error: 'COPY_FAILED',
        message: 'Владелец каталога не смог скопировать файл. Проверьте доступ к исходнику. ' + copyErr
      };
    }

    if (!driveFile || !driveFile.id) {
      return { ok: false, error: 'COPY_EMPTY', message: 'Drive не вернул id копии.' };
    }

    try {
      promoteDriveFileFromImportToCatalogRoot_(driveFile.id, importFolderId, rootFolderId, targetName);
      syncDrivePermissionsFromJson_(driveFile.id, permissionsJson);
    } catch (permErr) {
      try {
        Drive.Files.remove(driveFile.id, { supportsAllDrives: true });
      } catch (removeErr) {
        Logger.log('handleImportWebAppCopy_ rollback: ' + removeErr);
      }
      return {
        ok: false,
        error: 'PERMISSIONS_FAILED',
        message: 'Копия создана, но не удалось применить права: ' + permErr
      };
    }

    return {
      ok: true,
      data: {
        file_id: driveFile.id,
        name: driveFile.name,
        mime_type: driveFile.mimeType || ''
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'WEBAPP_COPY_FAILED',
      message: String(err)
    };
  }
}

/**
 * Переносит файл из _Import в root от имени владельца каталога (веб-приложение).
 * @param {string} rootFolderId
 * @param {string} importFolderId
 * @param {Object} payload
 * @returns {{ok: boolean, data?: object, error?: string, message?: string}}
 */
function handleImportWebAppPromote_(rootFolderId, importFolderId, payload) {
  try {
    var fileId = payload.file_id;
    var targetName = payload.target_name;
    if (!rootFolderId) {
      return { ok: false, error: 'NO_ROOT', message: 'Нет root_folder_id в запросе.' };
    }
    if (!importFolderId) {
      return { ok: false, error: 'NO_IMPORT', message: 'Нет import_folder_id в запросе.' };
    }
    if (!fileId || !targetName) {
      return { ok: false, error: 'BAD_PAYLOAD', message: 'Нет file_id или target_name.' };
    }

    var folderCheck = validateRootFolder_(rootFolderId);
    if (!folderCheck.ok) {
      return folderCheck;
    }
    var importCheck = validateRootFolder_(importFolderId);
    if (!importCheck.ok) {
      return importCheck;
    }

    var meta = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,parents'
    });
    var parents = meta.parents || [];
    var inRoot = false;
    var inImport = false;
    for (var p = 0; p < parents.length; p++) {
      if (parents[p] === rootFolderId) {
        inRoot = true;
      }
      if (parents[p] === importFolderId) {
        inImport = true;
      }
    }
    if (inRoot && !inImport) {
      return {
        ok: true,
        data: {
          file_id: fileId,
          name: meta.name,
          mime_type: meta.mimeType || '',
          already_promoted: true
        }
      };
    }
    if (!inImport) {
      return {
        ok: false,
        error: 'NOT_IN_IMPORT',
        message: 'Файл не находится в папке Import на Drive.'
      };
    }

    var runnerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
    var fileOwnerEmail = getDriveFileOwnerEmail_(fileId);
    if (runnerEmail && fileOwnerEmail && fileOwnerEmail !== runnerEmail) {
      return {
        ok: false,
        error: 'OWNERSHIP_PENDING',
        message: 'Владелец каталога ещё не принял файл. Примите приглашение на владение в Google Drive и повторите импорт.'
      };
    }

    var promoted = promoteDriveFileFromImportToCatalogRoot_(fileId, importFolderId, rootFolderId, targetName);
    return {
      ok: true,
      data: {
        file_id: fileId,
        name: promoted.name,
        mime_type: promoted.mimeType || ''
      }
    };
  } catch (err) {
    return {
      ok: false,
      error: 'PROMOTE_FAILED',
      message: String(err)
    };
  }
}
