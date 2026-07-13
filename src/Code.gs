/**
 * Cataloger Fast (v0.1)
 *
 * First implementation slice: setup schema in the current Spreadsheet.
 */

const FAST_SCHEMA_VERSION = '0.1';

const FAST_PROP_SCHEMA_VERSION = 'FAST_SCHEMA_VERSION';
const FAST_PROP_CATALOG_ROOT_FOLDER_ID = 'FAST_CATALOG_ROOT_FOLDER_ID';

const FAST_DEFAULT_CATALOG_ROOT_FOLDER_ID = '1lZKFCLzqrrzyiVWHYgF5dtaaOTF3mMT5';

const FAST_FILE_MODE_COPY = 'copy';
const FAST_FILE_MODE_MOVE = 'move';

const FAST_STATUS_PENDING_COPY = 'pending_copy';
const FAST_STATUS_PENDING_MOVE = 'pending_move';

const FAST_SHEETS = /** @type {const} */ ({
  Tree: 'Tree',
  Files: 'Files',
  CatalogACL: 'CatalogACL',
  DriveGrants: 'DriveGrants',
});

const FAST_SCHEMA = /** @type {const} */ ({
  [FAST_SHEETS.Tree]: [
    'folder_id',
    'parent_folder_id',
    'label',
    'prefix_code',
    'created_at',
    'updated_at',
  ],
  [FAST_SHEETS.Files]: [
    'catalog_id',
    'folder_id',
    'file_id',
    'source_file_id',
    'display_name',
    'status',
    'last_error',
    'created_at',
    'updated_at',
  ],
  [FAST_SHEETS.CatalogACL]: [
    'subject',
    'role',
    'folder_id',
    'catalog_id',
    'created_at',
    'updated_at',
  ],
  [FAST_SHEETS.DriveGrants]: [
    'file_id',
    'grantee_email',
    'permission_id',
    'granted_at',
    'expires_at',
  ],
});

/**
 * Ensures Fast schema exists in the current spreadsheet.
 *
 * Contract:
 * - Idempotent.
 * - Creates missing sheets and writes headers.
 * - If a sheet exists but header doesn't match exactly -> throws SCHEMA_MISMATCH.
 * - Stores schema version and catalog root folder id in DocumentProperties.
 *
 * @param {string} catalogRootFolderId Drive folder id for the catalog root (required on first run).
 * @returns {{ok:true, createdSheets:string[], schemaVersion:string, catalogRootFolderId:string}}
 */
function setupSchemaFast(catalogRootFolderId) {
  if (!catalogRootFolderId || typeof catalogRootFolderId !== 'string') {
    throw new Error('setupSchemaFast: catalogRootFolderId is required (Drive folder id string).');
  }

  const props = PropertiesService.getDocumentProperties();

  const existingSchemaVersion = props.getProperty(FAST_PROP_SCHEMA_VERSION);
  if (existingSchemaVersion && existingSchemaVersion !== FAST_SCHEMA_VERSION) {
    throw new Error(
      `SCHEMA_VERSION_MISMATCH: expected ${FAST_SCHEMA_VERSION}, got ${existingSchemaVersion}`
    );
  }

  const existingRoot = props.getProperty(FAST_PROP_CATALOG_ROOT_FOLDER_ID);
  if (existingRoot && existingRoot !== catalogRootFolderId) {
    throw new Error(
      `CATALOG_ROOT_MISMATCH: configured ${existingRoot}, requested ${catalogRootFolderId}`
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const createdSheets = [];

  Object.keys(FAST_SCHEMA).forEach((sheetName) => {
    const expectedHeader = FAST_SCHEMA[sheetName];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]);
      sheet.setFrozenRows(1);
      createdSheets.push(sheetName);
      return;
    }

    const actualHeader = _readHeaderRowStrict_(sheet, expectedHeader.length);
    if (!_arraysEqual_(actualHeader, expectedHeader)) {
      throw new Error(
        `SCHEMA_MISMATCH: sheet=${sheetName} expected=[${expectedHeader.join(
          ','
        )}] actual=[${actualHeader.join(',')}]`
      );
    }
  });

  props.setProperty(FAST_PROP_SCHEMA_VERSION, FAST_SCHEMA_VERSION);
  props.setProperty(FAST_PROP_CATALOG_ROOT_FOLDER_ID, catalogRootFolderId);

  return {
    ok: true,
    createdSheets,
    schemaVersion: FAST_SCHEMA_VERSION,
    catalogRootFolderId: props.getProperty(FAST_PROP_CATALOG_ROOT_FOLDER_ID) || '',
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} width
 * @returns {string[]}
 */
function _readHeaderRowStrict_(sheet, width) {
  const values = sheet.getRange(1, 1, 1, width).getValues();
  const row = values && values[0] ? values[0] : [];
  return row.map((v) => (v === null || v === undefined ? '' : String(v)).trim());
}

/**
 * @param {unknown[]} a
 * @param {unknown[]} b
 * @returns {boolean}
 */
function _arraysEqual_(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i]) !== String(b[i])) return false;
  }
  return true;
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function readCell_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

/**
 * @param {string} label
 * @returns {string}
 */
function sanitizeFolderLabel_(label) {
  return String(label || '')
    .trim()
    .replace(/[/\\]/g, '-');
}

/**
 * @param {string[]} headers
 * @returns {Object.<string, number>}
 */
function columnIndexMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    map[headers[i]] = i + 1;
  }
  return map;
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _getTreeSheetOrThrow_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FAST_SHEETS.Tree);
  if (!sheet) {
    throw new Error('SCHEMA_NOT_INITIALIZED: Tree sheet missing, run setupSchemaFast first.');
  }
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{folder_id:string, parent_folder_id:string, label:string, prefix_code:string}[]}
 */
function _readTreeDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const width = FAST_SCHEMA[FAST_SHEETS.Tree].length;
  const numRows = lastRow - 1;
  const values = sheet.getRange(2, 1, numRows, width).getValues();

  return values.map((row) => ({
    folder_id: readCell_(row[0]),
    parent_folder_id: readCell_(row[1]),
    label: readCell_(row[2]),
    prefix_code: readCell_(row[3]),
  }));
}

/**
 * Creates a virtual folder row in Tree (Sheets only, no Drive).
 *
 * @param {string} parentFolderId Parent folder UUID, or '' for root.
 * @param {string} label Folder label without prefixes.
 * @param {string=} prefixCode Optional prefix code for navigation/search.
 * @returns {{
 *   ok:true,
 *   folder_id:string,
 *   parent_folder_id:string,
 *   label:string,
 *   prefix_code:string,
 *   created_at:string,
 *   updated_at:string
 * }}
 */
function createFolderFast(parentFolderId, label, prefixCode) {
  const parentId = readCell_(parentFolderId);
  const sanitizedLabel = sanitizeFolderLabel_(label);
  if (!sanitizedLabel) {
    throw new Error('INVALID_LABEL: label is required.');
  }

  const prefix = readCell_(prefixCode);
  const sheet = _getTreeSheetOrThrow_();
  const rows = _readTreeDataRows_(sheet);

  if (parentId) {
    const parentExists = rows.some((row) => row.folder_id === parentId);
    if (!parentExists) {
      throw new Error(`PARENT_NOT_FOUND: ${parentId}`);
    }
  }

  const duplicate = rows.some(
    (row) => row.parent_folder_id === parentId && row.label === sanitizedLabel
  );
  if (duplicate) {
    throw new Error(`DUPLICATE_LABEL: "${sanitizedLabel}" under parent="${parentId}"`);
  }

  const now = new Date().toISOString();
  const folderId = Utilities.getUuid();
  sheet.appendRow([folderId, parentId, sanitizedLabel, prefix, now, now]);

  return {
    ok: true,
    folder_id: folderId,
    parent_folder_id: parentId,
    label: sanitizedLabel,
    prefix_code: prefix,
    created_at: now,
    updated_at: now,
  };
}

/**
 * @param {string} folderId
 * @returns {boolean}
 */
function _folderExistsInTree_(folderId) {
  const rows = _readTreeDataRows_(_getTreeSheetOrThrow_());
  return rows.some((row) => row.folder_id === folderId);
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _getFilesSheetOrThrow_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FAST_SHEETS.Files);
  if (!sheet) {
    throw new Error('SCHEMA_NOT_INITIALIZED: Files sheet missing, run setupSchemaFast first.');
  }
  return sheet;
}

/**
 * @returns {string}
 */
function _getCatalogOwnerEmail_() {
  const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  if (!owner) {
    throw new Error('CATALOG_OWNER_NOT_FOUND: spreadsheet owner is unavailable.');
  }
  return owner.getEmail().toLowerCase();
}

/**
 * @param {string} sourceFileId
 * @returns {string[]}
 */
function _getDriveFileOwnerEmails_(sourceFileId) {
  const file = Drive.Files.get(sourceFileId, {
    fields: 'owners(emailAddress)',
    supportsAllDrives: true,
  });
  const owners = file && file.owners ? file.owners : [];
  return owners
    .map((owner) => readCell_(owner.emailAddress).toLowerCase())
    .filter((email) => email);
}

/**
 * @param {string=} mode
 * @returns {'copy'|'move'}
 */
function _normalizeImportMode_(mode) {
  const normalized = readCell_(mode).toLowerCase();
  if (!normalized || normalized === FAST_FILE_MODE_COPY) {
    return FAST_FILE_MODE_COPY;
  }
  if (normalized === FAST_FILE_MODE_MOVE) {
    return FAST_FILE_MODE_MOVE;
  }
  throw new Error(`INVALID_MODE: expected copy or move, got "${mode}"`);
}

/**
 * Creates a file record in Files (Sheets only). Drive copy/move runs later in a worker.
 *
 * @param {string} folderId Virtual folder UUID from Tree.
 * @param {string} sourceFileId Drive file id of the source.
 * @param {string} displayName Display name in catalog (no prefixes).
 * @param {string=} mode 'copy' (default) or 'move'.
 * @returns {{
 *   ok:true,
 *   catalog_id:string,
 *   folder_id:string,
 *   source_file_id:string,
 *   display_name:string,
 *   status:string,
 *   file_id:string,
 *   created_at:string,
 *   updated_at:string
 * }}
 */
function createFileRecordFast(folderId, sourceFileId, displayName, mode) {
  const targetFolderId = readCell_(folderId);
  const sourceId = readCell_(sourceFileId);
  const sanitizedName = sanitizeFolderLabel_(displayName);
  const importMode = _normalizeImportMode_(mode);

  if (!targetFolderId) {
    throw new Error('FOLDER_NOT_FOUND: folderId is required.');
  }
  if (!sourceId) {
    throw new Error('INVALID_SOURCE_FILE_ID: sourceFileId is required.');
  }
  if (!sanitizedName) {
    throw new Error('INVALID_DISPLAY_NAME: displayName is required.');
  }
  if (!_folderExistsInTree_(targetFolderId)) {
    throw new Error(`FOLDER_NOT_FOUND: ${targetFolderId}`);
  }

  if (importMode === FAST_FILE_MODE_MOVE) {
    const catalogOwnerEmail = _getCatalogOwnerEmail_();
    const ownerEmails = _getDriveFileOwnerEmails_(sourceId);
    const ownedByCatalogOwner = ownerEmails.indexOf(catalogOwnerEmail) !== -1;
    if (!ownedByCatalogOwner) {
      throw new Error(`MOVE_NOT_ALLOWED: source file owner is not catalog owner (${catalogOwnerEmail}).`);
    }
  }

  const status =
    importMode === FAST_FILE_MODE_MOVE ? FAST_STATUS_PENDING_MOVE : FAST_STATUS_PENDING_COPY;
  const now = new Date().toISOString();
  const catalogId = Utilities.getUuid();
  const sheet = _getFilesSheetOrThrow_();

  sheet.appendRow([catalogId, targetFolderId, '', sourceId, sanitizedName, status, '', now, now]);

  return {
    ok: true,
    catalog_id: catalogId,
    folder_id: targetFolderId,
    source_file_id: sourceId,
    display_name: sanitizedName,
    status,
    file_id: '',
    created_at: now,
    updated_at: now,
  };
}

/**
 * Spreadsheet open hook: builds the Fast catalog menu.
 */
function onOpen() {
  buildCatalogFastMenu_();
}

/**
 * Manual menu refresh (useful after clasp push without reload).
 */
function buildCatalogFastMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📁 Каталогизатор Fast')
    .addItem('⚙️ Инициализация схемы', 'catalogFastInitSchemaMenu_')
    .addSeparator()
    .addItem('📂 Создать корневую папку', 'catalogFastCreateRootFolderMenu_')
    .addItem('📂 Создать папку', 'catalogFastCreateFolderMenu_')
    .addItem('📄 Добавить запись файла', 'catalogFastCreateFileRecordMenu_')
    .addSeparator()
    .addItem('🔄 Обновить меню', 'buildCatalogFastMenu_')
    .addToUi();
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 * @param {string} title
 * @param {string} message
 */
function catalogFastShowMessage_(ui, title, message) {
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 * @param {unknown} err
 */
function catalogFastShowError_(ui, err) {
  const message = err && typeof err === 'object' && 'message' in err ? String(err.message) : String(err);
  ui.alert('Ошибка', message, ui.ButtonSet.OK);
}

function catalogFastInitSchemaMenu_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = setupSchemaFast(FAST_DEFAULT_CATALOG_ROOT_FOLDER_ID);
    const created = result.createdSheets.length ? result.createdSheets.join(', ') : 'нет (уже были)';
    catalogFastShowMessage_(ui, 'Схема готова', `Созданы листы: ${created}`);
  } catch (err) {
    catalogFastShowError_(ui, err);
  }
}

function catalogFastCreateRootFolderMenu_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Создать корневую папку', 'Имя папки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  try {
    const result = createFolderFast('', response.getResponseText());
    catalogFastShowMessage_(
      ui,
      'Папка создана',
      `ID: ${result.folder_id}\nИмя: ${result.label}`
    );
  } catch (err) {
    catalogFastShowError_(ui, err);
  }
}

function catalogFastCreateFolderMenu_() {
  const ui = SpreadsheetApp.getUi();
  const parentResponse = ui.prompt(
    'Создать папку',
    'parent_folder_id (UUID родителя):',
    ui.ButtonSet.OK_CANCEL
  );
  if (parentResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const labelResponse = ui.prompt('Создать папку', 'Имя папки:', ui.ButtonSet.OK_CANCEL);
  if (labelResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  try {
    const result = createFolderFast(parentResponse.getResponseText(), labelResponse.getResponseText());
    catalogFastShowMessage_(
      ui,
      'Папка создана',
      `ID: ${result.folder_id}\nИмя: ${result.label}`
    );
  } catch (err) {
    catalogFastShowError_(ui, err);
  }
}

function catalogFastCreateFileRecordMenu_() {
  const ui = SpreadsheetApp.getUi();
  const folderResponse = ui.prompt(
    'Добавить запись файла',
    'folder_id (UUID виртуальной папки):',
    ui.ButtonSet.OK_CANCEL
  );
  if (folderResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const sourceResponse = ui.prompt(
    'Добавить запись файла',
    'source_file_id (Drive ID исходника):',
    ui.ButtonSet.OK_CANCEL
  );
  if (sourceResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const nameResponse = ui.prompt(
    'Добавить запись файла',
    'display_name (имя в каталоге):',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const modeResponse = ui.prompt(
    'Добавить запись файла',
    'mode: copy (default) или move:',
    ui.ButtonSet.OK_CANCEL
  );
  if (modeResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  try {
    const modeText = readCell_(modeResponse.getResponseText()) || FAST_FILE_MODE_COPY;
    const result = createFileRecordFast(
      folderResponse.getResponseText(),
      sourceResponse.getResponseText(),
      nameResponse.getResponseText(),
      modeText
    );
    catalogFastShowMessage_(
      ui,
      'Запись файла создана',
      `catalog_id: ${result.catalog_id}\nstatus: ${result.status}\nname: ${result.display_name}`
    );
  } catch (err) {
    catalogFastShowError_(ui, err);
  }
}

/**
 * One-time bootstrap for the current spreadsheet.
 * Run this from the Apps Script editor, then delete or leave unused.
 */
function initOnceFast() {
  return setupSchemaFast(FAST_DEFAULT_CATALOG_ROOT_FOLDER_ID);
}

