/**
 * §18 — выгрузка структуры каталога в новую Google Таблицу (только Управляющий).
 * Иконка + Тип + Размер (МБ) + L1…L10 + права; outline; бело-голубое чередование.
 *
 * @returns {{
 *   ok: true,
 *   spreadsheetId: string,
 *   url: string,
 *   title: string,
 *   rowCount: number
 * }}
 */
function exportCatalogStructure() {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);

  var engine = createAclEngine_();
  var rootId = getVirtualRootFolderId_();
  if (!engine.foldersById[rootId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Корневая папка каталога не найдена.');
  }

  var treeRows = [];
  Object.keys(engine.foldersById || {}).forEach(function (id) {
    treeRows.push(engine.foldersById[id]);
  });
  var fileRows = [];
  Object.keys(engine.filesByCatalogId || {}).forEach(function (id) {
    fileRows.push(engine.filesByCatalogId[id]);
  });
  var folderSizes = buildFolderSizeIndex_(treeRows, fileRows);

  var ownerEmails = collectCatalogOwnerEmails_();
  var header = [
    'Иконка',
    'Тип',
    'Размер',
    'L1',
    'L2',
    'L3',
    'L4',
    'L5',
    'L6',
    'L7',
    'L8',
    'L9',
    'L10',
    'Имя',
    'Утверждено',
    'Редакторы',
    'Комментаторы',
    'Читатели'
  ];
  var dataRows = [];
  var groups = [];
  var iconCache = {};
  walkExportFolder_(
    engine,
    rootId,
    0,
    [],
    true,
    ownerEmails,
    folderSizes,
    dataRows,
    groups,
    iconCache
  );

  var tz = Session.getScriptTimeZone() || 'Europe/Moscow';
  var title =
    'Каталог — структура ' +
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var ss = SpreadsheetApp.create(title);
  var sheet = ss.getActiveSheet();
  sheet.setName('Структура');

  var values = [header].concat(dataRows);
  sheet.getRange(1, 1, values.length, header.length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setColumnWidth(1, 36);
  sheet.setColumnWidth(2, 56);
  sheet.setColumnWidth(3, 72);
  for (var c = 4; c <= 13; c++) {
    sheet.setColumnWidth(c, 28);
  }
  sheet.setColumnWidth(14, 320);
  try {
    sheet.autoResizeColumns(15, header.length - 14);
  } catch (e) {
    // ignore
  }

  applyExportRowGroups_(sheet, groups, dataRows.length);
  try {
    sheet.expandAllRowGroups();
  } catch (e2) {
    // ignore if no groups
  }
  // Banding после outline — applyRowBanding с CellImage часто молча не рисует.
  applyExportRowBanding_(sheet, dataRows.length, header.length);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    url: ss.getUrl(),
    title: title,
    rowCount: dataRows.length
  };
}

/**
 * @returns {Object.<string, boolean>}
 */
function collectCatalogOwnerEmails_() {
  var out = {};
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  if (controllerEmail) {
    out[String(controllerEmail).toLowerCase()] = true;
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getOwner();
    if (owner && owner.getEmail) {
      var email = String(owner.getEmail() || '').trim();
      if (email) {
        out[email.toLowerCase()] = true;
      }
    }
  } catch (e) {
    // getOwner may fail in some contexts
  }
  return out;
}

/**
 * @param {Object} engine
 * @param {string} folderId
 * @param {number} depth 0 = корень → L1
 * @param {string[]} pathNames имена предков (без текущего)
 * @param {boolean} isRoot
 * @param {Object.<string, boolean>} ownerEmails
 * @param {Object.<string, number>} folderSizes bytes by folder_id
 * @param {Array<Array<*>>} outRows
 * @param {Array<{ start: number, end: number }>} groups 0-based indices in outRows
 * @param {Object.<string, Object>} iconCache
 */
function walkExportFolder_(
  engine,
  folderId,
  depth,
  pathNames,
  isRoot,
  ownerEmails,
  folderSizes,
  outRows,
  groups,
  iconCache
) {
  var folder = engine.foldersById[folderId];
  if (!folder) {
    return;
  }
  var name = String(folder.name || folderId);
  outRows.push(
    buildExportStructureRow_({
      icon: buildExportCellIcon_(
        iconCache,
        exportIconUrl_('folder', '', 'application/vnd.google-apps.folder'),
        'Папка'
      ),
      typeLabel: 'DIR',
      sizeMb: exportSizeMb_(folderSizes[folderId] || 0),
      depth: depth,
      name: name,
      pathNames: pathNames,
      approvedLabel: '—',
      acl: formatExportAclColumns_(engine, 'folder', folderId, isRoot, ownerEmails)
    })
  );

  var childPath = pathNames.concat([name]);
  var childrenStart = outRows.length;

  var childFolders = [];
  Object.keys(engine.foldersById || {}).forEach(function (id) {
    var row = engine.foldersById[id];
    if (String(row.parent_folder_id || '') === String(folderId)) {
      childFolders.push({
        id: id,
        name: String(row.name || id)
      });
    }
  });
  childFolders.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' });
  });
  childFolders.forEach(function (child) {
    walkExportFolder_(
      engine,
      child.id,
      depth + 1,
      childPath,
      false,
      ownerEmails,
      folderSizes,
      outRows,
      groups,
      iconCache
    );
  });

  var files = [];
  Object.keys(engine.filesByCatalogId || {}).forEach(function (catalogId) {
    var file = engine.filesByCatalogId[catalogId];
    if (String(file.folder_id || '') !== String(folderId)) {
      return;
    }
    files.push({
      id: catalogId,
      name: String(file.display_name || catalogId),
      mimeType: String(file.mime_type || ''),
      sizeBytes: parseNumber_(file.size_bytes) || 0,
      approved: parseBoolean_(file.approved),
      approvedBy: String(file.approved_by || '').trim()
    });
  });
  files.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' });
  });
  files.forEach(function (file) {
    var approvedLabel = '—';
    if (file.approved) {
      approvedLabel =
        resolveUserLabelFromEngine_(engine, file.approvedBy) || file.approvedBy || '—';
    }
    var typeLabel = exportTypeLabel_(file.name, file.mimeType);
    var mime = resolveExportMimeType_(file.name, file.mimeType);
    outRows.push(
      buildExportStructureRow_({
        icon: buildExportCellIcon_(
          iconCache,
          exportIconUrl_('file', file.name, mime),
          typeLabel
        ),
        typeLabel: typeLabel,
        sizeMb: exportSizeMb_(file.sizeBytes),
        depth: depth + 1,
        name: file.name,
        pathNames: childPath,
        approvedLabel: approvedLabel,
        acl: formatExportAclColumns_(engine, 'file', file.id, false, ownerEmails)
      })
    );
  });

  var childrenEnd = outRows.length - 1;
  // Корень не группируем — плюсик только у вложенных папок.
  if (!isRoot && childrenEnd >= childrenStart) {
    groups.push({ start: childrenStart, end: childrenEnd });
  }
}

/**
 * Байты → МБ (1024²), 2 знака; без суффикса — для формул Sheets.
 *
 * @param {number} bytes
 * @returns {number}
 */
function exportSizeMb_(bytes) {
  var n = Math.max(0, Number(bytes) || 0);
  return Math.round((n / (1024 * 1024)) * 100) / 100;
}

/**
 * Чередование белый / голубой на строках данных (шапка без заливки).
 * Явный setBackgrounds — надёжнее applyRowBanding при CellImage в колонке A.
 * getRange(row, column, numRows, numColumns).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} dataRowCount
 * @param {number} columnCount
 */
function applyExportRowBanding_(sheet, dataRowCount, columnCount) {
  if (!dataRowCount || dataRowCount < 1 || !columnCount) {
    return;
  }
  var white = '#FFFFFF';
  var blue = '#DDEBF7';
  var colors = [];
  for (var r = 0; r < dataRowCount; r++) {
    var fill = r % 2 === 0 ? white : blue;
    var row = [];
    for (var c = 0; c < columnCount; c++) {
      row.push(fill);
    }
    colors.push(row);
  }
  sheet.getRange(2, 1, dataRowCount, columnCount).setBackgrounds(colors);
}

/**
 * Группы: sheet row = dataIndex + 2 (строка 1 — заголовок).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<{ start: number, end: number }>} groups
 * @param {number} dataRowCount
 */
function applyExportRowGroups_(sheet, groups, dataRowCount) {
  if (!groups || !groups.length || dataRowCount < 2) {
    return;
  }
  // Сначала более глубокие/короткие диапазоны — вложенность outline.
  var sorted = groups.slice().sort(function (a, b) {
    var lenA = a.end - a.start;
    var lenB = b.end - b.start;
    if (lenA !== lenB) {
      return lenA - lenB;
    }
    return a.start - b.start;
  });
  sorted.forEach(function (g) {
    if (g.end < g.start) {
      return;
    }
    var sheetStart = g.start + 2;
    var sheetEnd = g.end + 2;
    try {
      sheet.getRange(sheetStart + ':' + sheetEnd).shiftRowGroupDepth(1);
    } catch (e) {
      try {
        sheet.getRange(sheetStart, 1, sheetEnd, 1).shiftRowGroupDepth(1);
      } catch (e2) {
        // ignore
      }
    }
  });
}

/**
 * @param {{
 *   icon: Object,
 *   typeLabel: string,
 *   sizeMb: number,
 *   depth: number,
 *   name: string,
 *   pathNames: string[],
 *   approvedLabel: string,
 *   acl: { editors: string, commenters: string, readers: string }
 * }} spec
 * @returns {Array<*>}
 */
function buildExportStructureRow_(spec) {
  var levels = ['', '', '', '', '', '', '', '', '', ''];
  var depth = Number(spec.depth) || 0;
  var col = Math.min(Math.max(depth, 0), 9);
  var cellName = String(spec.name || '');
  if (depth >= 10) {
    cellName = '…/' + cellName;
  }
  levels[col] = cellName;

  var sizeMb = Number(spec.sizeMb);
  if (!isFinite(sizeMb) || sizeMb < 0) {
    sizeMb = 0;
  }

  return [spec.icon, spec.typeLabel, sizeMb]
    .concat(levels)
    .concat([
      '',
      spec.approvedLabel || '—',
      spec.acl.editors,
      spec.acl.commenters,
      spec.acl.readers
    ]);
}

/**
 * @param {Object.<string, Object>} iconCache
 * @param {string} url
 * @param {string} alt
 * @returns {Object|string}
 */
function buildExportCellIcon_(iconCache, url, alt) {
  url = String(url || '').trim();
  if (!url) {
    return '';
  }
  if (iconCache[url]) {
    return iconCache[url];
  }
  try {
    var image = SpreadsheetApp.newCellImage()
      .setSourceUrl(url)
      .setAltTextTitle(String(alt || ''))
      .build();
    iconCache[url] = image;
    return image;
  } catch (e) {
    return '';
  }
}

/**
 * Иконки Drive third-party по MIME (16px).
 *
 * @param {'folder'|'file'} kind
 * @param {string} name
 * @param {string} mimeType
 * @returns {string}
 */
function exportIconUrl_(kind, name, mimeType) {
  var mime =
    kind === 'folder'
      ? 'application/vnd.google-apps.folder'
      : resolveExportMimeType_(name, mimeType);
  return 'https://drive-thirdparty.googleusercontent.com/16/type/' + mime;
}

/**
 * @param {string} name
 * @param {string} mimeType
 * @returns {string}
 */
function resolveExportMimeType_(name, mimeType) {
  var mime = String(mimeType || '').trim().toLowerCase();
  if (mime) {
    return mime;
  }
  var fileName = String(name || '').toLowerCase();
  var ext = '';
  var dot = fileName.lastIndexOf('.');
  if (dot > 0 && dot < fileName.length - 1) {
    ext = fileName.slice(dot + 1);
  }
  var byExt = {
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt: 'application/vnd.ms-powerpoint',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    zip: 'application/zip',
    rar: 'application/vnd.rar',
    '7z': 'application/x-7z-compressed',
    mp3: 'audio/mpeg',
    mp4: 'video/mp4',
    txt: 'text/plain',
    csv: 'text/csv'
  };
  if (ext && byExt[ext]) {
    return byExt[ext];
  }
  return 'application/octet-stream';
}

/**
 * @param {string} name
 * @param {string} mimeType
 * @returns {string}
 */
function exportTypeLabel_(name, mimeType) {
  var mime = String(mimeType || '').toLowerCase();
  var fileName = String(name || '');
  var ext = '';
  var dot = fileName.lastIndexOf('.');
  if (dot > 0 && dot < fileName.length - 1) {
    ext = fileName.slice(dot + 1).toUpperCase();
  }

  if (mime === 'application/pdf' || ext === 'PDF') {
    return 'PDF';
  }
  if (mime.indexOf('application/vnd.google-apps.') === 0) {
    var key = mime.slice('application/vnd.google-apps.'.length);
    if (key === 'document') {
      return 'Docs';
    }
    if (key === 'spreadsheet') {
      return 'Sheets';
    }
    if (key === 'presentation') {
      return 'Slides';
    }
    if (key === 'form') {
      return 'Forms';
    }
    if (key === 'drawing') {
      return 'Drawing';
    }
    if (key) {
      return key.charAt(0).toUpperCase() + key.slice(1);
    }
  }
  if (ext) {
    return ext;
  }
  if (mime.indexOf('/') >= 0) {
    return mime.split('/').pop() || 'файл';
  }
  return 'файл';
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @param {boolean} isRoot
 * @param {Object.<string, boolean>} ownerEmails
 * @returns {{ editors: string, commenters: string, readers: string }}
 */
function formatExportAclColumns_(engine, objectType, objectId, isRoot, ownerEmails) {
  if (isRoot) {
    var labels = getEffectiveAclDisplayFromEngine_(engine, objectType, objectId);
    return {
      editors: formatExportLabelList_(
        filterExportOwnerLabels_(engine, labels.editors || [], ownerEmails)
      ),
      commenters: formatExportLabelList_(
        filterExportOwnerLabels_(engine, labels.commenters || [], ownerEmails)
      ),
      readers: formatExportLabelList_(
        filterExportOwnerLabels_(engine, labels.readers || [], ownerEmails)
      )
    };
  }

  var rows = engine.aclByObject[objectType + ':' + objectId] || [];
  var editors = [];
  var commenters = [];
  var readers = [];
  var seen = {};

  rows.forEach(function (row) {
    var delta = normalizeAclDelta_(row);
    if (delta !== '+' && delta !== '-') {
      if (delta === '' || delta === 'base') {
        delta = '+';
      } else {
        return;
      }
    }
    var principalType = String(row.principal_type || '').trim().toLowerCase();
    var principalId = String(row.principal_id || '').trim();
    if (!principalId) {
      return;
    }
    if (principalType === 'user' && ownerEmails[principalId.toLowerCase()]) {
      return;
    }

    var level = normalizePermissionLevel_(row.permission_level);
    if (!level || level === 'none') {
      return;
    }

    var label =
      principalType === 'group'
        ? formatGroupAclLabel_(engine, principalId)
        : resolveUserLabelFromEngine_(engine, principalId);
    if (!label) {
      return;
    }
    var signed = delta + label;
    var dedupe = delta + ':' + principalType + ':' + principalId.toLowerCase() + ':' + level;
    if (seen[dedupe]) {
      return;
    }
    seen[dedupe] = true;

    if (level === 'editor') {
      editors.push(signed);
    } else if (level === 'commenter') {
      commenters.push(signed);
    } else if (level === 'reader') {
      readers.push(signed);
    }
  });

  return {
    editors: formatExportLabelList_(editors),
    commenters: formatExportLabelList_(commenters),
    readers: formatExportLabelList_(readers)
  };
}

/**
 * @param {Object} engine
 * @param {string[]} labels
 * @param {Object.<string, boolean>} ownerEmails
 * @returns {string[]}
 */
function filterExportOwnerLabels_(engine, labels, ownerEmails) {
  var ownerLabels = {};
  Object.keys(ownerEmails || {}).forEach(function (email) {
    var label = resolveUserLabelFromEngine_(engine, email);
    if (label) {
      ownerLabels[String(label).toLowerCase()] = true;
    }
    ownerLabels[email.toLowerCase()] = true;
  });
  return (labels || []).filter(function (label) {
    return !ownerLabels[String(label || '').toLowerCase()];
  });
}

/**
 * @param {string[]} labels
 * @returns {string}
 */
function formatExportLabelList_(labels) {
  if (!labels || !labels.length) {
    return '—';
  }
  return labels.join(', ');
}
