/**
 * §19 — Ремонт каталога: быстрая / глубокая проверка → отчёт; fix пакетами одного кода.
 */

/** @const {string} */
var REPAIR_LEVEL_QUICK_ = 'быстрая';
/** @const {string} */
var REPAIR_LEVEL_DEEP_ = 'глубокая';

/**
 * @param {'quick'|'deep'} mode
 * @returns {{
 *   ok: true,
 *   mode: string,
 *   spreadsheetId: string,
 *   url: string,
 *   title: string,
 *   issueCount: number,
 *   codes: Array<{ code: string, count: number, proposedFix: string, fixable: boolean }>
 * }}
 */
function runCatalogRepairCheck(mode) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);

  mode = String(mode || 'quick').trim().toLowerCase() === 'deep' ? 'deep' : 'quick';
  var issues = collectCatalogRepairIssues_(mode);
  var report = writeCatalogRepairReport_(mode, issues);
  return {
    ok: true,
    mode: mode,
    spreadsheetId: report.spreadsheetId,
    url: report.url,
    title: report.title,
    issueCount: issues.length,
    codes: summarizeRepairCodes_(issues)
  };
}

/**
 * Применить авто-fix для всех найденных сейчас проблем с данным кодом.
 *
 * @param {string} code
 * @param {'quick'|'deep'=} mode контекст перескана (для deep-кодов нужен deep)
 * @returns {{ ok: true, code: string, fixed: number, skipped: number }}
 */
function fixCatalogRepairByCode(code, mode) {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  assertIsCatalogController_(userEmail);

  code = String(code || '').trim();
  if (!code) {
    throw catalogError_('INVALID_INPUT', 'code is required.');
  }
  if (!isRepairCodeFixable_(code)) {
    throw catalogError_('NOT_ALLOWED', 'Для кода ' + code + ' авто-исправление не предусмотрено.');
  }

  mode = String(mode || 'quick').trim().toLowerCase() === 'deep' ? 'deep' : 'quick';
  if (String(code).indexOf('DRIVE_') === 0) {
    mode = 'deep';
  }

  var issues = collectCatalogRepairIssues_(mode).filter(function (issue) {
    return issue.code === code;
  });
  var fixed = 0;
  var skipped = 0;
  var needAclCacheRebuild = false;
  var bulkDone = {};

  issues.forEach(function (issue) {
    try {
      var bulkKey = '';
      if (
        issue.code === 'TREE_MISSING_ID' ||
        issue.code === 'FILE_MISSING_ID' ||
        issue.code === 'ACL_CACHE_MISMATCH'
      ) {
        bulkKey = issue.code;
      }
      if (bulkKey && bulkDone[bulkKey]) {
        fixed++;
        return;
      }
      if (applyRepairFixForIssue_(issue)) {
        fixed++;
        if (bulkKey) {
          bulkDone[bulkKey] = true;
        }
        if (String(issue.code).indexOf('ACL_') === 0) {
          needAclCacheRebuild = true;
        }
      } else {
        skipped++;
      }
    } catch (e) {
      skipped++;
    }
  });

  if (needAclCacheRebuild || code === 'ACL_CACHE_MISMATCH') {
    rebuildAllAclCachesFromEffective_();
  }
  if (fixed > 0) {
    bumpCatalogRev_();
  }

  return { ok: true, code: code, fixed: fixed, skipped: skipped };
}

/**
 * @param {'quick'|'deep'} mode
 * @returns {Array<Object>}
 */
function collectCatalogRepairIssues_(mode) {
  var treeRows = readSheetRecords_('Tree');
  var fileRows = readSheetRecords_('Files');
  var aclRows = readSheetRecords_('ACL');
  var userRows = readSheetRecords_('Users');
  var groupRows = readSheetRecords_('Groups');
  var memberRows = readSheetRecords_('GroupMembers');
  var rootId = getVirtualRootFolderId_();
  var engine = buildAclEngineFromRows_(
    treeRows,
    fileRows,
    aclRows,
    memberRows,
    userRows,
    groupRows
  );

  var issues = [];
  collectTreeRepairIssues_(treeRows, rootId, issues);
  collectFileRepairIssues_(fileRows, engine.foldersById, issues);
  collectAclRepairIssues_(aclRows, engine, rootId, issues);
  collectAclCacheRepairIssues_(treeRows, fileRows, engine, issues);
  collectPeopleRepairIssues_(userRows, groupRows, memberRows, issues);

  if (mode === 'deep') {
    collectDriveRepairIssues_(fileRows, issues);
  }
  return issues;
}

/**
 * @param {Object.<string, string>[]} treeRows
 * @param {string} rootId
 * @param {Array<Object>} issues
 */
function collectTreeRepairIssues_(treeRows, rootId, issues) {
  var byId = {};
  var seen = {};
  treeRows.forEach(function (row, idx) {
    var id = String(row.folder_id || '').trim();
    if (!id) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_MISSING_ID',
          'folder',
          '',
          '(строка ' + (idx + 2) + ')',
          'Пустой folder_id',
          'Удалить строку',
          { sheetRow: idx + 2 }
        )
      );
      return;
    }
    if (seen[id]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_DUP_ID',
          'folder',
          id,
          String(row.name || id),
          'Дубликат folder_id',
          'Удалить дублирующую строку (оставить первую)',
          { folderId: id, keepFirst: true }
        )
      );
    } else {
      seen[id] = true;
      byId[id] = row;
    }
    if (!String(row.name || '').trim()) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_EMPTY_NAME',
          'folder',
          id,
          id,
          'Пустое имя папки',
          'Задать «(без имени)»',
          { folderId: id }
        )
      );
    }
  });

  if (!byId[rootId]) {
    issues.push(
      repairIssue_(
        REPAIR_LEVEL_QUICK_,
        'TREE_NO_ROOT',
        'folder',
        rootId,
        rootId,
        'Нет корневой папки CATALOG_VIRTUAL_ROOT_FOLDER_ID',
        'Восстановить вручную / повторный setup',
        { folderId: rootId }
      )
    );
  }

  if (!byId['__TRASH__']) {
    issues.push(
      repairIssue_(
        REPAIR_LEVEL_QUICK_,
        'TREE_NO_TRASH',
        'folder',
        '__TRASH__',
        '## Корзина',
        'Нет системной корзины',
        'Создать строку __TRASH__ под корнем',
        { folderId: '__TRASH__', parentId: rootId }
      )
    );
  } else {
    var trash = byId['__TRASH__'];
    if (String(trash.parent_folder_id || '') !== String(rootId)) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_TRASH_BAD_PARENT',
          'folder',
          '__TRASH__',
          '## Корзина',
          'parent=' + String(trash.parent_folder_id || ''),
          'parent → корень',
          { folderId: '__TRASH__', parentId: rootId }
        )
      );
    }
    if (String(trash.name || '') !== '## Корзина') {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_EMPTY_NAME',
          'folder',
          '__TRASH__',
          String(trash.name || ''),
          'Имя корзины не «## Корзина»',
          'Переименовать в «## Корзина»',
          { folderId: '__TRASH__', name: '## Корзина' }
        )
      );
    }
  }

  Object.keys(byId).forEach(function (id) {
    if (id === rootId) {
      return;
    }
    var row = byId[id];
    var parent = String(row.parent_folder_id || '').trim();
    if (!parent) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_MULTI_ROOT',
          'folder',
          id,
          String(row.name || id),
          'Папка без родителя (не канонический корень)',
          'parent → корень',
          { folderId: id, parentId: rootId }
        )
      );
      return;
    }
    if (!byId[parent]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_BAD_PARENT',
          'folder',
          id,
          String(row.name || id),
          'parent_folder_id не найден: ' + parent,
          'parent → корень',
          { folderId: id, parentId: rootId }
        )
      );
      return;
    }
    if (treeHasParentCycle_(byId, id)) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'TREE_PARENT_CYCLE',
          'folder',
          id,
          String(row.name || id),
          'Цикл в parent_folder_id',
          'parent → корень',
          { folderId: id, parentId: rootId }
        )
      );
    }
  });
}

/**
 * @param {Object.<string, Object>} byId
 * @param {string} folderId
 * @returns {boolean}
 */
function treeHasParentCycle_(byId, folderId) {
  var seen = {};
  var cur = folderId;
  var guard = 0;
  while (cur && guard < 10000) {
    if (seen[cur]) {
      return true;
    }
    seen[cur] = true;
    var row = byId[cur];
    if (!row) {
      return false;
    }
    cur = String(row.parent_folder_id || '').trim();
    guard++;
  }
  return false;
}

/**
 * @param {Object.<string, string>[]} fileRows
 * @param {Object.<string, Object>} foldersById
 * @param {Array<Object>} issues
 */
function collectFileRepairIssues_(fileRows, foldersById, issues) {
  var seen = {};
  fileRows.forEach(function (row, idx) {
    var id = String(row.catalog_id || '').trim();
    var name = String(row.display_name || id);
    if (!id) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_MISSING_ID',
          'file',
          '',
          name || '(строка ' + (idx + 2) + ')',
          'Пустой catalog_id',
          'Удалить строку',
          { sheetRow: idx + 2 }
        )
      );
      return;
    }
    if (seen[id]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_DUP_ID',
          'file',
          id,
          name,
          'Дубликат catalog_id',
          'Удалить дублирующую строку (оставить первую)',
          { catalogId: id, keepFirst: true }
        )
      );
    } else {
      seen[id] = true;
    }

    var folderId = String(row.folder_id || '').trim();
    if (!folderId || !foldersById[folderId]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_BAD_FOLDER',
          'file',
          id,
          name,
          'folder_id не найден: ' + folderId,
          'Переместить в ## Корзина',
          { catalogId: id }
        )
      );
    }
    if (!String(row.display_name || '').trim()) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_EMPTY_NAME',
          'file',
          id,
          id,
          'Пустое display_name',
          'Задать «(без имени)»',
          { catalogId: id }
        )
      );
    }

    var approved = parseBoolean_(row.approved);
    var approvedBy = String(row.approved_by || '').trim();
    if (approved && !approvedBy) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_APPROVED_INCONSISTENT',
          'file',
          id,
          name,
          'approved=true без approved_by',
          'Сбросить approved',
          { catalogId: id, action: 'clear_approved' }
        )
      );
    } else if (!approved && approvedBy) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_APPROVED_INCONSISTENT',
          'file',
          id,
          name,
          'approved_by без approved',
          'Очистить approved_by / approved_at',
          { catalogId: id, action: 'clear_by' }
        )
      );
    }

    var status = String(row.status || '').trim().toLowerCase();
    if ((status === 'pending' || status === 'failed') && !String(row.file_id || '').trim()) {
      // stale pending without drive id — report
      if (status === 'pending') {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_QUICK_,
            'FILE_STATUS_STALE',
            'file',
            id,
            name,
            'status=pending без file_id',
            'Пометить failed',
            { catalogId: id }
          )
        );
      }
    }
    if (status === 'ready' && !String(row.file_id || '').trim()) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'FILE_EMPTY_DRIVE_ID',
          'file',
          id,
          name,
          'status=ready, пустой file_id',
          'Только отчёт (починка в глубокой / вручную)',
          { catalogId: id }
        )
      );
    }
  });
}

/**
 * @param {Object.<string, string>[]} aclRows
 * @param {Object} engine
 * @param {string} rootId
 * @param {Array<Object>} issues
 */
function collectAclRepairIssues_(aclRows, engine, rootId, issues) {
  var users = {};
  Object.keys(engine.userDisplayNameByEmail || {}).forEach(function (email) {
    users[String(email).toLowerCase()] = true;
  });
  readSheetRecords_('Users').forEach(function (u) {
    var e = String(u.email || '')
      .trim()
      .toLowerCase();
    if (e) {
      users[e] = true;
    }
  });
  var groups = {};
  Object.keys(engine.groupNameById || {}).forEach(function (gid) {
    groups[gid] = true;
  });

  var seenExact = {};
  var plusMinus = {};
  var multiLevel = {};

  aclRows.forEach(function (row) {
    var aclId = String(row.acl_id || '').trim();
    var objectType = String(row.object_type || '').trim().toLowerCase();
    var objectId = String(row.object_id || '').trim();
    var principalType = String(row.principal_type || '').trim().toLowerCase();
    var principalId = String(row.principal_id || '').trim();
    var level = normalizePermissionLevel_(row.permission_level);
    var delta = normalizeAclDelta_(row);
    var label = objectType + ':' + objectId;

    if (objectType !== 'folder' && objectType !== 'file') {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_BAD_TYPE',
          'acl',
          aclId,
          label,
          'object_type=' + objectType,
          'Удалить строку',
          { aclId: aclId }
        )
      );
      return;
    }
    if (principalType !== 'user' && principalType !== 'group') {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_BAD_TYPE',
          'acl',
          aclId,
          label,
          'principal_type=' + principalType,
          'Удалить строку',
          { aclId: aclId }
        )
      );
      return;
    }
    if (level === 'none' && delta !== '-') {
      // none without minus is odd but allowed as explicit deny via minus usually
    }
    if (delta !== '+' && delta !== '-' && delta !== 'base' && delta !== '') {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_BAD_TYPE',
          'acl',
          aclId,
          label,
          'delta=' + String(row.delta || ''),
          'Удалить строку',
          { aclId: aclId }
        )
      );
      return;
    }

    var exists =
      objectType === 'folder'
        ? !!engine.foldersById[objectId]
        : !!engine.filesByCatalogId[objectId];
    if (!exists) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_BAD_OBJECT',
          'acl',
          aclId,
          label,
          'object_id отсутствует в Tree/Files',
          'Удалить ACL-строку',
          { aclId: aclId }
        )
      );
    }

    if (principalType === 'user') {
      if (!users[principalId.toLowerCase()]) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_QUICK_,
            'ACL_BAD_PRINCIPAL',
            'acl',
            aclId,
            principalId,
            'user не найден в Users',
            'Удалить ACL-строку',
            { aclId: aclId }
          )
        );
      }
    } else if (!groups[principalId]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_BAD_PRINCIPAL',
          'acl',
          aclId,
          principalId,
          'group не найден в Groups',
          'Удалить ACL-строку',
          { aclId: aclId }
        )
      );
    }

    var isRoot = objectType === 'folder' && objectId === rootId;
    if (isRoot && (delta === '+' || delta === '-')) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_ROOT_DELTA',
          'acl',
          aclId,
          label,
          'у корня delta=' + delta,
          'Нормализовать delta → пусто (base)',
          { aclId: aclId }
        )
      );
    }
    if (!isRoot && (delta === '' || delta === 'base')) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_NONROOT_BASE',
          'acl',
          aclId,
          label,
          'устаревший явный ACL без дельты',
          'Удалить строку (наследовать мать)',
          { aclId: aclId }
        )
      );
    }

    var exactKey =
      objectType +
      '|' +
      objectId +
      '|' +
      principalType +
      '|' +
      principalId.toLowerCase() +
      '|' +
      level +
      '|' +
      delta;
    if (seenExact[exactKey]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_DUP_ROW',
          'acl',
          aclId,
          label,
          'полный дубликат ACL',
          'Удалить дубликат (оставить первую)',
          { aclId: aclId, keepFirst: true, exactKey: exactKey }
        )
      );
    } else {
      seenExact[exactKey] = aclId;
    }

    var pmKey =
      objectType +
      '|' +
      objectId +
      '|' +
      principalType +
      '|' +
      principalId.toLowerCase() +
      '|' +
      level;
    if (!plusMinus[pmKey]) {
      plusMinus[pmKey] = { plus: [], minus: [] };
    }
    if (delta === '+') {
      plusMinus[pmKey].plus.push(aclId);
    } else if (delta === '-') {
      plusMinus[pmKey].minus.push(aclId);
    }

    if (delta === '+' || delta === 'base' || delta === '') {
      var mlKey =
        objectType +
        '|' +
        objectId +
        '|' +
        principalType +
        '|' +
        principalId.toLowerCase();
      if (!multiLevel[mlKey]) {
        multiLevel[mlKey] = {};
      }
      if (!multiLevel[mlKey][level]) {
        multiLevel[mlKey][level] = [];
      }
      multiLevel[mlKey][level].push(aclId);
    }

    // redundant +/- vs parent effective
    if (!isRoot && exists && (delta === '+' || delta === '-')) {
      var parentId = getRepairObjectParentId_(engine, objectType, objectId);
      if (parentId) {
        var parentMap = getEffectiveAclMapFromEngine_(engine, 'folder', parentId);
        var pKey = aclPrincipalMapKey_(principalType, principalId);
        var parentLevel = parentMap[pKey] ? parentMap[pKey].level : 'none';
        if (delta === '+') {
          var rank = { none: 0, reader: 1, commenter: 2, editor: 3 };
          if (rank[parentLevel] >= rank[level]) {
            issues.push(
              repairIssue_(
                REPAIR_LEVEL_QUICK_,
                'ACL_REDUNDANT_PLUS',
                'acl',
                aclId,
                label,
                'у матери уже ' + parentLevel + ' ≥ ' + level,
                'Удалить noop +',
                { aclId: aclId }
              )
            );
          }
        } else if (delta === '-' && (!parentMap[pKey] || parentLevel === 'none')) {
          issues.push(
            repairIssue_(
              REPAIR_LEVEL_QUICK_,
              'ACL_REDUNDANT_MINUS',
              'acl',
              aclId,
              label,
              'principal нет у матери',
              'Удалить noop −',
              { aclId: aclId }
            )
          );
        }
      }
    }
  });

  Object.keys(plusMinus).forEach(function (key) {
    var pack = plusMinus[key];
    if (pack.plus.length && pack.minus.length) {
      pack.plus.forEach(function (aclId) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_QUICK_,
            'ACL_PLUS_MINUS_CONFLICT',
            'acl',
            aclId,
            key,
            '+ и − на одном principal/уровне',
            'Оставить −, удалить +',
            { aclId: aclId, keepMinus: true }
          )
        );
      });
    }
  });

  Object.keys(multiLevel).forEach(function (key) {
    var levels = Object.keys(multiLevel[key]).filter(function (l) {
      return l && l !== 'none';
    });
    if (levels.length <= 1) {
      return;
    }
    var best = levels[0];
    levels.forEach(function (l) {
      best = maxPermissionLevel_(best, l);
    });
    levels.forEach(function (l) {
      if (l === best) {
        return;
      }
      (multiLevel[key][l] || []).forEach(function (aclId) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_QUICK_,
            'ACL_MULTI_LEVEL',
            'acl',
            aclId,
            key,
            'несколько уровней; оставить max=' + best,
            'Удалить более слабый уровень',
            { aclId: aclId }
          )
        );
      });
    });
  });
}

/**
 * @param {Object} engine
 * @param {'folder'|'file'} objectType
 * @param {string} objectId
 * @returns {string}
 */
function getRepairObjectParentId_(engine, objectType, objectId) {
  if (objectType === 'file') {
    var file = engine.filesByCatalogId[objectId];
    return file ? String(file.folder_id || '') : '';
  }
  var folder = engine.foldersById[objectId];
  return folder ? String(folder.parent_folder_id || '') : '';
}

/**
 * @param {Object.<string, string>[]} treeRows
 * @param {Object.<string, string>[]} fileRows
 * @param {Object} engine
 * @param {Array<Object>} issues
 */
function collectAclCacheRepairIssues_(treeRows, fileRows, engine, issues) {
  var memo = {};
  treeRows.forEach(function (row) {
    var id = String(row.folder_id || '').trim();
    if (!id) {
      return;
    }
    var labels = aclRowsToCacheLabels_(
      engine,
      effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, 'folder', id, memo)),
      false
    );
    var expected = {
      e: formatAclCacheField_(labels.editors),
      c: formatAclCacheField_(labels.commenters),
      r: formatAclCacheField_(labels.readers)
    };
    var actual = {
      e: String(row.acl_editors || '').trim(),
      c: String(row.acl_commenters || '').trim(),
      r: String(row.acl_readers || '').trim()
    };
    if (expected.e !== actual.e || expected.c !== actual.c || expected.r !== actual.r) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_CACHE_MISMATCH',
          'folder',
          id,
          String(row.name || id),
          'кэш ≠ эффективные',
          'Пересчитать acl_*',
          { objectType: 'folder', objectId: id }
        )
      );
    }
  });

  fileRows.forEach(function (row) {
    var id = String(row.catalog_id || '').trim();
    if (!id) {
      return;
    }
    var approved = parseBoolean_(row.approved);
    var labels = aclRowsToCacheLabels_(
      engine,
      effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, 'file', id, memo)),
      approved
    );
    var expected = {
      e: formatAclCacheField_(labels.editors),
      c: formatAclCacheField_(labels.commenters),
      r: formatAclCacheField_(labels.readers)
    };
    var actual = {
      e: String(row.acl_editors || '').trim(),
      c: String(row.acl_commenters || '').trim(),
      r: String(row.acl_readers || '').trim()
    };
    if (expected.e !== actual.e || expected.c !== actual.c || expected.r !== actual.r) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'ACL_CACHE_MISMATCH',
          'file',
          id,
          String(row.display_name || id),
          'кэш ≠ эффективные',
          'Пересчитать acl_*',
          { objectType: 'file', objectId: id }
        )
      );
    }
  });
}

/**
 * @param {Object.<string, string>[]} userRows
 * @param {Object.<string, string>[]} groupRows
 * @param {Object.<string, string>[]} memberRows
 * @param {Array<Object>} issues
 */
function collectPeopleRepairIssues_(userRows, groupRows, memberRows, issues) {
  var seenEmail = {};
  var controllers = [];
  userRows.forEach(function (row) {
    var email = String(row.email || '')
      .trim()
      .toLowerCase();
    if (!email) {
      return;
    }
    if (seenEmail[email]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'USER_DUP_EMAIL',
          'user',
          email,
          email,
          'дубликат email',
          'Удалить дубликат (оставить первую)',
          { email: email, keepFirst: true }
        )
      );
    } else {
      seenEmail[email] = true;
    }
    if (String(row.login_role || '').trim().toLowerCase() === 'controller') {
      controllers.push(email);
    }
  });

  var controllerProp =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  controllerProp = String(controllerProp).trim().toLowerCase();

  if (!controllers.length) {
    issues.push(
      repairIssue_(
        REPAIR_LEVEL_QUICK_,
        'USER_NO_CONTROLLER',
        'user',
        controllerProp || '',
        controllerProp || '(нет)',
        'нет login_role=controller',
        'Назначить CONTROLLER_EMAIL / владельца',
        { email: controllerProp }
      )
    );
  } else if (controllers.length > 1) {
    controllers.forEach(function (email) {
      if (controllerProp && email === controllerProp) {
        return;
      }
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'USER_MULTI_CONTROLLER',
          'user',
          email,
          email,
          'лишний controller',
          'Понизить до manager',
          { email: email }
        )
      );
    });
  }

  var seenGroup = {};
  var groupIds = {};
  groupRows.forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    if (!gid) {
      return;
    }
    groupIds[gid] = true;
    if (seenGroup[gid]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'GROUP_DUP_ID',
          'group',
          gid,
          String(row.name || gid),
          'дубликат group_id',
          'Удалить дубликат',
          { groupId: gid, keepFirst: true }
        )
      );
    } else {
      seenGroup[gid] = true;
    }
  });

  memberRows.forEach(function (row) {
    var gid = String(row.group_id || '').trim();
    var email = String(row.email || '')
      .trim()
      .toLowerCase();
    if (!gid || !groupIds[gid]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'GROUP_BAD_MEMBER',
          'group',
          gid || email,
          email,
          'член группы: group_id не найден',
          'Удалить строку GroupMembers',
          { groupId: gid, email: email }
        )
      );
      return;
    }
    if (!email || !seenEmail[email]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_QUICK_,
          'GROUP_BAD_MEMBER',
          'group',
          gid,
          email || '(пусто)',
          'член группы не найден в Users',
          'Удалить строку GroupMembers',
          { groupId: gid, email: email }
        )
      );
    }
  });
}

/**
 * @param {Object.<string, string>[]} fileRows
 * @param {Array<Object>} issues
 */
function collectDriveRepairIssues_(fileRows, issues) {
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  controllerEmail = String(controllerEmail).trim().toLowerCase();
  var catalogRootId = getCatalogRootFolderId_();
  var knownDriveIds = {};
  var driveIdOwners = {};

  fileRows.forEach(function (row) {
    var catalogId = String(row.catalog_id || '').trim();
    var fileId = String(row.file_id || '').trim();
    var name = String(row.display_name || catalogId);
    if (!catalogId || !fileId) {
      return;
    }
    if (driveIdOwners[fileId]) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_DEEP_,
          'DRIVE_DUP_FILE_ID',
          'file',
          catalogId,
          name,
          'тот же file_id у ' + driveIdOwners[fileId],
          'Удалить дублирующую запись Files (оставить первую)',
          { catalogId: catalogId, fileId: fileId, keepFirst: true }
        )
      );
    } else {
      driveIdOwners[fileId] = catalogId;
    }
    knownDriveIds[fileId] = true;

    try {
      var file = DriveApp.getFileById(fileId);
      if (file.isTrashed()) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_DEEP_,
            'DRIVE_TRASHED_ON_DRIVE',
            'file',
            catalogId,
            name,
            'файл в корзине Drive',
            'folder_id → ## Корзина',
            { catalogId: catalogId }
          )
        );
      }
      try {
        var owner = file.getOwner();
        var ownerEmail = owner && owner.getEmail ? String(owner.getEmail() || '').toLowerCase() : '';
        if (controllerEmail && ownerEmail && ownerEmail !== controllerEmail) {
          issues.push(
            repairIssue_(
              REPAIR_LEVEL_DEEP_,
              'DRIVE_NOT_OWNER',
              'file',
              catalogId,
              name,
              'owner=' + ownerEmail,
              'Только отчёт (copy/setOwner — вручную/Jobs)',
              { catalogId: catalogId }
            )
          );
        }
      } catch (eOwner) {
        // ignore owner read failures
      }

      var driveName = String(file.getName() || '');
      if (driveName && name && driveName !== name) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_DEEP_,
            'DRIVE_NAME_MISMATCH',
            'file',
            catalogId,
            name,
            'Drive: «' + driveName + '»',
            'Переименовать на Drive → display_name каталога',
            { catalogId: catalogId, displayName: name, fileId: fileId }
          )
        );
      }

      var driveSize = 0;
      var driveMimeForSize = '';
      try {
        driveMimeForSize = String(file.getMimeType() || '');
        driveSize = resolveDriveFileSizeBytes_(file, driveMimeForSize);
      } catch (eSize) {
        driveSize = 0;
      }
      var catSize = parseNumber_(row.size_bytes) || 0;
      if (driveSize !== catSize) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_DEEP_,
            'DRIVE_SIZE_MISMATCH',
            'file',
            catalogId,
            name,
            'каталог=' +
              catSize +
              ', Drive=' +
              driveSize +
              (catSize <= 1 && isGoogleNativeMimeType_(driveMimeForSize)
                ? ' (возможен stub getSize)'
                : ''),
            'Обновить size_bytes в каталоге с Drive',
            {
              catalogId: catalogId,
              fileId: fileId,
              sizeBytes: driveSize,
              mimeType: driveMimeForSize
            }
          )
        );
      }

      var driveMime = '';
      try {
        driveMime = String(file.getMimeType() || '');
      } catch (eMime) {
        driveMime = '';
      }
      var catMime = String(row.mime_type || '').trim();
      if (driveMime && catMime && driveMime !== catMime) {
        issues.push(
          repairIssue_(
            REPAIR_LEVEL_DEEP_,
            'DRIVE_MIME_MISMATCH',
            'file',
            catalogId,
            name,
            'каталог=' + catMime + ', Drive=' + driveMime,
            'Обновить mime_type в каталоге с Drive',
            { catalogId: catalogId, fileId: fileId, mimeType: driveMime }
          )
        );
      }
    } catch (e) {
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_DEEP_,
          'DRIVE_FILE_MISSING',
          'file',
          catalogId,
          name,
          'нет доступа / файл не найден',
          'status=failed',
          { catalogId: catalogId }
        )
      );
    }
  });

  try {
    var folder = DriveApp.getFolderById(catalogRootId);
    var it = folder.getFiles();
    var guard = 0;
    while (it.hasNext() && guard < 5000) {
      guard++;
      var f = it.next();
      var fid = f.getId();
      if (knownDriveIds[fid] || f.isTrashed()) {
        continue;
      }
      issues.push(
        repairIssue_(
          REPAIR_LEVEL_DEEP_,
          'DRIVE_ORPHAN_IN_FOLDER',
          'file',
          fid,
          f.getName(),
          'есть на Drive, нет в Files',
          'Импорт вручную / игнор',
          { fileId: fid }
        )
      );
    }
  } catch (eRoot) {
    // catalog root missing — skip orphan scan
  }
}

/**
 * @param {Object} issue
 * @returns {boolean} true if changed something
 */
function applyRepairFixForIssue_(issue) {
  var code = issue.code;
  var p = issue.payload || {};

  if (code === 'TREE_MISSING_ID') {
    return deleteSheetRowsByPredicate_('Tree', function (row) {
      return !String(row.folder_id || '').trim();
    });
  }
  if (code === 'TREE_DUP_ID') {
    return deleteDuplicateSheetRows_('Tree', 'folder_id', String(p.folderId || issue.objectId));
  }
  if (
    code === 'TREE_BAD_PARENT' ||
    code === 'TREE_PARENT_CYCLE' ||
    code === 'TREE_MULTI_ROOT' ||
    code === 'TREE_TRASH_BAD_PARENT'
  ) {
    return updateTreeFolderFields_(String(p.folderId || issue.objectId), {
      parent_folder_id: String(p.parentId || getVirtualRootFolderId_())
    });
  }
  if (code === 'TREE_EMPTY_NAME') {
    return updateTreeFolderFields_(String(p.folderId || issue.objectId), {
      name: String(p.name || '(без имени)')
    });
  }
  if (code === 'TREE_NO_TRASH') {
    return ensureTrashFolderRow_(String(p.parentId || getVirtualRootFolderId_()));
  }
  if (code === 'TREE_NO_ROOT') {
    return false;
  }

  if (code === 'FILE_MISSING_ID') {
    return deleteSheetRowsByPredicate_('Files', function (row) {
      return !String(row.catalog_id || '').trim();
    });
  }
  if (code === 'FILE_DUP_ID') {
    return deleteDuplicateSheetRows_('Files', 'catalog_id', String(p.catalogId || issue.objectId));
  }
  if (code === 'FILE_BAD_FOLDER') {
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      folder_id: '__TRASH__'
    });
  }
  if (code === 'FILE_EMPTY_NAME') {
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      display_name: '(без имени)'
    });
  }
  if (code === 'FILE_STATUS_STALE') {
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      status: 'failed',
      last_error: 'repair: stale pending'
    });
  }
  if (code === 'FILE_APPROVED_INCONSISTENT') {
    if (p.action === 'clear_by') {
      return updateFileFields_(String(p.catalogId || issue.objectId), {
        approved_by: '',
        approved_at: ''
      });
    }
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      approved: false,
      approved_by: '',
      approved_at: ''
    });
  }
  if (code === 'FILE_EMPTY_DRIVE_ID') {
    return false;
  }

  if (
    code === 'ACL_PLUS_MINUS_CONFLICT' ||
    code === 'ACL_DUP_ROW' ||
    code === 'ACL_MULTI_LEVEL' ||
    code === 'ACL_BAD_OBJECT' ||
    code === 'ACL_BAD_TYPE' ||
    code === 'ACL_BAD_PRINCIPAL' ||
    code === 'ACL_NONROOT_BASE' ||
    code === 'ACL_REDUNDANT_PLUS' ||
    code === 'ACL_REDUNDANT_MINUS'
  ) {
    return deleteAclRowsByIds_([String(p.aclId || issue.objectId)]);
  }
  if (code === 'ACL_ROOT_DELTA') {
    return updateAclDelta_(String(p.aclId || issue.objectId), '');
  }
  if (code === 'ACL_CACHE_MISMATCH') {
    // handled by rebuild after batch; still mark as fixable no-op per row
    return true;
  }

  if (code === 'USER_DUP_EMAIL') {
    return deleteDuplicateSheetRows_('Users', 'email', String(p.email || issue.objectId), true);
  }
  if (code === 'USER_MULTI_CONTROLLER') {
    return updateUserRole_(String(p.email || issue.objectId), 'manager');
  }
  if (code === 'USER_NO_CONTROLLER') {
    var email = String(p.email || '').trim();
    if (!email) {
      return false;
    }
    return updateUserRole_(email, 'controller');
  }
  if (code === 'GROUP_DUP_ID') {
    return deleteDuplicateSheetRows_('Groups', 'group_id', String(p.groupId || issue.objectId));
  }
  if (code === 'GROUP_BAD_MEMBER') {
    return deleteGroupMemberRow_(String(p.groupId || ''), String(p.email || ''));
  }

  if (code === 'DRIVE_FILE_MISSING') {
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      status: 'failed',
      last_error: 'repair: missing on Drive'
    });
  }
  if (code === 'DRIVE_SIZE_MISMATCH') {
    var sizeBytes = p.sizeBytes;
    if (sizeBytes == null && p.fileId) {
      try {
        var df = DriveApp.getFileById(String(p.fileId));
        var mime =
          String(p.mimeType || '') || getDriveFileMimeType_(df) || '';
        sizeBytes = resolveDriveFileSizeBytes_(df, mime);
      } catch (eSz) {
        return false;
      }
    }
    if (sizeBytes == null) {
      return false;
    }
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      size_bytes: Number(sizeBytes) || 0
    });
  }
  if (code === 'DRIVE_MIME_MISMATCH') {
    var mimeType = p.mimeType;
    if (!mimeType && p.fileId) {
      try {
        mimeType = DriveApp.getFileById(String(p.fileId)).getMimeType();
      } catch (eMm) {
        return false;
      }
    }
    if (!mimeType) {
      return false;
    }
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      mime_type: String(mimeType)
    });
  }
  if (code === 'DRIVE_NAME_MISMATCH') {
    try {
      DriveApp.getFileById(String(p.fileId || '')).setName(String(p.displayName || ''));
      return true;
    } catch (e) {
      return false;
    }
  }
  if (code === 'DRIVE_TRASHED_ON_DRIVE') {
    return updateFileFields_(String(p.catalogId || issue.objectId), {
      folder_id: '__TRASH__'
    });
  }
  if (code === 'DRIVE_DUP_FILE_ID') {
    return deleteDuplicateSheetRowsByFileId_(
      String(p.fileId || ''),
      String(p.catalogId || issue.objectId)
    );
  }
  if (code === 'DRIVE_NOT_OWNER' || code === 'DRIVE_ORPHAN_IN_FOLDER') {
    return false;
  }
  return false;
}

/**
 * @param {string} code
 * @returns {boolean}
 */
function isRepairCodeFixable_(code) {
  var noFix = {
    TREE_NO_ROOT: true,
    FILE_EMPTY_DRIVE_ID: true,
    DRIVE_NOT_OWNER: true,
    DRIVE_ORPHAN_IN_FOLDER: true
  };
  return !noFix[code];
}

/**
 * @param {Array<Object>} issues
 * @returns {Array<{ code: string, count: number, proposedFix: string, fixable: boolean }>}
 */
function summarizeRepairCodes_(issues) {
  var map = {};
  var order = [];
  (issues || []).forEach(function (issue) {
    if (!map[issue.code]) {
      map[issue.code] = {
        code: issue.code,
        count: 0,
        proposedFix: issue.proposedFix || '',
        fixable: isRepairCodeFixable_(issue.code)
      };
      order.push(issue.code);
    }
    map[issue.code].count++;
  });
  return order.map(function (code) {
    return map[code];
  });
}

/**
 * @param {'quick'|'deep'} mode
 * @param {Array<Object>} issues
 * @returns {{ spreadsheetId: string, url: string, title: string }}
 */
function writeCatalogRepairReport_(mode, issues) {
  var tz = Session.getScriptTimeZone() || 'Europe/Moscow';
  var title =
    'Каталог — ремонт (' +
    (mode === 'deep' ? 'глубокая' : 'быстрая') +
    ') ' +
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var ss = SpreadsheetApp.create(title);
  var sheet = ss.getActiveSheet();
  sheet.setName('Отчёт');
  var header = [
    'уровень',
    'severity',
    'код',
    'объект',
    'путь/имя',
    'детали',
    'предложенный_fix',
    'автоfix'
  ];
  var rows = [header];
  (issues || []).forEach(function (issue) {
    rows.push([
      issue.level || '',
      issue.severity || 'error',
      issue.code || '',
      (issue.objectType ? issue.objectType + ':' : '') + String(issue.objectId || ''),
      issue.pathOrName || '',
      issue.details || '',
      issue.proposedFix || '',
      isRepairCodeFixable_(issue.code) ? 'да' : 'нет'
    ]);
  });
  sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  try {
    sheet.autoResizeColumns(1, header.length);
  } catch (e) {
    // ignore
  }
  if (rows.length > 1) {
    try {
      var colors = [];
      for (var r = 0; r < rows.length - 1; r++) {
        var fill = r % 2 === 0 ? '#FFFFFF' : '#DDEBF7';
        var rowColors = [];
        for (var c = 0; c < header.length; c++) {
          rowColors.push(fill);
        }
        colors.push(rowColors);
      }
      sheet.getRange(2, 1, rows.length - 1, header.length).setBackgrounds(colors);
    } catch (e2) {
      // ignore
    }
  }
  return { spreadsheetId: ss.getId(), url: ss.getUrl(), title: title };
}

/**
 * @param {string} level
 * @param {string} code
 * @param {string} objectType
 * @param {string} objectId
 * @param {string} pathOrName
 * @param {string} details
 * @param {string} proposedFix
 * @param {Object=} payload
 * @returns {Object}
 */
function repairIssue_(level, code, objectType, objectId, pathOrName, details, proposedFix, payload) {
  return {
    level: level,
    severity: 'error',
    code: code,
    objectType: objectType,
    objectId: objectId,
    pathOrName: pathOrName,
    details: details,
    proposedFix: proposedFix,
    payload: payload || {}
  };
}

function rebuildAllAclCachesFromEffective_() {
  var engine = createAclEngine_();
  var treeUpdates = [];
  var fileUpdates = [];
  var memo = {};
  Object.keys(engine.foldersById || {}).forEach(function (folderId) {
    var labels = aclRowsToCacheLabels_(
      engine,
      effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, 'folder', folderId, memo)),
      false
    );
    treeUpdates.push({
      folderId: folderId,
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    });
  });
  Object.keys(engine.filesByCatalogId || {}).forEach(function (catalogId) {
    var file = engine.filesByCatalogId[catalogId];
    var labels = aclRowsToCacheLabels_(
      engine,
      effectiveAclMapToRows_(getEffectiveAclMapFromEngine_(engine, 'file', catalogId, memo)),
      parseBoolean_(file && file.approved)
    );
    fileUpdates.push({
      catalogId: catalogId,
      aclEditors: formatAclCacheField_(labels.editors),
      aclCommenters: formatAclCacheField_(labels.commenters),
      aclReaders: formatAclCacheField_(labels.readers)
    });
  });
  writeTreeAclCacheBatch_(treeUpdates);
  writeFilesAclCacheBatch_(fileUpdates);
}

/**
 * @param {string} sheetName
 * @param {function(Object): boolean} pred true = delete
 * @returns {boolean}
 */
function deleteSheetRowsByPredicate_(sheetName, pred) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return false;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return false;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var kept = [headers];
  var any = false;
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    headers.forEach(function (h, hi) {
      obj[h] = values[i][hi];
    });
    if (pred(obj)) {
      any = true;
    } else {
      kept.push(values[i]);
    }
  }
  if (!any) {
    return false;
  }
  sheet.clearContents();
  if (kept.length) {
    sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
  }
  return true;
}

/**
 * Keep first occurrence of id; delete later duplicates.
 *
 * @param {string} sheetName
 * @param {string} idHeader
 * @param {string} idValue
 * @param {boolean=} caseInsensitive
 * @returns {boolean}
 */
function deleteDuplicateSheetRows_(sheetName, idHeader, idValue, caseInsensitive) {
  idValue = String(idValue || '').trim();
  if (!idValue) {
    return false;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return false;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return false;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf(idHeader);
  if (idCol < 0) {
    return false;
  }
  var kept = [headers];
  var seen = false;
  var any = false;
  for (var i = 1; i < values.length; i++) {
    var raw = String(values[i][idCol] || '').trim();
    var cur = caseInsensitive ? raw.toLowerCase() : raw;
    var want = caseInsensitive ? idValue.toLowerCase() : idValue;
    if (cur === want) {
      if (!seen) {
        kept.push(values[i]);
        seen = true;
      } else {
        any = true;
      }
    } else {
      kept.push(values[i]);
    }
  }
  if (!any) {
    return false;
  }
  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, headers.length).setValues(kept);
  return true;
}

/**
 * @param {string} folderId
 * @param {Object.<string, *>} fields
 * @returns {boolean}
 */
function updateTreeFolderFields_(folderId, fields) {
  return updateSheetRowById_('Tree', 'folder_id', folderId, fields);
}

/**
 * @param {string} catalogId
 * @param {Object.<string, *>} fields
 * @returns {boolean}
 */
function updateFileFields_(catalogId, fields) {
  return updateSheetRowById_('Files', 'catalog_id', catalogId, fields);
}

/**
 * @param {string} sheetName
 * @param {string} idHeader
 * @param {string} idValue
 * @param {Object.<string, *>} fields
 * @returns {boolean}
 */
function updateSheetRowById_(sheetName, idHeader, idValue, fields) {
  idValue = String(idValue || '').trim();
  if (!idValue) {
    return false;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return false;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return false;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf(idHeader);
  if (idCol < 0) {
    return false;
  }
  var changed = false;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol] || '').trim() !== idValue) {
      continue;
    }
    Object.keys(fields || {}).forEach(function (key) {
      var col = headers.indexOf(key);
      if (col < 0) {
        return;
      }
      values[i][col] = fields[key];
      changed = true;
    });
    break;
  }
  if (changed) {
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  }
  return changed;
}

/**
 * @param {string[]} aclIds
 * @returns {boolean}
 */
function deleteAclRowsByIds_(aclIds) {
  var set = {};
  (aclIds || []).forEach(function (id) {
    id = String(id || '').trim();
    if (id) {
      set[id] = true;
    }
  });
  if (!Object.keys(set).length) {
    return false;
  }
  return deleteSheetRowsByPredicate_('ACL', function (row) {
    return !!set[String(row.acl_id || '').trim()];
  });
}

/**
 * @param {string} aclId
 * @param {string} delta
 * @returns {boolean}
 */
function updateAclDelta_(aclId, delta) {
  return updateSheetRowById_('ACL', 'acl_id', aclId, { delta: delta });
}

/**
 * @param {string} email
 * @param {string} role
 * @returns {boolean}
 */
function updateUserRole_(email, role) {
  email = String(email || '')
    .trim()
    .toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet || !email) {
    return false;
  }
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var emailCol = headers.indexOf('email');
  var roleCol = headers.indexOf('login_role');
  if (emailCol < 0 || roleCol < 0) {
    return false;
  }
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][emailCol] || '').trim().toLowerCase() !== email) {
      continue;
    }
    sheet.getRange(i + 1, roleCol + 1).setValue(role);
    return true;
  }
  return false;
}

/**
 * @param {string} groupId
 * @param {string} email
 * @returns {boolean}
 */
function deleteGroupMemberRow_(groupId, email) {
  groupId = String(groupId || '').trim();
  email = String(email || '')
    .trim()
    .toLowerCase();
  return deleteSheetRowsByPredicate_('GroupMembers', function (row) {
    var g = String(row.group_id || '').trim();
    var e = String(row.email || '')
      .trim()
      .toLowerCase();
    if (groupId && email) {
      return g === groupId && e === email;
    }
    if (groupId && !email) {
      return g === groupId;
    }
    if (!groupId && email) {
      return e === email;
    }
    return false;
  });
}

/**
 * @param {string} parentId
 * @returns {boolean}
 */
function ensureTrashFolderRow_(parentId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    return false;
  }
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('folder_id');
  if (idCol < 0) {
    return false;
  }
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol] || '') === '__TRASH__') {
      return updateTreeFolderFields_('__TRASH__', {
        parent_folder_id: parentId,
        name: '## Корзина',
        is_system: true
      });
    }
  }
  var row = [];
  headers.forEach(function (h) {
    if (h === 'folder_id') {
      row.push('__TRASH__');
    } else if (h === 'parent_folder_id') {
      row.push(parentId);
    } else if (h === 'name') {
      row.push('## Корзина');
    } else if (h === 'folder_created_at') {
      row.push(new Date());
    } else if (h === 'is_system') {
      row.push(true);
    } else {
      row.push('');
    }
  });
  sheet.appendRow(row);
  return true;
}

/**
 * Keep first catalog row for fileId; delete the duplicate identified by catalogId.
 *
 * @param {string} fileId
 * @param {string} duplicateCatalogId
 * @returns {boolean}
 */
function deleteDuplicateSheetRowsByFileId_(fileId, duplicateCatalogId) {
  fileId = String(fileId || '').trim();
  duplicateCatalogId = String(duplicateCatalogId || '').trim();
  if (!fileId || !duplicateCatalogId) {
    return false;
  }
  var first = '';
  readSheetRecords_('Files').forEach(function (row) {
    if (String(row.file_id || '').trim() !== fileId) {
      return;
    }
    var cid = String(row.catalog_id || '').trim();
    if (!first) {
      first = cid;
    }
  });
  if (!first || first === duplicateCatalogId) {
    // delete this catalog id only if it's not the first
    var ids = {};
    var seen = false;
    readSheetRecords_('Files').forEach(function (row) {
      if (String(row.file_id || '').trim() !== fileId) {
        return;
      }
      var cid = String(row.catalog_id || '').trim();
      if (!seen) {
        seen = true;
        return;
      }
      if (cid === duplicateCatalogId) {
        ids[cid] = true;
      }
    });
    if (!Object.keys(ids).length) {
      ids[duplicateCatalogId] = true;
    }
    return rewriteSheetRemovingRows_('Files', 'catalog_id', ids);
  }
  var remove = {};
  remove[duplicateCatalogId] = true;
  return rewriteSheetRemovingRows_('Files', 'catalog_id', remove);
}
