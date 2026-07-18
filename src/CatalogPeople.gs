/**
 * §14.10 — список пользователей и групп для модального окна «Пользователи».
 * CRUD UI — следующим шагом; здесь чтение + права редактирования.
 *
 * @returns {{
 *   ok: true,
 *   canEditUsers: boolean,
 *   canEditGroups: boolean,
 *   users: Array<{
 *     email: string,
 *     displayName: string,
 *     loginRole: 'user'|'manager'|'controller',
 *     addedAt: (string|null),
 *     addedBy: string
 *   }>,
 *   groups: Array<{
 *     groupId: string,
 *     name: string,
 *     memberEmails: string[]
 *   }>
 * }}
 */
function getCatalogPeopleAdmin() {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var canEditGroups = canEditCatalogGroups_(userEmail, loginRole);

  var users = readSheetRecords_('Users').map(function (row) {
    var email = String(row.email || '').trim();
    var role = String(row.login_role || 'user').trim().toLowerCase();
    if (role !== 'controller' && role !== 'manager') {
      role = 'user';
    }
    return {
      email: email,
      displayName: resolveUserDisplayName_(row),
      loginRole: role,
      addedAt: formatCatalogDate_(row.added_at),
      addedBy: String(row.added_by || '')
    };
  });
  users.sort(function (a, b) {
    return String(a.displayName || a.email).localeCompare(
      String(b.displayName || b.email),
      'ru',
      { sensitivity: 'base' }
    );
  });

  var membersByGroup = {};
  readSheetRecords_('GroupMembers').forEach(function (row) {
    var groupId = String(row.group_id || '').trim();
    var email = String(row.email || '').trim();
    if (!groupId || !email) {
      return;
    }
    if (!membersByGroup[groupId]) {
      membersByGroup[groupId] = [];
    }
    membersByGroup[groupId].push(email);
  });

  var groups = readSheetRecords_('Groups').map(function (row) {
    var groupId = String(row.group_id || '').trim();
    var memberEmails = (membersByGroup[groupId] || []).slice();
    memberEmails.sort(function (a, b) {
      return a.localeCompare(b, 'ru', { sensitivity: 'base' });
    });
    return {
      groupId: groupId,
      name: String(row.name || '').trim() || groupId,
      memberEmails: memberEmails
    };
  });
  groups.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' });
  });

  return {
    ok: true,
    canEditUsers: true,
    canEditGroups: canEditGroups,
    users: users,
    groups: groups
  };
}

/**
 * @param {string} userEmail
 * @param {'user'|'manager'|'controller'} loginRole
 * @returns {boolean}
 */
function canEditCatalogGroups_(userEmail, loginRole) {
  if (loginRole === 'controller') {
    return true;
  }
  var controllerEmail =
    PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || '';
  return (
    !!userEmail &&
    !!controllerEmail &&
    userEmail.toLowerCase() === controllerEmail.toLowerCase()
  );
}

/**
 * @param {Object.<string, string>} row запись листа Users
 * @returns {string}
 */
function resolveUserDisplayName_(row) {
  var name = String((row && row.display_name) || '').trim();
  if (name) {
    return name;
  }
  return String((row && row.email) || '').trim();
}

/**
 * Пустой display_name → email (schema 0.3 / §14.10).
 */
function backfillEmptyUserDisplayNames_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    return;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var emailCol = headers.indexOf('email');
  var nameCol = headers.indexOf('display_name');
  if (emailCol < 0 || nameCol < 0) {
    return;
  }

  for (var i = 1; i < values.length; i++) {
    var email = String(values[i][emailCol] || '').trim();
    var name = String(values[i][nameCol] || '').trim();
    if (email && !name) {
      sheet.getRange(i + 1, nameCol + 1).setValue(email);
    }
  }
}

/**
 * Добавить / дописать пользователя в Users (header-aligned).
 *
 * @param {{
 *   email: string,
 *   loginRole?: 'user'|'manager'|'controller',
 *   addedBy?: string,
 *   displayName?: string
 * }} input
 */
function appendOrEnsureUserRow_(input) {
  input = input || {};
  var email = String(input.email || '').trim();
  if (!email) {
    return;
  }

  ensureCatalogSchemaUpToDate_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Users');
  }

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  while (headers.length && !headers[headers.length - 1]) {
    headers.pop();
  }

  var emailCol = headers.indexOf('email');
  var nameCol = headers.indexOf('display_name');
  if (emailCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Users sheet headers are invalid.');
  }

  var values = sheet.getDataRange().getValues();
  var normalized = email.toLowerCase();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][emailCol] || '').trim().toLowerCase() === normalized) {
      if (nameCol >= 0) {
        var existingName = String(values[i][nameCol] || '').trim();
        var wantedName = String(input.displayName || '').trim();
        if (!existingName && (wantedName || email)) {
          sheet.getRange(i + 1, nameCol + 1).setValue(wantedName || email);
        }
      }
      return;
    }
  }

  var displayName = String(input.displayName || '').trim() || email;
  var byHeader = {
    email: email,
    login_role: input.loginRole || 'user',
    added_at: new Date(),
    added_by: input.addedBy || '',
    display_name: displayName
  };
  var line = [];
  for (var c = 0; c < headers.length; c++) {
    var key = headers[c];
    line.push(Object.prototype.hasOwnProperty.call(byHeader, key) ? byHeader[key] : '');
  }
  sheet.appendRow(line);
}

/**
 * @param {GoogleAppsScript.Drive.User} user
 * @returns {string}
 */
function resolveDriveUserDisplayName_(user) {
  if (!user) {
    return '';
  }
  try {
    var name = String(user.getName() || '').trim();
    if (name) {
      return name;
    }
  } catch (e) {
    // ignore
  }
  try {
    return String(user.getEmail() || '').trim();
  } catch (e2) {
    return '';
  }
}

/**
 * §14.10 — добавить пользователя.
 *
 * @param {{ email: string, displayName?: string, loginRole?: 'user'|'manager' }} input
 * @returns {{ ok: true }}
 */
function createCatalogUser(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanRunCatalogOperations_(ctx.loginRole);

  input = input || {};
  var email = normalizeEmailStrict_(input.email);
  var loginRole = String(input.loginRole || 'user').trim().toLowerCase();
  if (loginRole !== 'user' && loginRole !== 'manager') {
    throw catalogError_('INVALID_INPUT', 'Роль может быть только Пользователь или Менеджер.');
  }
  if (findUserRowIndexByEmail_(email) >= 0) {
    throw catalogError_('ALREADY_EXISTS', 'Пользователь уже в каталоге: ' + email);
  }

  var displayName = String(input.displayName || '').trim() || email;
  appendOrEnsureUserRow_({
    email: email,
    loginRole: loginRole,
    addedBy: ctx.userEmail,
    displayName: displayName
  });
  return { ok: true };
}

/**
 * §14.10 — изменить имя и/или роль (не email, не роль Управляющего).
 *
 * @param {{ email: string, displayName?: string, loginRole?: 'user'|'manager' }} input
 * @returns {{ ok: true }}
 */
function updateCatalogUser(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanRunCatalogOperations_(ctx.loginRole);

  input = input || {};
  var email = normalizeEmailStrict_(input.email);
  var sheetMeta = getUsersSheetMeta_();
  var rowIndex = findUserRowIndexByEmail_(email);
  if (rowIndex < 0) {
    throw catalogError_('NOT_FOUND', 'Пользователь не найден: ' + email);
  }

  var values = sheetMeta.sheet.getDataRange().getValues();
  var currentRole = String(values[rowIndex][sheetMeta.roleCol] || '')
    .trim()
    .toLowerCase();
  if (currentRole === 'controller') {
    if (Object.prototype.hasOwnProperty.call(input, 'loginRole')) {
      var wanted = String(input.loginRole || '').trim().toLowerCase();
      if (wanted && wanted !== 'controller') {
        throw catalogError_('NOT_ALLOWED', 'Нельзя сменить роль Управляющего.');
      }
    }
  } else if (Object.prototype.hasOwnProperty.call(input, 'loginRole')) {
    var nextRole = String(input.loginRole || 'user').trim().toLowerCase();
    if (nextRole !== 'user' && nextRole !== 'manager') {
      throw catalogError_('INVALID_INPUT', 'Роль может быть только Пользователь или Менеджер.');
    }
    sheetMeta.sheet.getRange(rowIndex + 1, sheetMeta.roleCol + 1).setValue(nextRole);
  }

  if (Object.prototype.hasOwnProperty.call(input, 'displayName') && sheetMeta.nameCol >= 0) {
    var displayName = String(input.displayName || '').trim() || email;
    sheetMeta.sheet.getRange(rowIndex + 1, sheetMeta.nameCol + 1).setValue(displayName);
  }

  return { ok: true };
}

/**
 * §14.10 — удалить пользователя из каталога, групп и ACL.
 *
 * @param {{ email: string }} input
 * @returns {{ ok: true }}
 */
function deleteCatalogUser(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanRunCatalogOperations_(ctx.loginRole);

  var email = normalizeEmailStrict_((input || {}).email);
  var sheetMeta = getUsersSheetMeta_();
  var rowIndex = findUserRowIndexByEmail_(email);
  if (rowIndex < 0) {
    throw catalogError_('NOT_FOUND', 'Пользователь не найден: ' + email);
  }
  var values = sheetMeta.sheet.getDataRange().getValues();
  var role = String(values[rowIndex][sheetMeta.roleCol] || '')
    .trim()
    .toLowerCase();
  if (role === 'controller') {
    throw catalogError_('NOT_ALLOWED', 'Нельзя удалить Управляющего.');
  }
  if (email.toLowerCase() === ctx.userEmail.toLowerCase()) {
    throw catalogError_('NOT_ALLOWED', 'Нельзя удалить собственный аккаунт.');
  }

  sheetMeta.sheet.deleteRow(rowIndex + 1);
  deleteGroupMembershipsForEmail_(email);
  deleteAclForPrincipal_('user', email);
  return { ok: true };
}

/**
 * §14.10 — создать группу (только Управляющий).
 *
 * @param {{ name: string, memberEmails?: string[] }} input
 * @returns {{ ok: true, groupId: string }}
 */
function createCatalogGroup(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanEditCatalogGroupsActor_(ctx);

  input = input || {};
  var name = String(input.name || '').trim();
  if (!name) {
    throw catalogError_('INVALID_INPUT', 'Укажите имя группы.');
  }
  assertGroupNameUnique_(name, null);

  var groupId = Utilities.getUuid();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Groups');
  }
  sheet.appendRow([groupId, name, new Date(), ctx.userEmail]);

  var memberEmails = normalizeMemberEmailsInCatalog_(input.memberEmails || []);
  rewriteGroupMembers_(groupId, memberEmails);

  return { ok: true, groupId: groupId };
}

/**
 * §14.10 — переименовать группу и/или состав.
 *
 * @param {{ groupId: string, name?: string, memberEmails?: string[] }} input
 * @returns {{ ok: true }}
 */
function updateCatalogGroup(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanEditCatalogGroupsActor_(ctx);

  input = input || {};
  var groupId = String(input.groupId || '').trim();
  if (!groupId) {
    throw catalogError_('INVALID_INPUT', 'groupId is required.');
  }

  var groupsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups');
  if (!groupsSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Groups');
  }
  var values = groupsSheet.getDataRange().getValues();
  if (values.length < 2) {
    throw catalogError_('NOT_FOUND', 'Группа не найдена.');
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('group_id');
  var nameCol = headers.indexOf('name');
  if (idCol < 0 || nameCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Groups sheet headers are invalid.');
  }

  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol] || '').trim() === groupId) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex < 0) {
    throw catalogError_('NOT_FOUND', 'Группа не найдена.');
  }

  if (Object.prototype.hasOwnProperty.call(input, 'name')) {
    var name = String(input.name || '').trim();
    if (!name) {
      throw catalogError_('INVALID_INPUT', 'Укажите имя группы.');
    }
    assertGroupNameUnique_(name, groupId);
    groupsSheet.getRange(rowIndex + 1, nameCol + 1).setValue(name);
  }

  if (Object.prototype.hasOwnProperty.call(input, 'memberEmails')) {
    var memberEmails = normalizeMemberEmailsInCatalog_(input.memberEmails || []);
    rewriteGroupMembers_(groupId, memberEmails);
  }

  return { ok: true };
}

/**
 * §14.10 — удалить группу + членства + ACL.
 *
 * @param {{ groupId: string }} input
 * @returns {{ ok: true }}
 */
function deleteCatalogGroup(input) {
  var ctx = requirePeopleAdminActor_();
  assertCanEditCatalogGroupsActor_(ctx);

  var groupId = String((input || {}).groupId || '').trim();
  if (!groupId) {
    throw catalogError_('INVALID_INPUT', 'groupId is required.');
  }

  var groupsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups');
  if (!groupsSheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Groups');
  }
  var values = groupsSheet.getDataRange().getValues();
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var idCol = headers.indexOf('group_id');
  if (idCol < 0) {
    throw catalogError_('SCHEMA_MISMATCH', 'Groups sheet headers are invalid.');
  }

  var found = false;
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idCol] || '').trim() === groupId) {
      groupsSheet.deleteRow(i + 1);
      found = true;
      break;
    }
  }
  if (!found) {
    throw catalogError_('NOT_FOUND', 'Группа не найдена.');
  }

  rewriteGroupMembers_(groupId, []);
  deleteAclForPrincipal_('group', groupId);
  return { ok: true };
}

/**
 * @returns {{ userEmail: string, loginRole: 'user'|'manager'|'controller' }}
 */
function requirePeopleAdminActor_() {
  assertCatalogReady_();
  ensureCatalogSchemaUpToDate_();
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }
  var loginRole = getLoginRoleForUser_(userEmail);
  return { userEmail: userEmail, loginRole: loginRole };
}

/**
 * @param {{ userEmail: string, loginRole: string }} ctx
 */
function assertCanEditCatalogGroupsActor_(ctx) {
  if (!canEditCatalogGroups_(ctx.userEmail, ctx.loginRole)) {
    throw catalogError_('NOT_ALLOWED', 'Изменять группы может только Управляющий.');
  }
}

/**
 * @param {string} email
 * @returns {string}
 */
function normalizeEmailStrict_(email) {
  var value = String(email || '').trim().toLowerCase();
  if (!value || value.indexOf('@') < 1) {
    throw catalogError_('INVALID_INPUT', 'Укажите корректный email.');
  }
  return value;
}

/**
 * @returns {{ sheet: GoogleAppsScript.Spreadsheet.Sheet, emailCol: number, roleCol: number, nameCol: number }}
 */
function getUsersSheetMeta_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Users');
  }
  var headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
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
  return { sheet: sheet, emailCol: emailCol, roleCol: roleCol, nameCol: nameCol };
}

/**
 * @param {string} email
 * @returns {number} 0-based data row index in getValues(), or -1
 */
function findUserRowIndexByEmail_(email) {
  var meta = getUsersSheetMeta_();
  var values = meta.sheet.getDataRange().getValues();
  var normalized = String(email || '')
    .trim()
    .toLowerCase();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][meta.emailCol] || '').trim().toLowerCase() === normalized) {
      return i;
    }
  }
  return -1;
}

/**
 * @param {string} name
 * @param {string|null} exceptGroupId
 */
function assertGroupNameUnique_(name, exceptGroupId) {
  var rows = readSheetRecords_('Groups');
  var needle = String(name || '')
    .trim()
    .toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i].group_id || '').trim();
    if (exceptGroupId && id === exceptGroupId) {
      continue;
    }
    if (String(rows[i].name || '').trim().toLowerCase() === needle) {
      throw catalogError_('ALREADY_EXISTS', 'Группа с таким именем уже есть.');
    }
  }
}

/**
 * @param {string[]} emails
 * @returns {string[]}
 */
function normalizeMemberEmailsInCatalog_(emails) {
  var known = {};
  readSheetRecords_('Users').forEach(function (row) {
    var email = String(row.email || '').trim();
    if (email) {
      known[email.toLowerCase()] = email;
    }
  });
  var out = [];
  var seen = {};
  (emails || []).forEach(function (raw) {
    var key = String(raw || '')
      .trim()
      .toLowerCase();
    if (!key || !known[key] || seen[key]) {
      return;
    }
    seen[key] = true;
    out.push(known[key]);
  });
  return out;
}

/**
 * @param {string} groupId
 * @param {string[]} memberEmails
 */
function rewriteGroupMembers_(groupId, memberEmails) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GroupMembers');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: GroupMembers');
  }
  var values = sheet.getDataRange().getValues();
  if (values.length >= 1) {
    var headers = values[0].map(function (h) {
      return String(h).trim();
    });
    var groupCol = headers.indexOf('group_id');
    if (groupCol < 0) {
      throw catalogError_('SCHEMA_MISMATCH', 'GroupMembers headers are invalid.');
    }
    for (var i = values.length - 1; i >= 1; i--) {
      if (String(values[i][groupCol] || '').trim() === groupId) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  var now = new Date();
  (memberEmails || []).forEach(function (email) {
    sheet.appendRow([groupId, email, now]);
  });
}

/**
 * @param {string} email
 */
function deleteGroupMembershipsForEmail_(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GroupMembers');
  if (!sheet) {
    return;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var emailCol = headers.indexOf('email');
  if (emailCol < 0) {
    return;
  }
  var needle = String(email || '')
    .trim()
    .toLowerCase();
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][emailCol] || '').trim().toLowerCase() === needle) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * @param {'user'|'group'} principalType
 * @param {string} principalId
 */
function deleteAclForPrincipal_(principalType, principalId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACL');
  if (!sheet) {
    return;
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }
  var headers = values[0].map(function (h) {
    return String(h).trim();
  });
  var typeCol = headers.indexOf('principal_type');
  var idCol = headers.indexOf('principal_id');
  if (typeCol < 0 || idCol < 0) {
    return;
  }
  var needleType = String(principalType || '').trim().toLowerCase();
  var needleId = String(principalId || '').trim().toLowerCase();
  for (var i = values.length - 1; i >= 1; i--) {
    var rowType = String(values[i][typeCol] || '')
      .trim()
      .toLowerCase();
    var rowId = String(values[i][idCol] || '')
      .trim()
      .toLowerCase();
    if (rowType === needleType && rowId === needleId) {
      sheet.deleteRow(i + 1);
    }
  }
}
