/**
 * §2.4 / §13.1 — создать виртуальную папку (F7) в текущей локации.
 * На Drive отдельная папка **не** создаётся (§3.1 — дерево только в таблице).
 *
 * @param {{
 *   parentFolderId: string,
 *   name: string
 * }} input
 * @returns {{
 *   ok: true,
 *   folder: {
 *     id: string,
 *     parentFolderId: string,
 *     name: string,
 *     sizeBytes: number,
 *     modifiedAt: string,
 *     isSystem: boolean,
 *     editors: string[],
 *     commenters: string[],
 *     readers: string[]
 *   }
 * }}
 */
function createCatalogFolder(input) {
  assertCatalogReady_();

  input = input || {};
  var parentFolderId = String(input.parentFolderId || '').trim();
  var name = String(input.name || '').trim();

  if (!parentFolderId) {
    throw catalogError_('INVALID_INPUT', 'parentFolderId is required.');
  }
  if (!name) {
    throw catalogError_('INVALID_INPUT', 'Имя папки не может быть пустым.');
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw catalogError_('AUTH_REQUIRED', 'Google account email is required.');
  }

  var loginRole = getLoginRoleForUser_(userEmail);
  assertCanRunCatalogOperations_(loginRole);

  var engine = createAclEngine_();
  if (!engine.foldersById[parentFolderId]) {
    throw catalogError_('FOLDER_NOT_FOUND', 'Parent folder not found: ' + parentFolderId);
  }

  assertEditorOnFolderForMove_(engine, userEmail, loginRole, parentFolderId);

  var folderId = Utilities.getUuid();
  var now = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tree');
  if (!sheet) {
    throw catalogError_('SCHEMA_MISMATCH', 'Sheet missing: Tree');
  }

  sheet.appendRow([folderId, parentFolderId, name, now, false, '', '', '']);
  engine.foldersById[folderId] = {
    folder_id: folderId,
    parent_folder_id: parentFolderId,
    name: name,
    is_system: false
  };
  copyExplicitAclFromParentFolder_(engine, 'folder', folderId, parentFolderId);
  var acl = getEffectiveAclDisplayFromEngine_(engine, 'folder', folderId);
  bumpCatalogRev_();

  return {
    ok: true,
    folder: {
      id: folderId,
      parentFolderId: parentFolderId,
      name: name,
      sizeBytes: 0,
      modifiedAt: formatCatalogDate_(now),
      isSystem: false,
      editors: acl.editors || [],
      commenters: acl.commenters || [],
      readers: acl.readers || []
    }
  };
}
