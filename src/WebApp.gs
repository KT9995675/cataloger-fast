/**
 * §15.4 — состояние для UI при загрузке веб-приложения.
 *
 * @returns {{
 *   userEmail: string,
 *   isSpreadsheetOwner: boolean,
 *   catalog: ReturnType<typeof isCatalogInitialized>,
 *   needsSetup: boolean,
 *   needsFirstLaunch: boolean,
 *   setupBlocked: boolean,
 *   virtualRootFolderId: (string|null),
 *   loginRole: (('user'|'manager'|'controller')|null),
 *   canRunCatalogOperations: boolean,
 *   canEmptyTrash: boolean
 * }}
 */
function getWebAppBootstrap() {
  var email = Session.getActiveUser().getEmail() || '';
  try {
    ensureCatalogSchemaUpToDate_();
  } catch (e) {
    // sheets may be absent before first setup
  }
  var catalog = isCatalogInitialized();
  var hasPartialProps = hasPartialCatalogProps_(catalog);
  var isOwner = isSpreadsheetOwnerEmail_(email);

  // §15.5 — владелец может «Первый запуск» даже при хвостах шаблона (не initialized).
  var needsFirstLaunch = !catalog.initialized && isOwner;
  var needsSetup = needsFirstLaunch;
  var setupBlocked =
    !catalog.initialized &&
    !isOwner &&
    (hasPartialProps ||
      (!catalog.sheetsReady && hasPartialProps) ||
      (!!catalog.schemaVersion && catalog.schemaVersion !== SCHEMA_VERSION_));

  var virtualRootFolderId = catalog.initialized
    ? PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_) ||
      null
    : null;

  var loginRole = catalog.initialized ? resolveLoginRole_(email) : null;
  var controllerEmail = catalog.initialized
    ? PropertiesService.getDocumentProperties().getProperty(PROP_CONTROLLER_EMAIL_) || ''
    : '';
  var isController =
    !!email &&
    !!controllerEmail &&
    email.toLowerCase() === controllerEmail.toLowerCase();
  var canRunCatalogOperations =
    loginRole === 'manager' || loginRole === 'controller' || isController || isOwner;
  var canEmptyTrash = loginRole === 'controller' || isController || isOwner;

  return {
    userEmail: email,
    isSpreadsheetOwner: isOwner,
    catalog: catalog,
    needsSetup: needsSetup,
    needsFirstLaunch: needsFirstLaunch,
    setupBlocked: setupBlocked,
    virtualRootFolderId: virtualRootFolderId,
    loginRole: loginRole,
    canRunCatalogOperations: canRunCatalogOperations,
    canEmptyTrash: canEmptyTrash
  };
}

/**
 * @param {string} email
 * @returns {('user'|'manager'|'controller'|null)}
 */
function resolveLoginRole_(email) {
  if (!email) {
    return null;
  }
  try {
    return getLoginRoleForUser_(email);
  } catch (e) {
    return null;
  }
}

/**
 * @param {ReturnType<typeof isCatalogInitialized>} catalog
 * @returns {boolean}
 */
function hasPartialCatalogProps_(catalog) {
  return !!(
    catalog.schemaVersion ||
    catalog.catalogRootFolderId ||
    catalog.catalogVirtualRootFolderId
  );
}
