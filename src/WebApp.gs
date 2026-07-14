/**
 * §15.4 — состояние для UI при загрузке веб-приложения.
 *
 * @returns {{
 *   userEmail: string,
 *   isSpreadsheetOwner: boolean,
 *   catalog: ReturnType<typeof isCatalogInitialized>,
 *   needsSetup: boolean,
 *   setupBlocked: boolean,
 *   virtualRootFolderId: (string|null),
 *   loginRole: (('user'|'manager'|'controller')|null),
 *   canRunCatalogOperations: boolean
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

  var needsSetup = !catalog.initialized && !hasPartialProps;
  var setupBlocked =
    !catalog.initialized &&
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
  var isOwner = isSpreadsheetOwnerEmail_(email);
  var isController =
    !!email &&
    !!controllerEmail &&
    email.toLowerCase() === controllerEmail.toLowerCase();
  var canRunCatalogOperations =
    loginRole === 'manager' || loginRole === 'controller' || isController || isOwner;

  return {
    userEmail: email,
    isSpreadsheetOwner: isSpreadsheetOwnerEmail_(email),
    catalog: catalog,
    needsSetup: needsSetup,
    setupBlocked: setupBlocked,
    virtualRootFolderId: virtualRootFolderId,
    loginRole: loginRole,
    canRunCatalogOperations: canRunCatalogOperations
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
