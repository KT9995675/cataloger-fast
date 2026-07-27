/** @const DocumentProperties keys — §3.11 */
var PROP_SCHEMA_VERSION_ = 'SCHEMA_VERSION';
var PROP_CATALOG_ROOT_FOLDER_ID_ = 'CATALOG_ROOT_FOLDER_ID';
var PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_ = 'CATALOG_VIRTUAL_ROOT_FOLDER_ID';
var PROP_CONTROLLER_EMAIL_ = 'CONTROLLER_EMAIL';
var PROP_SETUP_AT_ = 'SETUP_AT';
/** Счётчик версии каталога для клиентского poll (§2.7) — см. также CatalogRev.gs */
var PROP_CATALOG_REV_ = 'CATALOG_REV';

/**
 * §15.1, §15.3 — инициализирован ли каталог (листы + обязательные DocumentProperties).
 *
 * @returns {{
 *   initialized: boolean,
 *   inconsistent: boolean,
 *   sheetsReady: boolean,
 *   schemaVersion: (string|null),
 *   catalogRootFolderId: (string|null),
 *   catalogVirtualRootFolderId: (string|null),
 *   controllerEmail: (string|null),
 *   setupAt: (string|null)
 * }}
 */
function isCatalogInitialized() {
  var props = PropertiesService.getDocumentProperties();
  var schemaVersion = props.getProperty(PROP_SCHEMA_VERSION_) || null;
  var catalogRootFolderId = props.getProperty(PROP_CATALOG_ROOT_FOLDER_ID_) || null;
  var catalogVirtualRootFolderId =
    props.getProperty(PROP_CATALOG_VIRTUAL_ROOT_FOLDER_ID_) || null;
  var controllerEmail = props.getProperty(PROP_CONTROLLER_EMAIL_) || null;
  var setupAt = props.getProperty(PROP_SETUP_AT_) || null;

  var sheetsReady = areCatalogSheetsPresent_();
  var hasRequiredProps = !!(
    schemaVersion &&
    catalogRootFolderId &&
    catalogVirtualRootFolderId
  );
  var schemaVersionOk = schemaVersion === SCHEMA_VERSION_;
  var initialized = sheetsReady && hasRequiredProps && schemaVersionOk;

  var inconsistent =
    (sheetsReady && !hasRequiredProps) ||
    (!sheetsReady && hasRequiredProps) ||
    (hasRequiredProps && !schemaVersionOk);

  return {
    initialized: initialized,
    inconsistent: inconsistent,
    sheetsReady: sheetsReady,
    schemaVersion: schemaVersion,
    catalogRootFolderId: catalogRootFolderId,
    catalogVirtualRootFolderId: catalogVirtualRootFolderId,
    controllerEmail: controllerEmail,
    setupAt: setupAt
  };
}
