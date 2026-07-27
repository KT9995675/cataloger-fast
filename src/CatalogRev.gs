/**
 * §2.7 / §0.5a — счётчик версии каталога (DocumentProperties `CATALOG_REV`).
 * Таблица = истина; клиентский снимок только читается.
 * Любая мутация Tree/Files/ACL (и Jobs done) → bump; UI раз в 5 мин сверяет rev.
 */

/**
 * Текущая версия каталога (лёгкий poll для UI).
 *
 * @returns {{ ok: true, catalogRev: number }}
 */
function getCatalogRevision() {
  assertCatalogReadyLight_();
  return {
    ok: true,
    catalogRev: getCatalogRev_()
  };
}

/**
 * @returns {number}
 */
function getCatalogRev_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(PROP_CATALOG_REV_);
  var n = parseInt(raw, 10);
  if (isNaN(n) || n < 0) {
    return 0;
  }
  return n;
}

/**
 * Увеличить счётчик после записи в каталог (структура / ACL / готовность Jobs).
 * @returns {number} новое значение
 */
function bumpCatalogRev_() {
  var props = PropertiesService.getDocumentProperties();
  var lock = LockService.getDocumentLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(3000);
  } catch (eLock) {
    acquired = false;
  }

  try {
    var current = getCatalogRev_();
    var next = current + 1;
    props.setProperty(PROP_CATALOG_REV_, String(next));
    return next;
  } finally {
    if (acquired) {
      try {
        lock.releaseLock();
      } catch (eRelease) {
        // ignore
      }
    }
  }
}
