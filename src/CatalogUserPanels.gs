/**
 * §21 — память левой/правой панели на пользователя (UserProperties).
 */

/** @const {string} */
var USER_PROP_PANEL_STATE_ = 'CATALOG_PANEL_STATE_V1';

/**
 * Прочитать сохранённые локации панелей текущего пользователя.
 *
 * @returns {{
 *   ok: true,
 *   left: Object|null,
 *   right: Object|null,
 *   activePanelIndex: number
 * }}
 */
function getUserPanelState() {
  assertCatalogReadyLight_();
  var raw = PropertiesService.getUserProperties().getProperty(USER_PROP_PANEL_STATE_);
  if (!raw) {
    return {
      ok: true,
      left: null,
      right: null,
      activePanelIndex: 0
    };
  }
  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    return {
      ok: true,
      left: null,
      right: null,
      activePanelIndex: 0
    };
  }
  return {
    ok: true,
    left: normalizeUserPanelLoc_(parsed && parsed.left),
    right: normalizeUserPanelLoc_(parsed && parsed.right),
    activePanelIndex: parsed && Number(parsed.activePanelIndex) === 1 ? 1 : 0
  };
}

/**
 * Сохранить локации панелей текущего пользователя.
 *
 * @param {{
 *   left?: Object,
 *   right?: Object,
 *   activePanelIndex?: number
 * }} input
 * @returns {{ ok: true }}
 */
function saveUserPanelState(input) {
  assertCatalogReadyLight_();
  input = input || {};
  var payload = {
    left: normalizeUserPanelLoc_(input.left),
    right: normalizeUserPanelLoc_(input.right),
    activePanelIndex: Number(input.activePanelIndex) === 1 ? 1 : 0,
    savedAt: new Date().toISOString()
  };
  PropertiesService.getUserProperties().setProperty(
    USER_PROP_PANEL_STATE_,
    JSON.stringify(payload)
  );
  return { ok: true };
}

/**
 * @param {*} loc
 * @returns {Object|null}
 */
function normalizeUserPanelLoc_(loc) {
  if (!loc || typeof loc !== 'object') {
    return null;
  }
  var mode = String(loc.mode || 'catalog').trim() === 'external' ? 'external' : 'catalog';
  if (mode === 'external') {
    var mirrorId = String(loc.mirrorCatalogId || '').trim();
    var driveId = String(loc.driveFolderId || '').trim();
    if (!mirrorId) {
      return null;
    }
    return {
      mode: 'external',
      mirrorCatalogId: mirrorId,
      driveFolderId: driveId || ''
    };
  }
  var folderId = String(loc.folderId || '').trim();
  if (!folderId) {
    return null;
  }
  return {
    mode: 'catalog',
    folderId: folderId
  };
}
