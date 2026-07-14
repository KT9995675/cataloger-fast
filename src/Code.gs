/**
 * §15.4 — точка входа веб-приложения.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cataloger Fast')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * @param {string} filename
 * @returns {string}
 */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
