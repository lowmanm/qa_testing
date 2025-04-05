// ====================
// App Entry Point & UI Bootstrap
// ====================

/**
 * Entry point when web app is accessed directly via URL.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Includes client-side HTML/JS/CSS partials into the main template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Launches the app in a modal dialog (used only if manually invoked from the script editor).
 */
function openQaApp() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Evaluation System');
}
