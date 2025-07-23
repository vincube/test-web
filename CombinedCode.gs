/**
 * Unified doGet to serve all three reports. Uses ?view=reports|machines|remote
 * Default is a simple menu.
 */
// Updated: 2025-07-23
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || 'menu';
  if (view === 'reports') {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Отчёты Vendista');
  } else if (view === 'machines') {
    return HtmlService.createTemplateFromFile('Index2').evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Управление автоматами Vendista');
  } else if (view === 'remote') {
    return HtmlService.createTemplateFromFile('Index3').evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Управление товарами и журнал пакетов');
  }
  return HtmlService.createTemplateFromFile('MainMenu').evaluate()
    .setTitle('Vendista Reports - Unified');
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}
