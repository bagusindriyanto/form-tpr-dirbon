function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return withErrorLogging_('doGet', () =>
    HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Form Update Data TPR Dirbon')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL),
  );
}
