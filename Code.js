function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Curefoods Reimbursement Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
