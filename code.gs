// ============================================
// CODE.GS - Main Entry Point
// ============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Write4U - CARS24 CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testConnection() {
  return { success: true, message: 'Backend is working!' };
}