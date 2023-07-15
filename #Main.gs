function doGet(e) {
  Logger.log(e);

  var xml = createRSStoSheet();

  // xml出力
  return ContentService.createTextOutput(xml)
    .setMimeType(ContentService.MimeType.XML);
}