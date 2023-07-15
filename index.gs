function doGet(e) {
  var xml = createRSStoSheet();

  // xml出力
  return ContentService.createTextOutput(xml)
    .setMimeType(ContentService.MimeType.XML);
}