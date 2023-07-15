function createRSStoSheet() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var channel = createXmlChannel(spreadSheet);

  var sheet = spreadSheet.getSheets()[0];

  var lastRow = sheet.getLastRow();
  
  let MaxRow = 200;
  var firstRow = (function()
  {
    if(lastRow < MaxRow + 1)
    {
      return 2;
    }
    else
    {
      return lastRow - MaxRow;
    }
  })();

  var range = sheet.getRange(firstRow, 1, lastRow, 5);
  var rows = range.getValues();
  rows = rows.filter(function (r) { return r[0]; } );
  rows = rows.reverse();

  getColumnIndexes(sheet);

  for (var i = 0; i < rows.length; i++) {
    var item = createXmlItem(rows[i]);
    channel.addContent(item);
  }

  var root = XmlService.createElement('rss')
    .setAttribute('version', "2.0")
    .addContent(channel);

  var document = XmlService.createDocument(root);
  var xml = XmlService.getPrettyFormat().format(document);

  Logger.log(xml);

  return xml;
}

let columnDate = -1;
let columnTitle = -1;
let columnDesc = -1;
let columnAuthor = -1;
let columnLink = -1;

function getColumnIndexes(sheet)
{
  var headers = sheet.getRange("1:1").getValues()[0];

  columnDate = getColumnIndex(['Date'], headers);
  columnTitle = getColumnIndex(['Title'], headers);
  columnDesc = getColumnIndex(['Title'], headers);
  columnAuthor = getColumnIndex(['Author'], headers);
  columnLink = getColumnIndex(['Link'], headers);
}

function getColumnIndex(columnNames, headers)
{
  var index = headers.findIndex(head =>
  {
    var ret = columnNames.includes(head);
    return ret;
  });

  if(index == -1)
  {
    return index;
  }

  return index;
}

function createXmlItem(row) {
  var date = row[columnDate];
  var title = row[columnTitle];
  var desc = row[columnDesc];
  var author = row[columnAuthor];
  var link = row[columnLink];

  var item = XmlService.createElement('item');

  item.addContent(XmlService.createElement('title').setText(title));
  item.addContent(XmlService.createElement('author').setText(author));
  item.addContent(XmlService.createElement('link').setText(link));
  item.addContent(XmlService.createElement('description').addContent(XmlService.createCdata(desc)));
  item.addContent(XmlService.createElement('pubDate').setText(date));

  return item;
}

function createXmlChannel(spreadSheet) {
  var channel = XmlService.createElement('channel');

  channel.addContent(XmlService.createElement('title').setText(spreadSheet.getName()));
  channel.addContent(XmlService.createElement('link').setText(spreadSheet.getUrl()));
  channel.addContent(XmlService.createElement('language').setText('ja'));

  return channel;
}
