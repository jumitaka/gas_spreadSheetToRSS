function createRSStoSheet() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  var channel = createXmlChannel(spreadSheet);

  var sheet = spreadSheet.getSheets()[0];

  const MaxRow = 200;
  var lastRow = sheet.getLastRow();
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

let columnData = {};

function getColumnIndexes(sheet)
{
  var headers = sheet.getRange("1:1").getValues()[0];

  columnData['title']
   = getColumnIndex(['Title'], headers);
  columnData['author']
   = getColumnIndex(['Author'], headers);
  columnData['link']
   = getColumnIndex(['Link'], headers);
  columnData['description']
   = getColumnIndex(['Title'], headers);
  columnData['pubDate']
   = getColumnIndex(['Date'], headers);
  columnData['guid']
   = getColumnIndex(['Link'], headers);
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
    throw new Error("該当列名なし:" + columnNames);
  }

  return index;
}

function createXmlItem(row) {
  var title = row[columnData['title']];
  var author = row[columnData['author']];
  var link = row[columnData['link']];
  var desc = row[columnData['description']];
  var date = row[columnData['pubDate']];  
  var guid = row[columnData['guid']];

  var item = XmlService.createElement('item');

  item.addContent(XmlService.createElement('title').setText(XmlService.createCdata(title)));
  item.addContent(XmlService.createElement('author').setText(author));
  item.addContent(XmlService.createElement('link').setText(link));
  item.addContent(XmlService.createElement('description').addContent(XmlService.createCdata(desc)));
  item.addContent(XmlService.createElement('pubDate').setText(convertDate(date)));
  item.addContent(XmlService.createElement('guid').setAttribute('isPermaLink', 'true').setText(guid));

  return item;
}

// RSSのpubDate用に日付を変換（RFC822)
function convertDate(dateString)
{
  dateString = dateString.replace('at ', '').replace('AM', ' AM').replace('PM', ' PM');

  var date = new Date(Date.parse(dateString));
  return date.toUTCString();
}

function createXmlChannel(spreadSheet) {
  var channel = XmlService.createElement('channel');

  channel.addContent(XmlService.createElement('title').setText(spreadSheet.getName()));
  channel.addContent(XmlService.createElement('link').setText(spreadSheet.getUrl()));
  channel.addContent(XmlService.createElement('language').setText('ja'));

  return channel;
}
