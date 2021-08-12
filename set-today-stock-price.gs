const getUrl = (symbol) => `https://marketdata.set.or.th/mkt/stockchartdata.do?symbol=${symbol}&time=${+new Date()}`

const stocks = ["E1VFVN3001"];

function updatePriceFor(symbol) {
  var response = UrlFetchApp.fetch(getUrl(symbol));
  var data = Utilities.parseCsv(response);
  var dayEnd = data[data.length-1];

  const [d, m, y] = dayEnd[0].split(' ')[0].split('/');
  const date = new Date((+y), (+m)-1, +d);
  const closedPrice = dayEnd[2];

  if (!closedPrice) {
    console.log("No data for date", date);
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(symbol);
  var lastRowNum = sheet.getLastRow();
  var lastDate = sheet.getRange(`A${lastRowNum}`).getValue();

  if (date.getTime() !== lastDate.getTime()) {
    console.log(`Updating data for ${symbol} @ ${date}`);

    var nextCell = sheet.getRange(`A${lastRowNum + 1}:B${lastRowNum + 1}`);
    nextCell.setValues([
      [date, closedPrice]
    ]);
  } else {
    console.log(`Data for ${symbol} @ ${date} already exist.`);
  }
}

function main() {
  for(let symbol of stocks) {
    updatePriceFor(symbol)
  }
}
