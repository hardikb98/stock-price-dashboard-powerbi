function updateHistoricalData() {

  var startTime = new Date();
  Logger.log("===== HISTORICAL DATA UPDATE STARTED =====");
  
  var sheetId = "1xvDkmo4i83m0krWpNj6d37BGiaqY8FsUCgemvPUj554";
  var ss = SpreadsheetApp.openById(sheetId);

  var tickerSheet = ss.getSheetByName("Sheet1");
  var outSheet = ss.getSheetByName("HistoricalData") || ss.insertSheet("HistoricalData");
  var tempSheet = ss.getSheetByName("Temp") || ss.insertSheet("Temp");

  var today = new Date();
  Logger.log("Today's date: " + today);

  // =========================
  // READ TICKERS
  // =========================
  Logger.log("Reading tickers...");
  
  var tickerData = tickerSheet.getDataRange().getValues();
  var headers = tickerData[0];
  var symbolIndex = headers.indexOf("Symbol");
  var typeIndex = headers.indexOf("Type");

  var trackers = tickerData.slice(1)
    .filter(r => r[symbolIndex])
    .map(r => ({ symbol: r[symbolIndex], type: r[typeIndex] }));

  Logger.log("Total tickers found: " + trackers.length);

  // =========================
  // READ EXISTING DATA ONCE
  // =========================
  Logger.log("Reading existing historical data...");

  var existing = outSheet.getDataRange().getValues();
  var lastDateBySymbol = {};

  for (var i = 1; i < existing.length; i++) {
    var symbol = existing[i][0];
    var date = new Date(existing[i][2]);

    if (!lastDateBySymbol[symbol] || date > lastDateBySymbol[symbol]) {
      lastDateBySymbol[symbol] = date;
    }
  }

  Logger.log("Existing symbols with data: " + Object.keys(lastDateBySymbol).length);

  var newRows = [];
  var totalRowsAdded = 0;

  // =========================
  // PROCESS EACH TICKER
  // =========================
  trackers.forEach(function(tracker, index) {

    var symbol = tracker.symbol;
    var type = tracker.type;

    Logger.log("----");
    Logger.log("Processing (" + (index+1) + "/" + trackers.length + "): " + symbol);

    var startDate = lastDateBySymbol[symbol] 
      ? new Date(lastDateBySymbol[symbol]) 
      : new Date(2025, 0, 1);

    startDate.setDate(startDate.getDate() + 1);

    if (startDate > today) {
      Logger.log(symbol + " is already up-to-date.");
      return;
    }

    Logger.log("Fetching data from: " + startDate.toDateString());

    var formula = '=GOOGLEFINANCE("' + symbol + '","all",DATE('
      + startDate.getFullYear() + ',' 
      + (startDate.getMonth()+1) + ',' 
      + startDate.getDate() 
      + '),TODAY(),"DAILY")';

    tempSheet.clear();
    tempSheet.getRange(1,1).setFormula(formula);

    SpreadsheetApp.flush();
    Utilities.sleep(2000);

    var tempData = tempSheet.getDataRange().getValues();
    var rowsAddedForSymbol = 0;

    for (var j = 1; j < tempData.length; j++) {
      var row = tempData[j];

      newRows.push([
        symbol,
        type,
        row[0],
        row[1] || "",
        row[2] || "",
        row[3] || "",
        row[4] || "",
        row[4] || "",
        row[5] || ""
      ]);

      rowsAddedForSymbol++;
    }

    Logger.log("Rows added for " + symbol + ": " + rowsAddedForSymbol);
    totalRowsAdded += rowsAddedForSymbol;

    tempSheet.clear();
  });

  // =========================
  // APPEND DATA
  // =========================
  if (newRows.length > 0) {
    Logger.log("Appending " + newRows.length + " new rows...");
    outSheet.getRange(
      outSheet.getLastRow() + 1,
      1,
      newRows.length,
      newRows[0].length
    ).setValues(newRows);
  } else {
    Logger.log("No new rows to append.");
  }

  var endTime = new Date();
  var duration = (endTime - startTime) / 1000;

  Logger.log("Total rows added: " + totalRowsAdded);
  Logger.log("Execution time: " + duration + " seconds");
  Logger.log("===== UPDATE COMPLETE =====");
}
