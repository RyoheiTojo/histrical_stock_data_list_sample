function findLastRow(colindex, sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var colValues = spreadsheet.getRange(1, colindex, spreadsheet.getLastRow(), 1).getValues();
  var lastRow = colValues.filter(String).length;
  
  return lastRow;
}

function clearSheet(sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  spreadsheet.clear()
}

function createHeader(spreadsheet) {
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('Symbol');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Market');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Date');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setValue('Open');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setValue('High');
  spreadsheet.getRange('F1').activate();
  spreadsheet.getCurrentCell().setValue('Low');
  spreadsheet.getRange('G1').activate();
  spreadsheet.getCurrentCell().setValue('Close');
  spreadsheet.getRange('H1').activate();
  spreadsheet.getCurrentCell().setValue('Volume');
}

function getSymbolList(rowindex, colindex, sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var rownum = findLastRow(colindex, sheetname)
  return spreadsheet.getRange(rowindex, colindex, rownum - (rowindex - 1), 1).getValues()
}

function getProperty(rowindex, colindex, sheetname) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  return spreadsheet.getRange(rowindex, colindex).getValue()
}

function drawSymbolData(rowindex, maxRowNum, symbol, market, startDate, endDate, targetAttribute, interval, sheetname) {
  Logger.log(symbol)
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  resultSheet.getRange(rowindex, 3).activate();
  resultSheet.getCurrentCell().setFormula("=GoogleFinance(\""+market+":"+symbol+"\", \""+targetAttribute+"\", \""+Utilities.formatDate(startDate, "JST", "YYYY/MM/dd")+"\", \""+Utilities.formatDate(endDate, "JST", "YYYY/MM/dd")+"\", \""+interval+"\")");
  
  var symbolMarketTitles = []
  for(j=0; j<maxRowNum;j++) {
    symbolMarketTitles.push([symbol, market])
  }
  resultSheet.getRange(rowindex,1, maxRowNum, 2).setValues(symbolMarketTitles)
}

function CreateStockStickData() {
  clearSheet("対象銘柄中間結果")
  var symbolList      = getSymbolList(1, 6, "対象銘柄情報")
  var targetAttribute = getProperty(1, 3, "対象銘柄情報")
  var startDate       = getProperty(2, 3, "対象銘柄情報")
  var endDate         = getProperty(3, 3, "対象銘柄情報")
  var interval        = getProperty(4, 3, "対象銘柄情報")

  var maxNumberOfRowsPerSymbol = (endDate - startDate)/(24 * 60 * 60 * 1000) + 1 + 1 // Including title row.

  for(i=0; i<symbolList.length; i++) {
    var splitResult = symbolList[i][0].split(":")
    drawSymbolData(i*maxNumberOfRowsPerSymbol + 1, maxNumberOfRowsPerSymbol, splitResult[1], splitResult[0], startDate, endDate, targetAttribute, interval, "対象銘柄中間結果")
  }
};

function FormatStockStickData() {
  var dataSheetName = "対象銘柄中間結果"
  var inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);

  var outputSheetName = "対象銘柄結果"
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outputSheetName);
  clearSheet(outputSheetName)
  createHeader(outputSheet)
  
  var rownum = findLastRow(1, dataSheetName)
  var inputData = inputSheet.getRange(1, 1, rownum, 8).getValues()

  var outputData = []
  for(var i=0; i<inputData.length; i++) {
    if(inputData[i][2] == "Date" || inputData[i][3] == "") continue

    if(inputData[i][1] == "CURRENCY") {
      outputData.push(Array(inputData[i][0], inputData[i][1], inputData[i][2], 0, 0, 0, inputData[i][6], 0))
    }else {
      outputData.push(inputData[i])
    }
  }
  outputSheet.getRange(2, 1, outputData.length, 8).setValues(outputData)
};
