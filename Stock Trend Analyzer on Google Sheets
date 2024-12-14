function start() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("SETUP");
  var lastRow = setupSheet.getLastRow();
  

  for(var y = 6; y <= lastRow; y++)
  {
    var checkbox = setupSheet.getRange(y,8).getValue();
    if(checkbox == true)
    {
      setupSheet.getRange("D"+y+":G"+y).clear();
      var symbol = setupSheet.getRange(y,3).getValue();
      var market = setupSheet.getRange(y,2).getValue();
      getData(symbol, market, y);
      setLink(symbol, market, y);
      setupSheet.getRange(y,8).setValue(false);
    }
  }
}

function checkAll() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("SETUP");
  var lastRow = setupSheet.getLastRow();
  
  for(var y = 6; y <= lastRow; y++)
  {
    setupSheet.getRange(y,8).setValue(true);
  }

}

function uncheckAll() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("SETUP");
  var lastRow = setupSheet.getLastRow();
  
  for(var y = 6; y <= lastRow; y++)
  {
    setupSheet.getRange(y,8).setValue(false);
  }

}

function setLink(symbol, market, row)
{
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("SETUP");
  var richValue = SpreadsheetApp.newRichTextValue()
  .setText(symbol)
  .setLinkUrl("https://www.google.com/finance/quote/"+symbol+":"+market)
  .build();
  setupSheet.getRange(row, 9).setRichTextValue(richValue);
}

function getData(symbol, market, row) {

  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("DATA");
  var setupSheet = ss.getSheetByName("SETUP");
  var percentGain = setupSheet.getRange(2,2).getValue();
  var percentDown = setupSheet.getRange(3,2).getValue();
  var daysBack = setupSheet.getRange(1,2).getValue();
  dataSheet.getRange("C1:D1000").clear();

  var cell = dataSheet.getRange("A1");
  cell.setFormula("=GOOGLEFINANCE(\""+market+":"+symbol+"\",\"price\",TODAY()-"+daysBack+",TODAY())");
  dataSheet.getRange("C1").setValue(symbol);

  var datalastRow = dataSheet.getLastRow();
  

  for(var i = 3; i <= datalastRow; i++)
  {
    var pastRow = i - 1;
    var cell = dataSheet.getRange(i,3);
    cell.setFormula("=B"+i+"-B"+pastRow+"");
  }

  for(var i = 3; i <= datalastRow; i++)
  {
    var pastRow = i - 1;
    var cell = dataSheet.getRange(i,4);
    cell.setFormula("=(C"+i+"/B"+pastRow+")*100");
  }

  var percentGainTotal = 0;
  var percentDownTotal = 0;
  var consecutiveDaysGain = 0;

  for(var i = 3; i <= datalastRow; i++)
  {
    var currentValue = dataSheet.getRange(i,4).getValue();
    if(currentValue > percentGain)
    {
      percentGainTotal++;
    }
    if(currentValue < percentDown * -1)
    {
      percentDownTotal++;
    }
  }

  for(var i = datalastRow; i >= 3; i--)
  {
    var currentValue = dataSheet.getRange(i,4).getValue();
    if(currentValue > 0)
    {
      consecutiveDaysGain++;
    }
    else
    {
      break;
    }
  }

  setupSheet.getRange(row,5).setValue(percentGainTotal);
  setupSheet.getRange(row,4).setValue(percentDownTotal);
  setupSheet.getRange(row,6).setValue(consecutiveDaysGain);
  setupSheet.getRange(row,7).setValue(datalastRow-1);

}
