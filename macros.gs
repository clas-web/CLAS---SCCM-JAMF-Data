function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().setFrozenRows(1);
  spreadsheet.getActiveSheet().setFrozenColumns(1);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow(), 1, 1508, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
};