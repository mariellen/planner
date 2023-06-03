function Sortbyordering() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('F1'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 10, ascending: true});
};